// Data layer. One Store, two back ends:
//  - local mode (no sync configured): localStorage + the baked-in seed, as in v1;
//  - sheet mode: the Google Sheet is the source of truth. State is loaded fresh
//    from the sheet on every page load (and the refresh button); every edit is
//    turned into a small cell-level operation and queued (localStorage) until
//    the sync engine writes it to the sheet.
// The UI only calls Store methods; it never knows which back end is active.

const KEY = "bakery.v1.state";
const SHEET_CACHE_KEY = "bakery.v2.sheetCache";
const PIN_KEY = "bakery.v1.pinHash";
const SESSION_KEY = "bakery.v1.unlocked";
export const DEFAULT_PIN_HASH =
  "a1fb4e703a9ef1fa4936801721ff285a97ac85330856674412e054892afe6972";

import {
  costPerUnit,
  calculateQuote,
  allRecipeCosts,
  recipeCostPerKg,
} from "./calculator.js";
import { parseWorkbook, orderToFields, compositeRecipeKey, newId } from "./sheet-model.js";

function clone(v) {
  return JSON.parse(JSON.stringify(v));
}

const SETTING_KEYS = [
  "fillingGPerKg", "crumbCoatGPerKgPerLayer", "finalFrostingGPerKgPerLayer", "wastage", "labourRate",
  "baseHours", "hoursPerKg", "hoursPerLayer", "hoursPerTopping", "deliveryRatePerKm", "tripMultiplier",
  "minDelivery", "markup", "gst", "rounding",
];
const OVERHEAD_KEYS = ["gas", "electricity", "water", "rent", "equipment"];

export function inventoryComputed(item, rawMaterials) {
  const mat = rawMaterials.find((m) => m.name === item.item);
  const rate = mat ? costPerUnit(mat) : 0;
  const current = Number(item.opening || 0) + Number(item.purchased || 0) - Number(item.used || 0);
  const reorder = Number(item.reorderLevel || 0);
  const status = reorder > 0 && current <= reorder ? "REORDER" : "OK";
  return {
    ...item,
    category: mat ? mat.category : item.category || "",
    unit: mat ? mat.unit : item.unit || "",
    current,
    costPerUnit: rate,
    stockValue: current * rate,
    status,
    estimated: !!(mat && mat.estimated),
  };
}

export class Store {
  constructor(seed, { sync = null } = {}) {
    this.seed = seed;
    this.sync = sync; // SyncEngine or null
    this.base = null; // sheet state as last loaded (for conflict checks)
    this.loadedFromSheet = false;
    this.state = this._load();
    this.listeners = new Set();
  }

  get sheetMode() {
    return !!(this.sync && this.sync.transport);
  }

  _defaults() {
    return {
      rawMaterials: clone(this.seed.rawMaterials),
      recipes: clone(this.seed.recipes),
      inventory: clone(this.seed.inventory),
      settings: clone(this.seed.settings),
      library: clone(this.seed.library),
      orders: clone(this.seed.orders || []),
    };
  }

  _key() {
    return this.sheetMode ? SHEET_CACHE_KEY : KEY;
  }

  _load() {
    try {
      const raw = localStorage.getItem(this._key());
      if (!raw) return this._withKeys(this._defaults());
      const parsed = JSON.parse(raw);
      const base = this._defaults();
      if (this.sheetMode && parsed.base) this.base = parsed.base;
      const st = parsed.state || parsed;
      return this._withKeys({
        ...base,
        ...st,
        settings: { ...base.settings, ...(st.settings || {}) },
      });
    } catch {
      return this._withKeys(this._defaults());
    }
  }

  _withKeys(st) {
    for (const r of st.recipes || []) if (!r.key) r.key = r.lineId || compositeRecipeKey(r);
    return st;
  }

  persist() {
    const payload = this.sheetMode ? { state: this.state, base: this.base } : this.state;
    localStorage.setItem(this._key(), JSON.stringify(payload));
    for (const fn of this.listeners) fn(this.state);
  }

  subscribe(fn) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  /** Replace state with a fresh read of the sheet (page load / refresh button). */
  applySheetGrid(grid) {
    const { state, problems } = parseWorkbook(grid);
    this._withKeys(state);
    for (const o of state.orders) {
      const live = this.quote(o, state);
      o.quoteSnapshot = o.quoteSnapshot ? { ...live, ...o.quoteSnapshot } : live;
    }
    // Orders that only exist on this phone (made in local mode / before sync) are uploaded once.
    const local = this._localOnlyOrders(state);
    this.base = clone(state);
    this.state = state;
    this.loadedFromSheet = true;
    this.problems = problems;
    for (const o of local) {
      this.state.orders.unshift(o);
      this._queueOrder(o, true);
    }
    this.persist();
    return { problems, uploaded: local.length };
  }

  _localOnlyOrders(sheetState) {
    let local = [];
    try {
      const raw = JSON.parse(localStorage.getItem(KEY) || "null");
      local = raw?.orders || [];
    } catch {}
    if (!local.length || localStorage.getItem("bakery.v2.localOrdersUploaded") === "1") return [];
    localStorage.setItem("bakery.v2.localOrdersUploaded", "1");
    const ids = new Set(sheetState.orders.map((o) => o.id));
    return local.filter((o) => !ids.has(o.id));
  }

  _enqueue(op) {
    if (this.sheetMode) this.sync.enqueue(op);
  }

  _baseOf(tab, key, field) {
    if (!this.base) return undefined;
    if (tab === "Raw Materials") return this.base.rawMaterials.find((m) => m.name === key)?.[field];
    if (tab === "Inventory") return this.base.inventory.find((m) => m.item === key)?.[field];
    if (tab === "Recipes") return this.base.recipes.find((m) => m.key === key)?.[field];
    if (tab === "Settings") {
      if (key.startsWith("overheads.")) return this.base.settings.overheads?.[key.split(".")[1]];
      return this.base.settings[key];
    }
    return undefined;
  }

  _set(tab, key, field, value) {
    this._enqueue({ kind: "set", tab, key, field, value, base: this._baseOf(tab, key, field) });
  }

  /** Called by the sync engine after ops reach the sheet: they become the new base. */
  markFlushed(done) {
    if (!this.base) return;
    for (const op of done) {
      if (op.kind !== "set") continue;
      const list = { "Raw Materials": ["rawMaterials", "name"], Inventory: ["inventory", "item"], Recipes: ["recipes", "key"] }[op.tab];
      if (list) {
        const row = this.base[list[0]].find((x) => x[list[1]] === op.key);
        if (row) row[op.field] = op.value;
      } else if (op.tab === "Settings") {
        if (op.key.startsWith("overheads.")) this.base.settings.overheads[op.key.split(".")[1]] = op.value;
        else this.base.settings[op.key] = op.value;
      }
    }
    this.persist();
  }

  resetToSeed() {
    this.state = this._withKeys(this._defaults());
    this.persist();
  }

  importState(payload) {
    if (!payload || typeof payload !== "object") throw new Error("Invalid file");
    const base = this._defaults();
    this.state = this._withKeys({
      ...base,
      ...payload,
      settings: { ...base.settings, ...(payload.settings || {}) },
      rawMaterials: payload.rawMaterials || base.rawMaterials,
      recipes: payload.recipes || base.recipes,
      inventory: payload.inventory || base.inventory,
      orders: payload.orders || [],
      library: payload.library || base.library,
    });
    this.persist();
  }

  exportState() {
    return clone(this.state);
  }

  get data() {
    return {
      rawMaterials: this.state.rawMaterials,
      recipes: this.state.recipes,
      settings: this.state.settings,
    };
  }

  get visibleOrders() {
    return this.state.orders.filter((o) => o.status !== "Deleted");
  }

  recipeCosts() {
    return allRecipeCosts(this.state.recipes, this.state.rawMaterials);
  }

  costOf(name) {
    return recipeCostPerKg(name, this.state.recipes, this.state.rawMaterials);
  }

  quote(input, st = this.state) {
    return calculateQuote(input, { rawMaterials: st.rawMaterials, recipes: st.recipes, settings: st.settings });
  }

  inventoryRows() {
    return this.state.inventory.map((i) => inventoryComputed(i, this.state.rawMaterials));
  }

  lowStock() {
    return this.inventoryRows().filter((i) => i.status === "REORDER");
  }

  updateMaterial(name, patch) {
    const i = this.state.rawMaterials.findIndex((m) => m.name === name);
    if (i < 0) return;
    const cur = this.state.rawMaterials[i];
    for (const f of ["category", "packSize", "unit", "packPrice", "notes", "priceStatus"]) {
      if (f in patch && patch[f] !== cur[f]) this._set("Raw Materials", name, f, patch[f]);
    }
    this.state.rawMaterials[i] = { ...cur, ...patch };
    this.persist();
  }

  addMaterial(mat) {
    if (!mat.name) throw new Error("Item name is required");
    if (this.state.rawMaterials.some((m) => m.name === mat.name)) {
      throw new Error("Item name must be unique");
    }
    this.state.rawMaterials.push(mat);
    const inv = { item: mat.name, opening: 0, purchased: 0, used: 0, reorderLevel: 0 };
    if (!this.state.inventory.some((i) => i.item === mat.name)) this.state.inventory.push(inv);
    this._enqueue({
      kind: "append",
      tab: "Raw Materials",
      record: { category: mat.category, name: mat.name, packSize: mat.packSize, unit: mat.unit, packPrice: mat.packPrice, notes: mat.notes || "", priceStatus: mat.priceStatus || "" },
      inventory: { opening: 0, purchased: 0, used: 0, reorderLevel: 0 },
    });
    this.persist();
  }

  updateInventory(item, patch) {
    const i = this.state.inventory.findIndex((x) => x.item === item);
    if (i < 0) return;
    const cur = this.state.inventory[i];
    for (const f of ["opening", "purchased", "used", "reorderLevel"]) {
      if (f in patch && Number(patch[f]) !== Number(cur[f])) this._set("Inventory", item, f, Number(patch[f]));
    }
    this.state.inventory[i] = { ...cur, ...patch };
    this.persist();
  }

  deductStock(usage) {
    // usage: [{ item, qty }]
    const totals = new Map();
    for (const u of usage) totals.set(u.item, (totals.get(u.item) || 0) + Number(u.qty || 0));
    for (const [item, qty] of totals) {
      const i = this.state.inventory.findIndex((x) => x.item === item);
      if (i < 0) continue;
      const used = Math.round((Number(this.state.inventory[i].used || 0) + qty) * 1000) / 1000;
      this.updateInventory(item, { used });
    }
    this.persist();
  }

  updateRecipeLine(key, patch) {
    const line = this.state.recipes.find((r) => r.key === key);
    if (!line) return;
    for (const f of ["ingredient", "qtyPerKg"]) {
      if (f in patch && patch[f] !== line[f]) this._set("Recipes", key, f, f === "qtyPerKg" ? Number(patch[f]) : patch[f]);
    }
    Object.assign(line, patch);
    if (patch.ingredient) line.unit = this.state.rawMaterials.find((m) => m.name === patch.ingredient)?.unit || line.unit;
    this.persist();
  }

  addRecipeLine({ name, type, ingredient, qtyPerKg }) {
    const unit = this.state.rawMaterials.find((m) => m.name === ingredient)?.unit || "";
    const lineId = newId("R");
    const line = { lineId, key: lineId, name, type, ingredient, qtyPerKg: Number(qtyPerKg) || 0, unit };
    // keep it next to the recipe's other lines in the app (the sheet gets a new row at the end)
    const last = this.state.recipes.map((r) => r.name).lastIndexOf(name);
    this.state.recipes.splice(last >= 0 ? last + 1 : this.state.recipes.length, 0, line);
    this._enqueue({ kind: "append", tab: "Recipes", record: { name, type, ingredient, qtyPerKg: line.qtyPerKg, lineId } });
    this.persist();
    return line;
  }

  _queueOrder(o, isNew) {
    this._enqueue({ kind: "order", id: o.id, fields: orderToFields(o), updatedAt: o.updatedAt, isNew: !!isNew });
  }

  addOrder(order) {
    const id = order.id || newId("ORD");
    const row = {
      id,
      createdAt: order.createdAt || new Date().toISOString(),
      customer: order.customer || "",
      phone: order.phone || "",
      size: order.size,
      layers: Number(order.layers),
      sponge: order.sponge,
      filling: order.filling,
      frosting: order.frosting,
      topping1: order.topping1 || "",
      topping2: order.topping2 || "None",
      topping3: order.topping3 || "None",
      box: order.box || "",
      board: order.board || "",
      candles: !!order.candles,
      carryBag: !!order.carryBag,
      delivery: order.delivery || "Pickup",
      distanceKm: Number(order.distanceKm || 0),
      hoursOverride: order.hoursOverride ?? null,
      deliveryDate: order.deliveryDate || "",
      status: order.status || "Enquiry",
      advancePaid: Number(order.advancePaid || 0),
      notes: order.notes || "",
      quoteSnapshot: order.quoteSnapshot || null,
      updatedAt: new Date().toISOString(),
    };
    if (!row.quoteSnapshot) {
      row.quoteSnapshot = this.quote(row);
    }
    this.state.orders.unshift(row);
    this._queueOrder(row, true);
    this.persist();
    return row;
  }

  updateOrder(id, patch) {
    const i = this.state.orders.findIndex((o) => o.id === id);
    if (i < 0) return;
    const { recalculate, ...rest } = patch;
    const next = { ...this.state.orders[i], ...rest, updatedAt: new Date().toISOString() };
    const quoteFields = [
      "size","layers","sponge","filling","frosting","topping1","topping2","topping3",
      "box","board","candles","carryBag","delivery","distanceKm","hoursOverride",
    ];
    if (!patch.quoteSnapshot && (quoteFields.some((k) => k in patch) || recalculate)) {
      next.quoteSnapshot = this.quote(next);
    }
    this.state.orders[i] = next;
    this._queueOrder(next, false);
    this.persist();
    return next;
  }

  deleteOrder(id) {
    if (this.sheetMode) {
      // soft delete: the row stays in the sheet with Status = Deleted
      this.updateOrder(id, { status: "Deleted" });
      return;
    }
    this.state.orders = this.state.orders.filter((o) => o.id !== id);
    this.persist();
  }

  updateSettings(patch) {
    const cur = this.state.settings;
    for (const k of SETTING_KEYS) {
      if (k in patch && Number(patch[k]) !== Number(cur[k])) this._set("Settings", k, "value", Number(patch[k]));
    }
    if (patch.overheads) {
      for (const k of OVERHEAD_KEYS) {
        if (k in patch.overheads && Number(patch.overheads[k]) !== Number(cur.overheads?.[k])) {
          this._set("Settings", "overheads." + k, "value", Number(patch.overheads[k]));
        }
      }
    }
    this.state.settings = { ...cur, ...patch, overheads: { ...cur.overheads, ...(patch.overheads || {}) } };
    this.persist();
  }

  dashboard(now = new Date()) {
    const startOfWeek = new Date(now);
    const day = (startOfWeek.getDay() + 6) % 7; // Monday = 0
    startOfWeek.setHours(0, 0, 0, 0);
    startOfWeek.setDate(startOfWeek.getDate() - day);
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const inRange = (o, from) => {
      const d = new Date(o.deliveryDate || o.createdAt);
      return d >= from && d <= now;
    };
    const sum = (list, fn) => list.reduce((s, o) => s + fn(o), 0);
    const n = (v) => (isFinite(Number(v)) ? Number(v) : 0);
    const money = (o) => {
      const q = o.quoteSnapshot || {};
      return { revenue: n(q.finalQuote), cost: n(q.totalCost), profit: n(q.profit) };
    };
    const orders = this.visibleOrders.filter((o) => o.status !== "Cancelled");
    const week = orders.filter((o) => inRange(o, startOfWeek));
    const month = orders.filter((o) => inRange(o, startOfMonth));
    const pending = orders.filter((o) => !["Delivered", "Paid", "Cancelled"].includes(o.status));
    const roll = (list) => ({
      count: list.length,
      revenue: sum(list, (o) => money(o).revenue),
      cost: sum(list, (o) => money(o).cost),
      profit: sum(list, (o) => money(o).profit),
    });
    return {
      week: roll(week),
      month: roll(month),
      pending: pending.length,
      lowStock: this.lowStock().length,
      upcoming: pending
        .filter((o) => o.deliveryDate)
        .sort((a, b) => String(a.deliveryDate).localeCompare(String(b.deliveryDate)))
        .slice(0, 6),
    };
  }

  profitByMonth() {
    const map = new Map();
    for (const o of this.visibleOrders) {
      if (o.status === "Cancelled") continue;
      const d = new Date(o.deliveryDate || o.createdAt);
      if (isNaN(d)) continue;
      const key = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
      const q = o.quoteSnapshot || {};
      const cur = map.get(key) || { month: key, orders: 0, revenue: 0, cost: 0, profit: 0 };
      cur.orders += 1;
      cur.revenue += Number(q.finalQuote || 0);
      cur.cost += Number(q.totalCost || 0);
      cur.profit += Number(q.profit || 0);
      map.set(key, cur);
    }
    return [...map.values()].sort((a, b) => b.month.localeCompare(a.month));
  }
}

export async function sha256Hex(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function currentPinHash() {
  return localStorage.getItem(PIN_KEY) || DEFAULT_PIN_HASH;
}

export async function checkPin(pin) {
  const hash = await sha256Hex(String(pin || ""));
  return hash === currentPinHash();
}

export async function setPin(pin) {
  const hash = await sha256Hex(String(pin || ""));
  localStorage.setItem(PIN_KEY, hash);
}

export function isUnlocked() {
  return sessionStorage.getItem(SESSION_KEY) === "1";
}

export function unlockSession() {
  sessionStorage.setItem(SESSION_KEY, "1");
}

export function lockSession() {
  sessionStorage.removeItem(SESSION_KEY);
}

export function ordersToSheetRows(orders) {
  const headers = [
    "ID","Created","Customer","Phone","Size","Layers","Sponge","Filling","Frosting",
    "Topping 1","Topping 2","Topping 3","Box","Board","Candles","Carry bag",
    "Delivery","Distance km","Delivery date","Status","Advance paid (₹)","Notes",
    "Quote (₹)","Cost (₹)","Profit (₹)","Hours","Raw material (₹)","Labour (₹)","Delivery charge (₹)",
  ];
  const rows = orders.map((o) => {
    const q = o.quoteSnapshot || {};
    return [
      o.id, o.createdAt, o.customer, o.phone, o.size, o.layers, o.sponge, o.filling, o.frosting,
      o.topping1, o.topping2, o.topping3, o.box, o.board, o.candles ? "Yes" : "No", o.carryBag ? "Yes" : "No",
      o.delivery, o.distanceKm, o.deliveryDate, o.status, o.advancePaid, o.notes,
      q.finalQuote, q.totalCost, q.profit, q.hoursUsed, q.rawMaterialCost, q.labour, q.delivery,
    ];
  });
  return [headers, ...rows];
}

export function toCsv(rows) {
  return rows
    .map((r) =>
      r
        .map((c) => {
          const s = c == null ? "" : String(c);
          return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
        })
        .join(",")
    )
    .join("\n");
}
