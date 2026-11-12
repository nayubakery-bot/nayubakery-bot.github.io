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
import { parseWorkbook, orderToFields, compositeRecipeKey, newId, SECTIONS, paymentStatusOf, DEFAULT_CATEGORIES, categoryToFields, normalizeCategory } from "./sheet-model.js";

const MENU_KEY = {
  sizes: (x) => x.name, layers: (x) => String(x), recipeMenu: (x) => x.name, toppings: (x) => x.name,
  boxes: (x) => x, boards: (x) => x, addons: (x) => x.name,
};
export const PRICE_KINDS = { sizes: "Size", toppings: "Topping", boxes: "Box", boards: "Board", addons: "Add-on" };

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
      categories: clone(DEFAULT_CATEGORIES),
      customers: [],
      priceList: [],
      rules: [],
      zones: [],
      recipeNotes: [],
      purchases: [],
      capacity: {},
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
    // Categories live in the sheet. First sync: create the Categories tab with the defaults.
    let seedCategories = false;
    if (!state.categoriesFromSheet) {
      state.categories = clone(DEFAULT_CATEGORIES);
      seedCategories = typeof this.sync?.transport?.ensureTabs === "function";
    } else if (!state.categories.length) {
      state.categories = clone(DEFAULT_CATEGORIES); // tab emptied by hand: keep the app usable, write nothing
    }
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
    if (seedCategories) for (const c of state.categories) this._enqueue({ kind: "row", tab: "Categories", key: { name: c.name }, fields: categoryToFields(c) });
    this.persist();
    return { problems, uploaded: local.length, seededCategories: seedCategories };
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

  // ---------- product categories ----------
  get categories() {
    const list = this.state.categories && this.state.categories.length ? this.state.categories : DEFAULT_CATEGORIES;
    return list;
  }
  activeCategories() {
    return this.categories.filter((c) => c.active !== false);
  }
  category(name, st = this.state) {
    const list = st.categories && st.categories.length ? st.categories : DEFAULT_CATEGORIES;
    const nm = String(name || "").trim().toLowerCase();
    return list.find((c) => c.name.toLowerCase() === nm) || null;
  }
  /** Is this order (or category name) priced per unit rather than by the cake calculator? */
  isUnitOrder(o, st = this.state) {
    if (!o) return false;
    if (o.quantity != null && o.quantity !== "") return true;
    const c = this.category(o.category, st);
    return !!(c && c.pricing !== "Calculator");
  }
  unitQuote(input, st = this.state) {
    const c = this.category(input.category, st) || {};
    const qty = Math.max(Number(input.quantity) || 0, 0);
    const unitPrice = input.unitPrice != null && input.unitPrice !== "" ? Number(input.unitPrice) || 0 : Number(c.unitPrice) || 0;
    const unitCost = input.unitCost != null && input.unitCost !== "" ? Number(input.unitCost) || 0 : Number(c.costPerUnit) || 0;
    const subtotal = Math.round(qty * unitPrice * 100) / 100;
    const delivery = Math.max(Number(input.deliveryCharge) || 0, 0);
    const totalCost = Math.round(qty * unitCost * 100) / 100;
    const finalQuote = subtotal + delivery;
    return {
      kind: "unit", category: input.category || c.name || "", unit: c.unit || "", quantity: qty, unitPrice, unitCost,
      subtotal, delivery, finalQuote, totalCost, rawMaterialCost: totalCost, labour: 0, overheads: 0,
      profit: Math.round((subtotal - totalCost) * 100) / 100, hoursUsed: 0, costKnown: unitCost > 0,
    };
  }
  saveCategory(c, originalName) {
    const list = (this.state.categories = this.state.categories && this.state.categories.length ? this.state.categories : clone(DEFAULT_CATEGORIES));
    const rec = normalizeCategory({ ...c, active: c.active === false ? "No" : "Yes" });
    if (!rec.name) throw new Error("Give the category a name");
    const old = originalName ? list.find((x) => x.name === originalName) : null;
    if (list.some((x) => x.name.toLowerCase() === rec.name.toLowerCase() && x !== old)) throw new Error(`“${rec.name}” is already a category`);
    if (rec.pricing === "Per unit" && !(rec.unitPrice > 0)) throw new Error("Enter the price per unit");
    if (old && old.pricing === "Calculator" && rec.pricing !== "Calculator" && list.filter((x) => x.pricing === "Calculator").length === 1) throw new Error("Keep at least one category on the cake calculator");
    if (old) Object.assign(old, rec); else list.push(rec);
    const key = { name: old ? originalName : rec.name };
    const fields = { ...categoryToFields(rec) };
    if (old && originalName !== rec.name) fields.name = rec.name;
    this._enqueue({ kind: "row", tab: "Categories", key, fields });
    this.persist();
    return rec;
  }
  removeCategory(name) {
    const list = this.state.categories || [];
    const c = list.find((x) => x.name === name);
    if (!c) return;
    if (c.pricing === "Calculator" && list.filter((x) => x.pricing === "Calculator").length === 1) throw new Error("Keep at least one category on the cake calculator");
    this.state.categories = list.filter((x) => x !== c);
    this._enqueue({ kind: "row", tab: "Categories", key: { name }, remove: true });
    this.persist();
  }

  quote(input, st = this.state) {
    if (this.isUnitOrder(input, st)) return this.unitQuote(input, st);
    return calculateQuote(input, {
      rawMaterials: st.rawMaterials, recipes: st.recipes, settings: st.settings,
      priceList: st.priceList || [], rules: st.rules || [], zones: st.zones || [],
    });
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
    if (this.isUnitOrder(order)) return this._addUnitOrder(order, id);
    const row = {
      category: order.category || this.categories.find((c) => c.pricing === "Calculator")?.name || "Cake",
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
      priceOverride: order.priceOverride ?? null,
      rules: order.rules || [],
      zone: order.zone || "",
      extraAddons: order.extraAddons || [],
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
    this._rememberCustomer(row);
    this.persist();
    return row;
  }

  _addUnitOrder(order, id) {
    const c = this.category(order.category) || {};
    const qty = Number(order.quantity);
    if (!(qty > 0)) throw new Error("Enter how many");
    const row = {
      id,
      category: order.category || c.name || "",
      quantity: qty,
      unitPrice: order.unitPrice != null && order.unitPrice !== "" ? Number(order.unitPrice) : Number(c.unitPrice) || 0,
      createdAt: order.createdAt || new Date().toISOString(),
      customer: order.customer || "",
      phone: order.phone || "",
      size: "", layers: "", sponge: "", filling: "", frosting: "", topping1: "None", topping2: "None", topping3: "None",
      box: "", board: "", candles: false, carryBag: false,
      delivery: order.delivery || "Pickup",
      deliveryCharge: Number(order.deliveryCharge || 0),
      distanceKm: 0, hoursOverride: null, priceOverride: null, rules: [], zone: "", extraAddons: [],
      deliveryDate: order.deliveryDate || "",
      status: order.status || "Enquiry",
      advancePaid: Number(order.advancePaid || 0),
      notes: order.notes || "",
      updatedAt: new Date().toISOString(),
    };
    row.quoteSnapshot = this.unitQuote(row);
    this.state.orders.unshift(row);
    this._queueOrder(row, true);
    this._rememberCustomer(row);
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
      "priceOverride","rules","zone","extraAddons","quantity","unitPrice","deliveryCharge","category",
    ];
    if (!patch.quoteSnapshot && (quoteFields.some((k) => k in patch) || recalculate)) {
      next.quoteSnapshot = this.quote(next);
    }
    this.state.orders[i] = next;
    this._queueOrder(next, false);
    if ("customer" in patch || "phone" in patch) this._rememberCustomer(next);
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

  // ================= v3: everything editable from the app =================

  _capKey(section, item) {
    return section === "recipeMenu" ? "recipeMenu:" + item.type : section;
  }
  _checkRoom(section, item) {
    if (!this.sheetMode) return;
    const cap = this.state.capacity || {};
    const c = cap[this._capKey(section, item)] || (section === "recipeMenu" ? cap.recipeMenu : null);
    if (c && c.found && c.free <= 0) {
      const label = SECTIONS[section].label;
      const err = new Error(`The ${label} list in her sheet has no empty row left (${c.range}). Tap “Make room” to add 5 empty rows inside that block, or clear an old row in the sheet.`);
      err.full = true;
      throw err;
    }
    if (c) c.free -= 1;
  }

  /** Sheet list full: insert `count` empty rows inside it (Sheets widens the Calculator's ranges). */
  makeRoom(section, type, count = 5) {
    if (!this.sheetMode) return;
    this._enqueue({ kind: "makeRoom", section, type: type || "", count });
    const cap = this.state.capacity?.[section === "recipeMenu" && type ? "recipeMenu:" + type : section];
    if (cap) cap.free += count;
    this.persist();
  }
  roomLeft(section, type) {
    if (!this.sheetMode) return null;
    const cap = this.state.capacity?.[section === "recipeMenu" && type ? "recipeMenu:" + type : section];
    return cap && cap.found ? cap : null;
  }

  /** Menu lists in Settings: sizes, layers, recipeMenu, toppings, boxes, boards, addons. */
  addMenuItem(section, item) {
    const s = this.state.settings;
    const list = s[section];
    const key = MENU_KEY[section](section === "layers" ? Number(item.value) : section === "boxes" || section === "boards" ? item.name : item);
    if (!key || key === "0") throw new Error("Give it a name first");
    if (list.some((x) => MENU_KEY[section](x) === key)) throw new Error(`“${key}” is already in the list`);
    if (section === "recipeMenu" && !["Sponge", "Filling", "Frosting"].includes(item.type)) throw new Error("Pick Sponge, Filling or Frosting");
    this._checkRoom(section, item);
    // packaging / topping / add-on items are priced from a Raw Materials row
    if (section === "boxes" || section === "boards") this._ensureMaterial(item.name, item.newMaterial, "Packaging");
    if (section === "toppings" && item.material) this._ensureMaterial(item.material, item.newMaterial, "Toppings");
    if (section === "addons" && item.material) this._ensureMaterial(item.material, item.newMaterial, "Packaging");
    let rec, fields;
    if (section === "sizes") { rec = { name: item.name, kg: Number(item.kg) || 0, servings: String(item.servings || "") }; fields = { kg: rec.kg, servings: rec.servings }; }
    else if (section === "layers") { rec = Number(item.value); fields = {}; }
    else if (section === "recipeMenu") { rec = { type: item.type, name: item.name }; fields = { type: item.type }; }
    else if (section === "toppings") {
      const unit = this.state.rawMaterials.find((m) => m.name === item.material)?.unit || "";
      rec = { name: item.name, material: item.material || "", unit, fixedQty: Number(item.fixedQty) || 0, qtyPerKg: Number(item.qtyPerKg) || 0 };
      fields = { material: rec.material, unit, fixedQty: rec.fixedQty, qtyPerKg: rec.qtyPerKg };
    } else if (section === "boxes" || section === "boards") { rec = item.name; fields = {}; }
    else if (section === "addons") { rec = { name: item.name, material: item.material || item.name, qty: Number(item.qty) || 1 }; fields = { material: rec.material, qty: rec.qty }; }
    if (section === "layers") { list.push(rec); list.sort((a, b) => a - b); }
    else if (section === "recipeMenu") {
      const last = list.map((m) => m.type).lastIndexOf(item.type);
      list.splice(last >= 0 ? last + 1 : list.length, 0, rec);
    } else list.push(rec);
    this._enqueue({ kind: "menu", section, action: "add", key: section === "layers" ? rec : key, fields });
    if (item.sellingPrice != null && item.sellingPrice !== "" && PRICE_KINDS[section]) this.setSellingPrice(PRICE_KINDS[section], key, item.sellingPrice);
    this.persist();
    return rec;
  }

  updateMenuItem(section, key, patch) {
    const list = this.state.settings[section];
    const i = list.findIndex((x) => MENU_KEY[section](x) === String(key));
    if (i < 0) return;
    const cur = list[i];
    const newName = patch.name != null ? String(patch.name).trim() : null;
    if (newName !== null && newName !== String(key)) {
      if (!newName) throw new Error("Name can't be empty");
      if (list.some((x) => MENU_KEY[section](x) === newName)) throw new Error(`“${newName}” is already in the list`);
    }
    const fields = {};
    let next = cur;
    if (section === "sizes") {
      next = { ...cur };
      if ("kg" in patch) fields.kg = next.kg = Number(patch.kg) || 0;
      if ("servings" in patch) fields.servings = next.servings = String(patch.servings);
    } else if (section === "toppings") {
      next = { ...cur };
      for (const f of ["fixedQty", "qtyPerKg"]) if (f in patch) fields[f] = next[f] = Number(patch[f]) || 0;
      if ("material" in patch) {
        this._ensureMaterial(patch.material, patch.newMaterial, "Toppings");
        fields.material = next.material = patch.material;
        next.unit = this.state.rawMaterials.find((m) => m.name === patch.material)?.unit || "";
        fields.unit = next.unit;
      }
    } else if (section === "addons") {
      next = { ...cur };
      if ("qty" in patch) fields.qty = next.qty = Number(patch.qty) || 0;
      if ("material" in patch) { this._ensureMaterial(patch.material, patch.newMaterial, "Packaging"); fields.material = next.material = patch.material; }
    } else if (section === "recipeMenu") {
      next = { ...cur };
    }
    if (newName !== null && newName !== String(key)) {
      const keyField = section === "layers" ? "value" : "name";
      fields[keyField] = section === "layers" ? Number(newName) : newName;
      if (section === "layers") next = Number(newName);
      else if (section === "boxes" || section === "boards") next = newName;
      else next = { ...next, name: newName };
      this._renameEverywhere(section, String(key), newName);
    }
    list[i] = next;
    if (Object.keys(fields).length) this._enqueue({ kind: "menu", section, action: "update", key: section === "layers" ? Number(key) : String(key), fields });
    this.persist();
    return next;
  }

  removeMenuItem(section, key) {
    const list = this.state.settings[section];
    const i = list.findIndex((x) => MENU_KEY[section](x) === String(key));
    if (i < 0) return;
    const [gone] = list.splice(i, 1);
    const cap = this.state.capacity?.[this._capKey(section, gone)];
    if (cap) cap.free += 1;
    this._enqueue({ kind: "menu", section, action: "remove", key: section === "layers" ? Number(key) : String(key) });
    if (PRICE_KINDS[section]) {
      const p = (this.state.priceList || []).find((x) => x.kind === PRICE_KINDS[section] && x.item === String(key));
      if (p) this.setSellingPrice(PRICE_KINDS[section], String(key), null);
    }
    this.persist();
  }

  /** Rename follow-through: recipe lines, price list, recipe notes, packaging material. */
  _renameEverywhere(section, oldName, newName) {
    if (section === "recipeMenu") {
      for (const line of this.state.recipes) if (line.name === oldName) {
        this._set("Recipes", line.key, "name", newName);
        line.name = newName;
      }
      const note = (this.state.recipeNotes || []).find((n) => n.recipe === oldName);
      if (note) {
        this._enqueue({ kind: "row", tab: "Recipe Notes", key: { recipe: oldName }, fields: { recipe: newName } });
        note.recipe = newName;
      }
    }
    if (section === "boxes" || section === "boards") {
      if (this.state.rawMaterials.some((m) => m.name === oldName) && !this.state.rawMaterials.some((m) => m.name === newName)) this.renameMaterial(oldName, newName);
    }
    const kind = PRICE_KINDS[section];
    const p = kind && (this.state.priceList || []).find((x) => x.kind === kind && x.item === oldName);
    if (p) {
      this._enqueue({ kind: "row", tab: "Price List", key: { kind, item: oldName }, fields: { item: newName, updatedAt: new Date().toISOString() } });
      p.item = newName;
    }
  }

  _ensureMaterial(name, spec, category) {
    if (!name || this.state.rawMaterials.some((m) => m.name === name)) return;
    const m = spec || {};
    this.addMaterial({
      category: m.category || category, name, packSize: Number(m.packSize) || 1, unit: m.unit || "nos",
      packPrice: Number(m.packPrice) || 0, notes: "", priceStatus: "", estimated: false,
    });
  }

  /** Rename a raw material and every place that points at it by name. */
  renameMaterial(oldName, newName) {
    const m = this.state.rawMaterials.find((x) => x.name === oldName);
    if (!m || !newName || oldName === newName) return;
    if (this.state.rawMaterials.some((x) => x.name === newName)) throw new Error("Item name must be unique");
    this._set("Raw Materials", oldName, "name", newName);
    m.name = newName;
    const inv = this.state.inventory.find((i) => i.item === oldName);
    if (inv) inv.item = newName; // Inventory mirrors the name by formula
    for (const line of this.state.recipes) if (line.ingredient === oldName) {
      this._set("Recipes", line.key, "ingredient", newName);
      line.ingredient = newName;
    }
    for (const t of this.state.settings.toppings) if (t.material === oldName) this.updateMenuItem("toppings", t.name, { material: newName });
    for (const a of this.state.settings.addons) if (a.material === oldName) this.updateMenuItem("addons", a.name, { material: newName });
    this.persist();
  }

  // ---------- selling prices ----------
  sellingPrice(kind, item) {
    const p = (this.state.priceList || []).find((x) => x.kind === kind && x.item === item);
    return p && Number(p.price) > 0 ? Number(p.price) : null;
  }
  setSellingPrice(kind, item, price) {
    const list = (this.state.priceList = this.state.priceList || []);
    const i = list.findIndex((x) => x.kind === kind && x.item === item);
    const val = price === "" || price == null || !(Number(price) > 0) ? null : Math.round(Number(price) * 100) / 100;
    if (val == null) {
      if (i >= 0) list.splice(i, 1);
      this._enqueue({ kind: "row", tab: "Price List", key: { kind, item }, remove: true });
    } else {
      const rec = { kind, item, price: val, updatedAt: new Date().toISOString() };
      if (i >= 0) list[i] = rec; else list.push(rec);
      this._enqueue({ kind: "row", tab: "Price List", key: { kind, item }, fields: { price: val, updatedAt: rec.updatedAt } });
    }
    this.persist();
  }

  // ---------- "price went up" ----------
  /** What changes if this item's pack price (and optionally pack size) changes. Nothing is saved. */
  priceImpact(name, packPrice, packSize) {
    const m = this.state.rawMaterials.find((x) => x.name === name);
    if (!m) return null;
    const after = { ...this.state, rawMaterials: this.state.rawMaterials.map((x) => (x.name === name ? { ...x, packPrice: Number(packPrice), packSize: packSize ? Number(packSize) : x.packSize } : x)) };
    const before = this.recipeCosts();
    const afterCosts = allRecipeCosts(after.recipes, after.rawMaterials);
    const recipes = Object.keys(before)
      .filter((r) => this.state.recipes.some((l) => l.name === r && l.ingredient === name))
      .map((r) => ({ name: r, before: before[r], after: afterCosts[r] }));
    const uses = [];
    for (const t of this.state.settings.toppings) if (t.material === name) uses.push("Topping: " + t.name);
    for (const b of [...this.state.settings.boxes, ...this.state.settings.boards]) if (b === name) uses.push(b);
    for (const a of this.state.settings.addons) if (a.material === name) uses.push("Add-on: " + a.name);
    const sample = this.seed.sampleOrder;
    const sampleQuote = sample ? { before: this.quote(sample).finalQuote, after: this.quote(sample, after).finalQuote } : null;
    const openOrders = this.visibleOrders.filter((o) => !["Delivered", "Paid", "Cancelled"].includes(o.status) && (this.quote(o).finalQuote !== this.quote(o, after).finalQuote)).length;
    return { item: m, recipes, uses, sampleQuote, openOrders, costPerUnitBefore: costPerUnit(m), costPerUnitAfter: costPerUnit({ ...m, packPrice: Number(packPrice), packSize: packSize ? Number(packSize) : m.packSize }) };
  }
  priceWentUp(name, packPrice, packSize) {
    const patch = { packPrice: Number(packPrice) };
    if (packSize) patch.packSize = Number(packSize);
    const m = this.state.rawMaterials.find((x) => x.name === name);
    if (m && /estimated|sample/i.test(m.priceStatus || "")) patch.priceStatus = "Confirmed " + new Date().toISOString().slice(0, 10);
    if (m && m.estimated) patch.estimated = false;
    this.updateMaterial(name, patch);
  }

  // ---------- recipes ----------
  createRecipe({ name, type, lines = [], method = "" }) {
    name = String(name || "").trim();
    if (!name) throw new Error("Give the recipe a name");
    if (this.state.recipes.some((r) => r.name === name) || this.state.settings.recipeMenu.some((m) => m.name === name)) throw new Error("A recipe with that name already exists");
    if (!["Sponge", "Filling", "Frosting"].includes(type)) throw new Error("Pick Sponge, Filling or Frosting");
    const good = lines.filter((l) => l.ingredient && Number(l.qtyPerKg) > 0);
    if (!good.length) throw new Error("Add at least one ingredient with its quantity per 1 kg");
    this.addMenuItem("recipeMenu", { type, name }); // throws if the menu list is full
    for (const l of good) this.addRecipeLine({ name, type, ingredient: l.ingredient, qtyPerKg: l.qtyPerKg });
    if (method) this.setRecipeNotes(name, method, type);
    this.persist();
  }
  removeRecipeLine(key) {
    const i = this.state.recipes.findIndex((r) => r.key === key);
    if (i < 0) return;
    this.state.recipes.splice(i, 1);
    this._enqueue({ kind: "clearRecipeLine", key });
    this.persist();
  }
  recipeNotes(name) {
    return (this.state.recipeNotes || []).find((n) => n.recipe === name)?.method || "";
  }
  setRecipeNotes(name, method, type) {
    const list = (this.state.recipeNotes = this.state.recipeNotes || []);
    const rec = list.find((n) => n.recipe === name);
    const updatedAt = new Date().toISOString();
    if (rec) Object.assign(rec, { method, updatedAt });
    else list.push({ recipe: name, type: type || "", method, updatedAt });
    this._enqueue({ kind: "row", tab: "Recipe Notes", key: { recipe: name }, fields: { type: type || rec?.type || "", method, updatedAt } });
    this.persist();
  }

  // ---------- inventory ----------
  logPurchase({ item, packs, pricePerPack, notes = "", date }) {
    const m = this.state.rawMaterials.find((x) => x.name === item);
    if (!m) throw new Error("Pick an item");
    const n = Number(packs);
    if (!(n > 0)) throw new Error("How many packs did she buy?");
    const price = pricePerPack === "" || pricePerPack == null ? Number(m.packPrice) : Number(pricePerPack);
    const qty = Math.round(n * Number(m.packSize) * 1000) / 1000;
    let inv = this.state.inventory.find((i) => i.item === item);
    if (!inv) { inv = { item, opening: 0, purchased: 0, used: 0, reorderLevel: 0 }; this.state.inventory.push(inv); }
    this.updateInventory(item, { purchased: Math.round((Number(inv.purchased || 0) + qty) * 1000) / 1000 });
    if (isFinite(price) && price > 0 && price !== Number(m.packPrice)) this.priceWentUp(item, price);
    const rec = {
      id: newId("P"), date: date || new Date().toISOString().slice(0, 10), item, packs: n, qty, unit: m.unit,
      pricePerPack: price, total: Math.round(price * n * 100) / 100, notes,
    };
    (this.state.purchases = this.state.purchases || []).push(rec);
    const { id, ...fields } = rec;
    this._enqueue({ kind: "row", tab: "Purchases", key: { id }, fields });
    this.persist();
    return rec;
  }

  // ---------- pricing rules ----------
  saveRule(rule) {
    const list = (this.state.rules = this.state.rules || []);
    const name = String(rule.name || "").trim();
    if (!name) throw new Error("Give the rule a name");
    const rec = {
      id: rule.id || newId("RULE"), name, effect: rule.effect === "Discount" ? "Discount" : "Charge",
      type: rule.type === "₹" ? "₹" : "%", amount: Number(rule.amount) || 0, active: !!rule.active,
      applyTo: rule.applyTo === "When ticked" ? "When ticked" : "Every order", notes: rule.notes || "",
    };
    if (list.some((r) => r.name === name && r.id !== rec.id)) throw new Error("A rule with that name already exists");
    const i = list.findIndex((r) => r.id === rec.id);
    if (i >= 0) list[i] = rec; else list.push(rec);
    const { id, ...fields } = rec;
    this._enqueue({ kind: "row", tab: "Pricing Rules", key: { id }, fields: { ...fields, active: rec.active ? "Yes" : "No" } });
    this.persist();
    return rec;
  }
  removeRule(id) {
    this.state.rules = (this.state.rules || []).filter((r) => r.id !== id);
    this._enqueue({ kind: "row", tab: "Pricing Rules", key: { id }, remove: true });
    this.persist();
  }
  /** Rules ticked by default on a new order. */
  defaultRules() {
    return (this.state.rules || []).filter((r) => r.active && r.applyTo === "Every order").map((r) => r.name);
  }

  // ---------- delivery zones ----------
  saveZone(z) {
    const list = (this.state.zones = this.state.zones || []);
    const name = String(z.name || "").trim();
    if (!name) throw new Error("Give the zone a name");
    const rec = { id: z.id || newId("Z"), name, charge: Number(z.charge) || 0, notes: z.notes || "" };
    if (list.some((x) => x.name === name && x.id !== rec.id)) throw new Error("A zone with that name already exists");
    const i = list.findIndex((x) => x.id === rec.id);
    if (i >= 0) list[i] = rec; else list.push(rec);
    const { id, ...fields } = rec;
    this._enqueue({ kind: "row", tab: "Delivery Zones", key: { id }, fields });
    this.persist();
    return rec;
  }
  removeZone(id) {
    this.state.zones = (this.state.zones || []).filter((z) => z.id !== id);
    this._enqueue({ kind: "row", tab: "Delivery Zones", key: { id }, remove: true });
    this.persist();
  }

  // ---------- customers ----------
  _phoneKey(p) {
    return String(p || "").replace(/\D/g, "").slice(-10);
  }
  findCustomer({ name, phone }) {
    const list = this.state.customers || [];
    const pk = this._phoneKey(phone);
    if (pk) { const c = list.find((x) => this._phoneKey(x.phone) === pk); if (c) return c; }
    const nm = String(name || "").trim().toLowerCase();
    return nm ? list.find((x) => String(x.name).trim().toLowerCase() === nm && (!pk || !this._phoneKey(x.phone))) : null;
  }
  saveCustomer(c) {
    const list = (this.state.customers = this.state.customers || []);
    const name = String(c.name || "").trim();
    if (!name) throw new Error("Customer name is required");
    const prev = c.id ? list.find((x) => x.id === c.id) : null;
    const rec = { id: c.id || newId("C"), name, phone: String(c.phone || "").trim(), notes: c.notes || "", createdAt: prev?.createdAt || c.createdAt || new Date().toISOString().slice(0, 10) };
    if (prev) Object.assign(prev, rec); else list.push(rec);
    const { id, ...fields } = rec;
    this._enqueue({ kind: "row", tab: "Customers", key: { id }, fields });
    this.persist();
    return rec;
  }
  removeCustomer(id) {
    this.state.customers = (this.state.customers || []).filter((c) => c.id !== id);
    this._enqueue({ kind: "row", tab: "Customers", key: { id }, remove: true });
    this.persist();
  }
  _rememberCustomer(o) {
    const name = String(o.customer || "").trim();
    if (!name || /^quote$|^sample order/i.test(name)) return;
    const c = this.findCustomer({ name, phone: o.phone });
    if (!c) this.saveCustomer({ name, phone: o.phone });
    else if (!c.phone && o.phone) this.saveCustomer({ ...c, phone: o.phone });
  }
  ordersOf(c) {
    const pk = this._phoneKey(c.phone);
    const nm = String(c.name || "").trim().toLowerCase();
    return this.visibleOrders.filter((o) => (pk && this._phoneKey(o.phone) === pk) || (!pk && String(o.customer || "").trim().toLowerCase() === nm));
  }
  /** Saved customers plus anyone found only in past orders. */
  customerRows() {
    const out = (this.state.customers || []).map((c) => ({ ...c, saved: true }));
    for (const o of this.visibleOrders) {
      const name = String(o.customer || "").trim();
      if (!name || /^quote$|^sample order/i.test(name)) continue;
      const pk = this._phoneKey(o.phone);
      const hit = out.find((c) => (pk && this._phoneKey(c.phone) === pk) || (!pk && !this._phoneKey(c.phone) && c.name.toLowerCase() === name.toLowerCase()));
      if (!hit) out.push({ id: "", name, phone: o.phone || "", notes: "", saved: false });
    }
    return out.map((c) => {
      const orders = this.ordersOf(c).filter((o) => o.status !== "Cancelled");
      const last = orders.map((o) => o.deliveryDate || o.createdAt).sort().pop() || "";
      return { ...c, orders: orders.length, spent: orders.reduce((s, o) => s + Number(o.quoteSnapshot?.finalQuote || 0), 0), last };
    }).sort((a, b) => String(b.last).localeCompare(String(a.last)) || a.name.localeCompare(b.name));
  }

  // ---------- payments + status ----------
  paymentOf(o) {
    const total = Number(o.quoteSnapshot?.finalQuote || 0);
    const paid = Number(o.advancePaid || 0);
    return { total, paid, balance: Math.max(total - paid, 0), status: paymentStatusOf(paid, total) };
  }
  recordPayment(id, amount) {
    const o = this.state.orders.find((x) => x.id === id);
    const a = Number(amount);
    if (!o || !(a > 0)) throw new Error("Enter the amount received");
    return this.updateOrder(id, { advancePaid: Math.round((Number(o.advancePaid || 0) + a) * 100) / 100 });
  }
  markPaidInFull(id) {
    const o = this.state.orders.find((x) => x.id === id);
    if (!o) return;
    const patch = { advancePaid: Number(o.quoteSnapshot?.finalQuote || o.advancePaid || 0) };
    if (o.status === "Delivered") patch.status = "Paid";
    return this.updateOrder(id, patch);
  }
  setStatus(id, status) {
    return this.updateOrder(id, { status });
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
    const byCat = (list) => {
      const m = new Map();
      for (const o of list) {
        const k = o.category || "Cake";
        const c = m.get(k) || { category: k, orders: 0, revenue: 0 };
        c.orders += 1; c.revenue += money(o).revenue; m.set(k, c);
      }
      return [...m.values()].sort((a, b) => b.revenue - a.revenue);
    };
    const roll = (list) => ({
      byCategory: byCat(list),
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

  profitByCategory() {
    const map = new Map();
    for (const o of this.visibleOrders) {
      if (o.status === "Cancelled") continue;
      const key = o.category || "Cake";
      const q = o.quoteSnapshot || {};
      const cur = map.get(key) || { category: key, orders: 0, qty: 0, revenue: 0, cost: 0, profit: 0 };
      cur.orders += 1;
      if (this.isUnitOrder(o)) cur.qty += Number(o.quantity) || 0;
      cur.revenue += Number(q.finalQuote || 0);
      cur.cost += Number(q.totalCost || 0);
      cur.profit += Number(q.profit || 0);
      map.set(key, cur);
    }
    return [...map.values()].sort((a, b) => b.revenue - a.revenue);
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
    "Category","Quantity","Unit price (₹)",
  ];
  const rows = orders.map((o) => {
    const q = o.quoteSnapshot || {};
    return [
      o.id, o.createdAt, o.customer, o.phone, o.size, o.layers, o.sponge, o.filling, o.frosting,
      o.topping1, o.topping2, o.topping3, o.box, o.board, o.candles ? "Yes" : "No", o.carryBag ? "Yes" : "No",
      o.delivery, o.distanceKm, o.deliveryDate, o.status, o.advancePaid, o.notes,
      q.finalQuote, q.totalCost, q.profit, q.hoursUsed, q.rawMaterialCost, q.labour, q.delivery,
      o.category || "Cake", o.quantity ?? "", o.unitPrice ?? "",
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
