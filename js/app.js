import { PRICE_KINDS } from "./store.js";
import { findAddon } from "./calculator.js";
import { Store, checkPin, setPin, isUnlocked, unlockSession, lockSession, ordersToSheetRows, toCsv } from "./store.js";
import { rupees, rupeesWhole, scaleRecipe, parseSizeKg, sizeLabel, estimateServings } from "./calculator.js";
import { SyncEngine, loadSyncConfig, saveSyncConfig, makeTransport, DEFAULT_SHEET_ID } from "./sync.js";
import { cachedGoogleToken, signInWithGoogle, clearGoogleToken } from "./transports.js";

const STATUSES = ["Enquiry", "Confirmed", "Baking", "Delivered", "Paid", "Cancelled"];
const $ = (sel, el = document) => el.querySelector(sel);

let store;
let engine;
let syncCfg;
let route = { name: "dashboard" };

const SYNC_LABEL = { local: "Local only", idle: "Connecting…", synced: "Synced", syncing: "Syncing…", pending: "Syncing…", offline: "Offline", error: "Sync error", signin: "Sign in needed" };
function fmtTime(iso) {
  if (!iso) return "never";
  const d = new Date(iso);
  const sameDay = d.toDateString() === new Date().toDateString();
  return d.toLocaleString("en-IN", sameDay ? { hour: "numeric", minute: "2-digit" } : { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
}
function syncPill() {
  const st = engine?.status || { state: "local" };
  const pending = st.pending ? ` · ${st.pending} waiting` : "";
  const when = st.state === "local" ? "" : ` · ${fmtTime(st.lastSynced)}`;
  return `<span class="sync-pill s-${st.state}" title="${escapeHtml(st.message || "")}">${SYNC_LABEL[st.state] || st.state}${when}${pending}</span>`;
}
function updateSyncPill() {
  const el = document.getElementById("sync-slot");
  if (el) el.innerHTML = syncPill();
}
async function refreshFromSheet() {
  if (!engine?.transport) return;
  await engine.pull();
  render();
}

function fmtDate(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
}

function cakeLabel(o) {
  if (store.isUnitOrder(o)) {
    return `${Number(o.quantity) || 0} × ${o.category || "Item"}`;
  }
  const cat = o.category && o.category !== "Cake" ? o.category + " · " : "";
  return cat + `${o.size || ""} ${o.sponge || ""}`.replace("Sponge", "").trim();
}
function catSlug(name) { return encodeURIComponent(String(name || "").trim()); }
const SIZE_UNIT_KEY = "bakery.v4.sizeUnit";
function sizeUnitPref() { return localStorage.getItem(SIZE_UNIT_KEY) === "g" ? "g" : "kg"; }

function iconSvg() {
  return `<img class="logo" src="./icons/icon-192.png" alt="" width="40" height="40">`;
}

function topbar() {
  return `<header class="topbar">${iconSvg()}
    <h1 class="brand">Rise & Whisk<small>Home bakery book</small></h1>
    <div class="spacer"></div>
    <div class="sync-wrap"><span id="sync-slot">${syncPill()}</span>
      ${engine?.transport ? `<button class="icon-btn" id="sync-refresh" title="Reload data from the Google Sheet" aria-label="Reload from sheet"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7"/></svg></button>` : ""}
    </div>
  </header>`;
}

function nav(active) {
  const item = (id, label, path) => `<button class="${active === id ? "on" : ""}" data-go="${path}">${icons[id] || ""}${label}</button>`;
  return `<nav class="nav">${item("home","Home","#/dashboard")}${item("orders","Orders","#/orders")}${item("quote","Quote","#/calculator")}${item("recipes","Recipes","#/recipes")}${item("more","More","#/more")}</nav>`;
}

const icons = {
  home: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 11.5 12 4l8 7.5V20a1 1 0 0 1-1 1h-5v-6H10v6H5a1 1 0 0 1-1-1z"/></svg>`,
  orders: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M7 7h10M7 12h10M7 17h6"/><rect x="4" y="4" width="16" height="16" rx="2"/></svg>`,
  quote: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 1v22M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>`,
  recipes: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 3h6a2 2 0 0 1 2 2v16l-5-2-5 2V5a2 2 0 0 1 2-2z"/></svg>`,
  more: `<svg viewBox="0 0 24 24" fill="currentColor"><circle cx="6" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="18" cy="12" r="1.6"/></svg>`,
};

function quoteBox(q) {
  const mode = q.priceOverride ? "Your price for this order" : q.menu ? "Final quote · menu price" : "Final quote";
  const note = q.priceOverride
    ? `Calculated quote was ${rupeesWhole(q.calculatedQuote)}`
    : q.menu ? `Cost-plus calculator would say ${rupeesWhole(excelRound(q))}` : "";
  return `<div class="quote-box">
    <div class="lbl">${mode}</div>
    <div class="quote-num">${rupeesWhole(q.finalQuote)}</div>
    ${note ? `<div class="tiny" style="color:#f3e2cf;margin:-6px 0 8px">${note}</div>` : ""}
    <div class="quote-grid">
      <div><span>Raw materials</span><br><strong>${rupees(q.rawMaterialCost)}</strong></div>
      <div><span>Labour</span><br><strong>${rupees(q.labour)}</strong></div>
      <div><span>Overheads</span><br><strong>${rupees(q.overheads)}</strong></div>
      <div><span>Profit</span><br><strong>${rupees(q.profit)}</strong></div>
      <div><span>Delivery${q.deliveryZone ? " · " + escapeHtml(q.deliveryZone) : ""}</span><br><strong>${rupees(q.delivery)}</strong></div>
      <div><span>Per kg (ex-del.)</span><br><strong>${rupees(q.pricePerKgExclDelivery)}</strong></div>
      ${(q.rules || []).map((r) => `<div><span>${escapeHtml(r.name)}</span><br><strong>${r.value < 0 ? "−" : "+"}${rupees(Math.abs(r.value))}</strong></div>`).join("")}
    </div>
  </div>`;
}
function excelRound(q) {
  // what the plain cost-plus calculator gives (for comparison with the menu price)
  const s = store.state.settings;
  const before = q.calculatedBeforeDelivery + (q.rulesTotal || 0) + q.delivery;
  const v = before * (1 + Number(s.gst || 0));
  const r = Number(s.rounding || 0);
  return r > 0 ? Math.ceil(v / r - 1e-12) * r : v;
}

function breakdownTable(q) {
  const rows = [
    ["Sponge", q.lines.sponge],
    ["Filling", q.lines.filling],
    ["Crumb coat", q.lines.crumbCoat],
    ["Final frosting", q.lines.frosting],
    ...q.lines.toppings.filter(t => t.name && t.name !== "None").map((t,i) => ["Topping "+(i+1), t]),
    ["Cake box", q.lines.box],
    ["Cake board", q.lines.board],
    ["Candles & knife", q.lines.candles],
    ["Carry bag", q.lines.carryBag],
    ...(q.lines.extraAddons || []).map((a) => ["Add-on", a]),
  ];
  return `<table class="breakdown"><thead><tr><th>Component</th><th>Selection</th><th class="num">Qty</th><th class="num">Cost</th></tr></thead><tbody>
    ${rows.map(([l,x]) => `<tr><td>${l}</td><td>${x.name || "—"}</td><td class="num">${x.qty ? Number(x.qty).toLocaleString("en-IN",{maximumFractionDigits:3}) + " " + (x.unit||"") : "—"}</td><td class="num">${rupees(x.cost)}</td></tr>`).join("")}
    <tr class="total-row"><td colspan="3">Ingredients (incl. wastage)</td><td class="num">${rupees(q.ingredientCost)}</td></tr>
    <tr><td colspan="3">Packaging</td><td class="num">${rupees(q.packagingCost)}</td></tr>
    <tr><td colspan="3">Total cost of making</td><td class="num">${rupees(q.totalCost)}</td></tr>
    ${q.menu ? q.menu.parts.map((p) => `<tr><td colspan="3">Menu price · ${escapeHtml(p.label)}</td><td class="num">${rupees(p.price)}</td></tr>`).join("") : ""}
    ${(q.rules || []).map((r) => `<tr><td colspan="3">${escapeHtml(r.name)} (${r.type === "%" ? r.amount + "%" : "₹" + r.amount} ${r.effect === "Discount" ? "off" : "extra"})</td><td class="num">${r.value < 0 ? "−" : ""}${rupees(Math.abs(r.value))}</td></tr>`).join("")}
    <tr><td colspan="3">Before rounding</td><td class="num">${rupees(q.beforeRounding)}</td></tr>
    <tr class="total-row"><td colspan="3">Final quote</td><td class="num">${rupeesWhole(q.finalQuote)}</td></tr>
  </tbody></table>
  <p class="tiny">${q.weightOk ? "Weight check OK — sponge + filling + crumb coat + frosting matches cake weight." : "CHECK — component weights do not add up to cake weight."}
  Hours used: ${q.hoursUsed} (suggested ${q.hoursSuggested}). GST ${rupees(q.gst)}.</p>`;
}

function options(list, selected, labelFn = (x) => x) {
  return list.map((x) => {
    const v = typeof x === "string" ? x : labelFn(x);
    return `<option ${v === selected ? "selected" : ""}>${escapeHtml(v)}</option>`;
  }).join("");
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function calcInputsFromForm(form, settings) {
  const g = (n) => form.elements[n]?.value;
  let size = g("size");
  let weightKg = 0, servings;
  if (size === "__custom") {
    const unit = g("sizeUnit") === "g" ? "g" : "kg";
    const amt = Number(g("customSize"));
    weightKg = amt > 0 ? (unit === "g" ? amt / 1000 : amt) : 0;
    size = amt > 0 ? `${amt} ${unit}` : "";
    servings = estimateServings(weightKg, settings.sizes);
  } else {
    const sz = settings.sizes.find((s) => s.name === size);
    weightKg = sz ? sz.kg : parseSizeKg(size);
  }
  return {
    category: g("category") || "Cake",
    customer: g("customer"),
    phone: g("phone"),
    size,
    weightKg,
    ...(servings !== undefined ? { servings } : {}),
    layers: Number(g("layers")),
    sponge: g("sponge"),
    filling: g("filling"),
    frosting: g("frosting"),
    topping1: g("topping1"),
    topping2: g("topping2"),
    topping3: g("topping3"),
    box: g("box"),
    board: g("board"),
    candles: form.elements.candles?.checked,
    carryBag: form.elements.carryBag?.checked,
    delivery: g("delivery"),
    distanceKm: Number(g("distanceKm") || 0),
    hoursOverride: g("hoursOverride") === "" ? null : g("hoursOverride"),
    priceOverride: g("priceOverride") === "" || g("priceOverride") == null ? null : Number(g("priceOverride")),
    zone: g("delivery") === "Home delivery" ? g("zone") || "" : "",
    rules: [...form.querySelectorAll('input[name="rule"]:checked')].map((x) => x.value),
    extraAddons: [...form.querySelectorAll('input[name="addon"]:checked')].map((x) => x.value),
    deliveryDate: g("deliveryDate"),
    advancePaid: Number(g("advancePaid") || 0),
    notes: g("notes"),
    status: g("status") || "Enquiry",
  };
}

function cakeForm(pre = {}, settings) {
  const menu = settings.recipeMenu;
  const sponges = menu.filter((r) => r.type === "Sponge").map((r) => r.name);
  const fillings = menu.filter((r) => r.type === "Filling").map((r) => r.name);
  const frostings = menu.filter((r) => r.type === "Frosting").map((r) => r.name);
  const tops = settings.toppings.map((t) => t.name);
  if (!tops.includes("None")) tops.unshift("None");
  const first = (list, want) => (list.includes(want) ? want : list[0]);
  const p = { size: first(settings.sizes.map((x) => x.name), "0.5 kg"), layers: 2, sponge: first(sponges, "Chocolate Sponge"), filling: first(fillings, "Chocolate Ganache"), frosting: first(frostings, "Chocolate Ganache Frosting"), topping1: "None", topping2: "None", topping3: "None", box: first(settings.boxes, "Cake box 8x8x5 in (up to 1 kg)"), board: first(settings.boards, "Cake board 8 in"), delivery: "Home delivery", distanceKm: 8, carryBag: true, candles: false, status: "Enquiry", ...pre };
  const candle = findAddon(settings.addons, "candles");
  const bag = findAddon(settings.addons, "carryBag");
  const extra = (settings.addons || []).filter((a) => a !== candle && a !== bag);
  const zones = store.state.zones || [];
  const ruleList = store.state.rules || [];
  const ticked = p.rules ?? store.defaultRules();
  const shownRules = ruleList.filter((r) => r.active || ticked.includes(r.name));
  const custs = store.customerRows();
  const keep = (list, v) => (v && !list.includes(v) ? [...list, v] : list); // an old order keeps its removed option
  // size: presets from the sheet, or a manual size in grams / kg
  const presetNames = settings.sizes.map((x) => x.name);
  const isCustom = p.size && !presetNames.includes(p.size) && parseSizeKg(p.size) > 0;
  const unit = isCustom ? (/\d\s*k/i.test(p.size) ? "kg" : "g") : sizeUnitPref();
  const customAmt = isCustom ? String(p.size).match(/[\d.]+/)[0] : "";
  const sizeOpts = [...settings.sizes, ...(!isCustom && p.size && !presetNames.includes(p.size) ? [{ name: p.size, kg: 0 }] : [])]
    .map((z) => `<option value="${escapeHtml(z.name)}" data-kg="${Number(z.kg) || 0}" ${!isCustom && z.name === p.size ? "selected" : ""}>${escapeHtml(unit === "g" && z.kg ? sizeLabel(z.kg, "g") : z.name)}</option>`).join("")
    + `<option value="__custom" ${isCustom ? "selected" : ""}>Custom size…</option>`;
  return `<form id="cake-form">
    <input type="hidden" name="category" value="${escapeHtml(p.category || "Cake")}">
    <datalist id="cust-list">${custs.map((c) => `<option value="${escapeHtml(c.name)}">${escapeHtml(c.phone || "")}</option>`).join("")}</datalist>
    <div class="card">
      <h2>Order details</h2>
      <div class="row-2">
        <div><label>Customer</label><input name="customer" list="cust-list" autocomplete="off" value="${escapeHtml(p.customer||"")}" placeholder="Name (pick a saved one)" required></div>
        <div><label>Phone</label><input name="phone" type="tel" value="${escapeHtml(p.phone||"")}" placeholder="WhatsApp number"></div>
      </div>
      <p class="tiny" id="cust-hint"></p>
      <div class="row-2">
        <div><label>Delivery date</label><input name="deliveryDate" type="date" value="${escapeHtml(p.deliveryDate||"")}"></div>
        <div><label>Status</label><select name="status">${options(STATUSES, p.status)}</select></div>
      </div>
      <div class="row-2">
        <div><label>Advance paid (₹)</label><input name="advancePaid" type="number" min="0" step="1" value="${p.advancePaid||0}"></div>
        <div><label>Notes</label><input name="notes" value="${escapeHtml(p.notes||"")}" placeholder="Message, colour, time"></div>
      </div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Cake</h2>
      <div class="row-2">
        <div><label class="size-lbl">Size <span class="seg" id="size-unit" role="group" aria-label="Size unit"><button type="button" data-unit="g" class="${unit === "g" ? "on" : ""}">g</button><button type="button" data-unit="kg" class="${unit === "kg" ? "on" : ""}">kg</button></span></label>
          <input type="hidden" name="sizeUnit" value="${unit}">
          <select name="size">${sizeOpts}</select>
          <div class="custom-size" id="custom-size" ${isCustom ? "" : "hidden"}><input name="customSize" type="number" min="0" step="any" inputmode="decimal" value="${escapeHtml(customAmt)}" placeholder="${unit === "g" ? "e.g. 750" : "e.g. 1.2"}"><span class="unit-tag">${unit}</span></div>
          <p class="tiny" id="size-hint"></p></div>
        <div><label>Layers</label><select name="layers">${options(keep(settings.layers.map(String), String(p.layers)), String(p.layers))}</select></div>
      </div>
      <label>Sponge</label><select name="sponge">${options(keep(sponges, p.sponge), p.sponge)}</select>
      <label>Filling</label><select name="filling">${options(keep(fillings, p.filling), p.filling)}</select>
      <label>Frosting</label><select name="frosting">${options(keep(frostings, p.frosting), p.frosting)}</select>
      <div class="row-3">
        <div><label>Topping 1</label><select name="topping1">${options(keep(tops, p.topping1||"None"), p.topping1||"None")}</select></div>
        <div><label>Topping 2</label><select name="topping2">${options(keep(tops, p.topping2||"None"), p.topping2||"None")}</select></div>
        <div><label>Topping 3</label><select name="topping3">${options(keep(tops, p.topping3||"None"), p.topping3||"None")}</select></div>
      </div>
      <div class="row-2">
        <div><label>Cake box</label><select name="box">${options(keep(settings.boxes, p.box), p.box)}</select></div>
        <div><label>Cake board</label><select name="board">${options(keep(settings.boards, p.board), p.board)}</select></div>
      </div>
      ${candle ? `<label class="check"><input type="checkbox" name="candles" ${p.candles?"checked":""}> ${escapeHtml(candle.name)}</label>` : `<input type="hidden" name="candles">`}
      ${bag ? `<label class="check"><input type="checkbox" name="carryBag" ${p.carryBag?"checked":""}> ${escapeHtml(bag.name)}</label>` : `<input type="hidden" name="carryBag">`}
      ${extra.map((a) => `<label class="check"><input type="checkbox" name="addon" value="${escapeHtml(a.name)}" ${(p.extraAddons||[]).includes(a.name)?"checked":""}> ${escapeHtml(a.name)}</label>`).join("")}
      <div class="row-2">
        <div><label>Delivery</label><select name="delivery">${options(["Home delivery","Pickup"], p.delivery)}</select></div>
        ${zones.length || p.zone ? `<div><label>Delivery charge</label><select name="zone"><option value="">By distance (₹/km)</option>${keep(zones.map((z) => z.name), p.zone).map((z) => { const zz = zones.find((x) => x.name === z); return `<option value="${escapeHtml(z)}" ${z === p.zone ? "selected" : ""}>${escapeHtml(z)}${zz ? " · ₹" + zz.charge : ""}</option>`; }).join("")}</select></div>` : ""}
      </div>
      <div><label>Distance, one way (km)</label><input name="distanceKm" type="number" min="0" step="0.1" value="${p.distanceKm||0}"></div>
      ${shownRules.length ? `<label>Pricing rules</label><div class="chips">${shownRules.map((r) => `<label class="chip-check"><input type="checkbox" name="rule" value="${escapeHtml(r.name)}" ${ticked.includes(r.name)?"checked":""}> ${escapeHtml(r.name)} <span class="muted">${r.effect === "Discount" ? "−" : "+"}${r.type === "%" ? r.amount + "%" : "₹" + r.amount}</span></label>`).join("")}</div>` : ""}
      <label>Hours needed (leave blank for suggested)</label>
      <input name="hoursOverride" type="number" min="0" step="0.25" value="${p.hoursOverride ?? ""}" placeholder="Auto from size, layers, toppings">
      <label>Your price for this order (₹) — optional</label>
      <input name="priceOverride" type="number" min="0" step="1" value="${p.priceOverride ?? ""}" placeholder="Leave blank to use the quote below">
    </div>
    <div id="live-quote" style="margin-top:12px"></div>
  </form>`;
}

function renderLiveQuote(form) {
  const input = calcInputsFromForm(form, store.state.settings);
  const q = store.quote(input);
  const hint = $("#size-hint");
  if (hint) hint.textContent = input.weightKg > 0 ? `${sizeLabel(q.weightKg, "kg")} = ${sizeLabel(q.weightKg, "g")}${q.servings ? " · serves " + q.servings : ""}` : form.elements.size?.value === "__custom" ? "Type the cake weight" : "";
  $("#live-quote").innerHTML = `<div class="card">${quoteBox(q)}${breakdownTable(q)}</div>`;
  return { input, q };
}

function viewDashboard() {
  const d = store.dashboard();
  const upcoming = d.upcoming.length
    ? d.upcoming.map((o) => `<div class="order-item" data-go="#/order/${o.id}"><div><h3>${escapeHtml(o.customer||"No name")}</h3><div class="meta">${fmtDate(o.deliveryDate)} · ${escapeHtml(cakeLabel(o))}</div></div><div><span class="pill ${o.status}">${o.status}</span><div class="meta" style="text-align:right;margin-top:4px">${rupeesWhole(o.quoteSnapshot?.finalQuote||0)}</div></div></div>`).join("")
    : `<p class="empty">No upcoming deliveries yet.</p>`;
  const low = store.lowStock();
  return `${topbar()}<h2 class="page-title">Today's bake</h2>
    <div class="stats">
      <div class="card"><div class="stat-lbl">This week · orders</div><div class="stat-val">${d.week.count}</div></div>
      <div class="card"><div class="stat-lbl">This week · revenue</div><div class="stat-val">${rupeesWhole(d.week.revenue)}</div></div>
      <div class="card"><div class="stat-lbl">This month · profit</div><div class="stat-val">${rupeesWhole(d.month.profit)}</div></div>
      <div class="card"><div class="stat-lbl">Pending / low stock</div><div class="stat-val">${d.pending} <span class="muted">/</span> ${d.lowStock}</div></div>
    </div>
    <div class="card" style="margin-top:14px">
      <h2>This month</h2>
      <div class="quote-grid" style="color:inherit">
        <div><span class="muted">Revenue</span><br><strong>${rupees(d.month.revenue)}</strong></div>
        <div><span class="muted">Cost</span><br><strong>${rupees(d.month.cost)}</strong></div>
        <div><span class="muted">Profit</span><br><strong>${rupees(d.month.profit)}</strong></div>
        <div><span class="muted">Orders</span><br><strong>${d.month.count}</strong></div>
      </div>
      ${d.month.byCategory.length > 1 || (d.month.byCategory[0] && d.month.byCategory[0].category !== "Cake") ? `<p class="tiny" style="margin-top:8px">${d.month.byCategory.map((c) => `${escapeHtml(c.category)}: ${c.orders} · ${rupeesWhole(c.revenue)}`).join(" &nbsp;·&nbsp; ")}</p>` : ""}
    </div>
    ${low.length ? `<div class="card" style="margin-top:14px"><h2>Low stock <span class="badge-count">${low.length}</span></h2>
      ${low.map(i => `<div class="inv-item" data-go="#/inventory"><strong>${escapeHtml(i.item)}</strong> <span class="pill reorder">REORDER</span><div class="meta">${i.current} ${i.unit} left · reorder at ${i.reorderLevel}</div></div>`).join("")}
    </div>` : ""}
    <div class="card" style="margin-top:14px"><h2>Upcoming</h2><div class="list">${upcoming}</div>
      <div class="btn-row"><button class="btn" data-go="#/order/new">New order</button><button class="btn secondary" data-go="#/calculator">Price a cake</button></div>
    </div>
    ${nav("home")}`;
}

function viewOrders() {
  const q = (route.q || "").toLowerCase();
  const st = route.status || "All";
  let list = store.visibleOrders;
  if (st !== "All") list = list.filter((o) => o.status === st);
  if (q) list = list.filter((o) => (o.customer+" "+o.phone+" "+o.id+" "+cakeLabel(o)).toLowerCase().includes(q));
  const items = list.length
    ? list.map((o) => `<div class="order-item" data-go="#/order/${o.id}">
        <div><h3>${escapeHtml(o.customer||"No name")}</h3><div class="meta">${o.id} · ${fmtDate(o.deliveryDate||o.createdAt)} · ${escapeHtml(cakeLabel(o))}</div></div>
        <div style="text-align:right"><span class="pill ${o.status}">${o.status}</span><div class="meta" style="margin-top:4px">${rupeesWhole(o.quoteSnapshot?.finalQuote||0)}</div></div>
        <div class="quick-row">${payPill(o)}${nextStatus(o.status) ? `<button class="mini-btn" data-next="${o.id}" data-status="${nextStatus(o.status)}">Mark ${nextStatus(o.status)}</button>` : ""}</div>
      </div>`).join("")
    : `<p class="empty">No orders yet. Log the first one from a quote or the button below.</p>`;
  return `${topbar()}<h2 class="page-title">Orders</h2>
    <input class="search" id="order-search" placeholder="Search name, phone or id" value="${escapeHtml(route.q||"")}">
    <div class="filters" id="status-filters">${["All", ...STATUSES].map((s) => `<button data-status="${s}" class="${st===s?"on":""}">${s}</button>`).join("")}</div>
    <div class="list">${items}</div>
    <div class="btn-row"><button class="btn full" data-go="#/order/new">New order</button></div>
    ${nav("orders")}`;
}

function viewOrderNew() {
  const cats = store.activeCategories();
  const qs = route.cust ? `?cust=${encodeURIComponent(route.cust)}` : "";
  if (route.cat) {
    const cat = store.category(route.cat);
    if (!cat) return `${topbar()}<p class="tiny"><a href="#/order/new${qs}">← Categories</a></p><p class="empty">“${escapeHtml(route.cat)}” isn't a category any more.</p>${nav("orders")}`;
    return cat.pricing === "Calculator" ? viewOrderNewCake(cat) : viewOrderNewUnit(cat);
  }
  return `${topbar()}<h2 class="page-title">New order</h2>
    <p class="muted">What is she making?</p>
    <div class="cat-grid">${cats.map((c) => `<button class="cat-tile" data-go="#/order/new/${catSlug(c.name)}${qs}">
        <strong>${escapeHtml(c.name)}</strong>
        <span>${c.pricing === "Calculator" ? "Size, layers, flavours, toppings" : `${rupeesWhole(c.unitPrice)} per ${escapeHtml(c.unit || "piece")}`}</span>
      </button>`).join("") || `<p class="empty">No active categories. Add one in More → Menu &amp; prices.</p>`}</div>
    <p class="tiny">Categories come from the <strong>Categories</strong> tab of her sheet${store.sheetMode ? "" : " (on this phone until the sheet is connected)"}; add or change them in <a href="#/menu">Menu &amp; prices</a>.</p>
    ${nav("orders")}`;
}
function prefillCustomer() {
  const c = route.cust ? store.customerRows().find((x) => custKey(x) === route.cust) : null;
  return c ? { customer: c.name, phone: c.phone } : {};
}
function viewOrderNewCake(cat) {
  const qs = route.cust ? `?cust=${encodeURIComponent(route.cust)}` : "";
  return `${topbar()}<p class="tiny"><a href="#/order/new${qs}">← Categories</a></p><h2 class="page-title">New ${escapeHtml(cat.name.toLowerCase())} order</h2>
    ${cakeForm({ ...prefillCustomer(), category: cat.name }, store.state.settings)}
    <div class="btn-row"><button class="btn" id="save-order">Save order</button><button class="btn secondary" data-go="#/orders">Cancel</button></div>
    ${nav("orders")}`;
}
function viewOrderNewUnit(cat) {
  const qs = route.cust ? `?cust=${encodeURIComponent(route.cust)}` : "";
  return `${topbar()}<p class="tiny"><a href="#/order/new${qs}">← Categories</a></p><h2 class="page-title">New ${escapeHtml(cat.name.toLowerCase())} order</h2>
    ${unitForm({ ...prefillCustomer(), category: cat.name, unitPrice: cat.unitPrice })}
    <div class="btn-row"><button class="btn" id="save-unit-order">Save order</button><button class="btn secondary" data-go="#/orders">Cancel</button></div>
    ${nav("orders")}`;
}

/** Order form for per-unit categories (granola bars, cookies…): customer + quantity. */
function unitForm(pre = {}) {
  const cat = store.category(pre.category) || { name: pre.category, unit: "piece", unitPrice: pre.unitPrice };
  const p = { quantity: 1, delivery: "Pickup", deliveryCharge: 0, status: "Enquiry", advancePaid: 0, ...pre };
  const unitPrice = p.unitPrice != null && p.unitPrice !== "" ? Number(p.unitPrice) : Number(cat.unitPrice) || 0;
  const custs = store.customerRows();
  const u = cat.unit || "piece";
  return `<form id="unit-form">
    <input type="hidden" name="category" value="${escapeHtml(p.category)}">
    <input type="hidden" name="unitPrice" value="${unitPrice}">
    <datalist id="cust-list">${custs.map((c) => `<option value="${escapeHtml(c.name)}">${escapeHtml(c.phone || "")}</option>`).join("")}</datalist>
    <div class="card">
      <h2>${escapeHtml(p.category)}</h2>
      <div class="row-2">
        <div><label>How many ${escapeHtml(u)}s</label><input name="quantity" type="number" min="1" step="1" inputmode="numeric" value="${escapeHtml(p.quantity)}" class="qty-input" required></div>
        <div><label>Price per ${escapeHtml(u)}</label><div class="meta" style="padding:10px 0"><strong>${rupeesWhole(unitPrice)}</strong> <span class="tiny">from the Categories tab</span></div></div>
      </div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Order details</h2>
      <div class="row-2">
        <div><label>Customer</label><input name="customer" list="cust-list" autocomplete="off" value="${escapeHtml(p.customer || "")}" placeholder="Name (pick a saved one)" required></div>
        <div><label>Phone</label><input name="phone" type="tel" value="${escapeHtml(p.phone || "")}" placeholder="WhatsApp number"></div>
      </div>
      <p class="tiny" id="cust-hint"></p>
      <div class="row-2">
        <div><label>Delivery date</label><input name="deliveryDate" type="date" value="${escapeHtml(p.deliveryDate || "")}"></div>
        <div><label>Status</label><select name="status">${options(STATUSES, p.status)}</select></div>
      </div>
      <div class="row-2">
        <div><label>Delivery</label><select name="delivery">${options(["Pickup", "Home delivery"], p.delivery)}</select></div>
        <div><label>Delivery charge (₹)</label><input name="deliveryCharge" type="number" min="0" step="1" value="${Number(p.deliveryCharge) || 0}"></div>
      </div>
      <div class="row-2">
        <div><label>Advance paid (₹)</label><input name="advancePaid" type="number" min="0" step="1" value="${p.advancePaid || 0}"></div>
        <div><label>Notes</label><input name="notes" value="${escapeHtml(p.notes || "")}" placeholder="Flavour, packing, time"></div>
      </div>
    </div>
    <div id="unit-quote" style="margin-top:12px"></div>
  </form>`;
}
function unitInputsFromForm(form) {
  const g = (n) => form.elements[n]?.value;
  return {
    category: g("category"), quantity: Number(g("quantity")) || 0, unitPrice: Number(g("unitPrice")) || 0,
    customer: g("customer"), phone: g("phone"), deliveryDate: g("deliveryDate"), status: g("status") || "Enquiry",
    delivery: g("delivery") || "Pickup", deliveryCharge: g("delivery") === "Home delivery" ? Number(g("deliveryCharge")) || 0 : 0,
    advancePaid: Number(g("advancePaid") || 0), notes: g("notes"),
  };
}
function unitQuoteBox(q) {
  const u = q.unit || "piece";
  return `<div class="quote-box">
    <div class="lbl">Order total · ${escapeHtml(q.category)}</div>
    <div class="quote-num">${rupeesWhole(q.finalQuote)}</div>
    <div class="quote-grid">
      <div><span>Quantity</span><br><strong>${q.quantity} ${escapeHtml(u)}${q.quantity === 1 ? "" : "s"}</strong></div>
      <div><span>Price each</span><br><strong>${rupees(q.unitPrice)}</strong></div>
      <div><span>Subtotal</span><br><strong>${rupees(q.subtotal)}</strong></div>
      <div><span>Delivery</span><br><strong>${rupees(q.delivery)}</strong></div>
      <div><span>Cost</span><br><strong>${q.costKnown || q.totalCost ? rupees(q.totalCost) : "not set"}</strong></div>
      <div><span>Profit</span><br><strong>${rupees(q.profit)}</strong></div>
    </div>
  </div>
  ${q.costKnown || q.totalCost ? "" : `<p class="tiny">Add a “Cost per unit” for ${escapeHtml(q.category)} in <a href="#/menu">Menu &amp; prices</a> to see real profit.</p>`}`;
}
function renderUnitQuote(form) {
  const input = unitInputsFromForm(form);
  const q = store.unitQuote(input);
  const box = $("#unit-quote");
  if (box) box.innerHTML = unitQuoteBox(q);
  return { input, q };
}

const NEXT = { Enquiry: "Confirmed", Confirmed: "Baking", Baking: "Delivered", Delivered: "Paid" };
function nextStatus(st) { return NEXT[st] || ""; }
function payPill(o) {
  const p = store.paymentOf(o);
  const cls = p.status === "Paid in full" ? "OK" : p.status === "Advance paid" ? "Confirmed" : "Enquiry";
  return `<span class="pill ${cls}">${p.status}${p.status !== "Paid in full" && p.balance ? " · due " + rupeesWhole(p.balance) : ""}</span>`;
}

function orderHead(o) {
  return `${topbar()}<h2 class="page-title">${escapeHtml(o.customer||"Order")}</h2>
    <p class="meta"><span class="pill cat">${escapeHtml(o.category || "Cake")}</span> ${o.id} · Created ${fmtDate(o.createdAt)}${o.phone ? ` · <a href="tel:${escapeHtml(o.phone)}">${escapeHtml(o.phone)}</a>` : ""}</p>
    <div class="card" style="margin-bottom:12px">
      <h2>Status</h2>
      <div class="chips" id="status-chips">${STATUSES.map((st) => `<button class="chip ${o.status === st ? "on" : ""}" data-set-status="${st}" data-id="${o.id}">${st}</button>`).join("")}</div>
    </div>`;
}
function viewUnitOrder(o) {
  const q = o.quoteSnapshot || store.quote(o);
  return `${orderHead(o)}
    ${unitQuoteBox(q)}
    ${paymentCard(o, q)}
    <div style="margin-top:12px">${unitForm(o)}</div>
    <div class="btn-row">
      <button class="btn" id="save-unit-order" data-id="${o.id}">Update order</button>
      <button class="btn danger" id="delete-order" data-id="${o.id}">Delete</button>
    </div>
    ${nav("orders")}`;
}
function viewOrder(id) {
  const o = store.state.orders.find((x) => x.id === id);
  if (!o) return `${topbar()}<p>Order not found.</p>${nav("orders")}`;
  if (store.isUnitOrder(o)) return viewUnitOrder(o);
  const q = o.quoteSnapshot || store.quote(o);
  return `${orderHead(o)}
    ${quoteBox(q)}
    ${paymentCard(o, q)}
    ${(() => { const live = store.quote(o); return q.fromSheet && Math.round(live.finalQuote) !== Math.round(q.finalQuote) ? `<p class="tiny warn-msg">Saved quote is ${rupeesWhole(q.finalQuote)} (from the sheet). With today’s prices it would be ${rupeesWhole(live.finalQuote)}. Tap “Update order” to re-price.</p>` : ""; })()}
    ${cakeForm(o, store.state.settings)}
    <div class="btn-row">
      <button class="btn" id="save-order" data-id="${o.id}">Update order</button>
      <button class="btn secondary" id="use-stock" data-id="${o.id}">Deduct ingredients from stock</button>
      <button class="btn danger" id="delete-order" data-id="${o.id}">Delete</button>
    </div>
    <div class="card" style="margin-top:12px">${breakdownTable(q)}</div>
    ${nav("orders")}`;
}
function paymentCard(o, q) {
  const pay = store.paymentOf({ ...o, quoteSnapshot: q });
  return `<div class="card" style="margin-top:12px">
      <h2>Payment ${payPill({ ...o, quoteSnapshot: q })}</h2>
      <div class="quote-grid" style="color:inherit">
        <div><span class="muted">Order total</span><br><strong>${rupees(pay.total)}</strong></div>
        <div><span class="muted">Received so far</span><br><strong>${rupees(pay.paid)}</strong></div>
        <div><span class="muted">Balance due</span><br><strong>${rupees(pay.balance)}</strong></div>
        <div><span class="muted">Profit</span><br><strong>${rupees(q.profit)}</strong></div>
      </div>
      <div class="row-2" style="margin-top:6px">
        <div><label>Payment received (₹)</label><input id="pay-amt" type="number" min="0" step="1" placeholder="${pay.paid ? "Balance " + Math.round(pay.balance) : "Advance"}"></div>
        <div><label>&nbsp;</label><button class="btn secondary full" id="pay-add" data-id="${o.id}">Add payment</button></div>
      </div>
      <div class="btn-row"><button class="btn sage" id="pay-full" data-id="${o.id}" ${pay.balance ? "" : "disabled"}>Paid in full</button></div>
      <p class="tiny">Saved in the Orders tab: “Advance paid (₹)” is the total received; “Payment status” and “Balance (₹)” are filled in for you.</p>
    </div>`;
}

function viewCalculator() {
  const sample = store.seed.sampleOrder;
  return `${topbar()}<h2 class="page-title">Price calculator</h2>
    <p class="muted">Matches her Excel calculator exactly — wastage, labour, overheads, markup, GST, rounding and delivery.</p>
    ${cakeForm({ ...sample, customer: "", phone: "", deliveryDate: "" }, store.state.settings)}
    <div class="btn-row">
      <button class="btn" id="save-quote-order">Save as new order</button>
      <button class="btn secondary" id="load-sample">Load sample birthday cake</button>
    </div>
    ${nav("quote")}`;
}

function viewRecipes() {
  const costs = store.recipeCosts();
  const names = [...new Set(store.state.recipes.map((r) => r.name))];
  const chosen = route.recipe || names[0];
  const lines = store.state.recipes.filter((r) => r.name === chosen);
  const type = store.state.settings.recipeMenu.find((m) => m.name === chosen)?.type || lines[0]?.type || "";
  const target = Number(route.kg || 1);
  const scaled = scaleRecipe(chosen, target, store.state.recipes, store.state.rawMaterials);
  const matNames = store.state.rawMaterials.map((m) => m.name);
  return `${topbar()}<h2 class="page-title">Recipe costing</h2>
    <div class="card">
      <label>Recipe</label>
      <select id="recipe-pick">${names.map((n) => `<option ${n===chosen?"selected":""}>${escapeHtml(n)}</option>`).join("")}</select>
      <p><span class="pill">${escapeHtml(type)}</span> · Cost per 1 kg: <strong>${rupees(costs[chosen]||0)}</strong></p>
      <table class="breakdown"><thead><tr><th>Ingredient</th><th class="num">Qty / 1 kg</th><th class="num">Cost</th><th></th></tr></thead><tbody>
        ${lines.map((l) => {
          const mat = store.state.rawMaterials.find((m) => m.name === l.ingredient);
          const rate = mat ? (Number(mat.packPrice)/Number(mat.packSize)||0) : 0;
          return `<tr><td><select class="mini" data-rline="${escapeHtml(l.key)}" data-field="ingredient">${options(matNames, l.ingredient)}</select>${mat?.estimated?'<span class="tag-est">ESTIMATED</span>':""}${!mat?'<span class="tag-est">MISSING</span>':""}</td><td class="num"><input class="mini num" data-rline="${escapeHtml(l.key)}" data-field="qtyPerKg" type="number" step="any" min="0" value="${l.qtyPerKg}"> ${l.unit||""}</td><td class="num">${rupees(l.qtyPerKg*rate)}</td><td><button class="x-btn sm" data-rdel="${escapeHtml(l.key)}" title="Remove this line" aria-label="Remove line">✕</button></td></tr>`;
        }).join("")}
        <tr class="total-row"><td colspan="2">Per 1 kg</td><td class="num">${rupees(costs[chosen]||0)}</td><td></td></tr>
      </tbody></table>
      <p class="tiny">Edit a quantity or ingredient and it is saved${store.sheetMode ? " to the Recipes tab of the sheet" : ""}. Quantities are for 1 kg of this recipe, in the item’s own unit.</p>
      <div class="row-3">
        <div><label>Add ingredient</label><select id="rl-ing">${options(matNames, matNames[0])}</select></div>
        <div><label>Qty / 1 kg</label><input id="rl-qty" type="number" step="any" min="0" value="0"></div>
        <div><label>&nbsp;</label><button class="btn secondary" id="rl-add" data-recipe="${escapeHtml(chosen)}" data-type="${escapeHtml(type)}">Add line</button></div>
      </div>
      <label>Method notes</label>
      <textarea id="r-notes" data-recipe="${escapeHtml(chosen)}" data-type="${escapeHtml(type)}" placeholder="Steps, oven temperature, baking time, tips…">${escapeHtml(store.recipeNotes(chosen))}</textarea>
      <p class="tiny">Notes save when you leave the box${store.sheetMode ? " (Recipe Notes tab)" : ""}.</p>
      <div class="btn-row"><button class="btn ghost" data-go="#/recipe-new?type=${encodeURIComponent(type || "Sponge")}">+ New recipe</button></div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Batch scaler</h2>
      <p class="tiny">Enter how many kilograms of this recipe you want to make. You get the scaled list and how many packs to buy.</p>
      <label>Target quantity (kg)</label>
      <input id="scale-kg" type="number" min="0.01" step="0.01" value="${target}">
      <p>For <strong>${target} kg</strong> of ${escapeHtml(chosen)}: ingredient cost <strong>${rupees(scaled.totalCost)}</strong>. If you buy full packs: <strong>${rupees(scaled.packSpend)}</strong>.</p>
      <table class="breakdown"><thead><tr><th>Ingredient</th><th class="num">Scaled qty</th><th class="num">Packs to buy</th><th class="num">Pack spend</th></tr></thead><tbody>
        ${scaled.lines.map((l) => `<tr><td>${escapeHtml(l.ingredient)}${l.estimated?'<span class="tag-est">ESTIMATED</span>':""}${l.missing?'<span class="tag-est">MISSING</span>':""}</td><td class="num">${Number(l.qty).toLocaleString("en-IN",{maximumFractionDigits:2})} ${l.unit}</td><td class="num">${l.packsToBuy} × ${l.packSize} ${l.unit}</td><td class="num">${rupees(l.packSpend)}</td></tr>`).join("")}
      </tbody></table>
    </div>
    ${nav("recipes")}`;
}

function viewMaterials() {
  const cats = [...new Set(store.state.rawMaterials.map((m) => m.category))];
  const rows = store.state.rawMaterials.map((m) => {
    const rate = Number(m.packPrice) / Number(m.packSize) || 0;
    return `<div class="mat-item">
      <h3>${escapeHtml(m.name)} ${m.estimated?'<span class="tag-est">ESTIMATED — confirm</span>':""}</h3>
      <div class="meta">${escapeHtml(m.category)} · ₹${rate.toLocaleString("en-IN",{maximumFractionDigits:4})} / ${escapeHtml(m.unit)}</div>
      <div class="row-3" style="margin-top:8px">
        <div><label>Pack size</label><input data-mat="${escapeHtml(m.name)}" data-field="packSize" type="number" min="0" step="any" value="${m.packSize}"></div>
        <div><label>Pack price (₹)</label><input data-mat="${escapeHtml(m.name)}" data-field="packPrice" type="number" min="0" step="any" value="${m.packPrice}"></div>
        <div><label>Unit</label><input data-mat="${escapeHtml(m.name)}" data-field="unit" value="${escapeHtml(m.unit)}"></div>
      </div>
      ${m.notes ? `<p class="tiny">${escapeHtml(m.notes)}</p>` : ""}
    </div>`;
  }).join("");
  return `${topbar()}<h2 class="page-title">Raw materials</h2>
    <p class="muted">Change a pack price and every recipe cost and cake quote updates. Confirm any item marked ESTIMATED.</p>
    <div class="list">${rows}</div>
    <div class="card" style="margin-top:12px">
      <h2>Add item</h2>
      <div class="row-2"><div><label>Category</label><select id="new-cat">${options(cats, cats[0])}</select></div>
      <div><label>Item name</label><input id="new-name"></div></div>
      <div class="row-3"><div><label>Pack size</label><input id="new-size" type="number" value="1000"></div>
      <div><label>Unit</label><input id="new-unit" value="g"></div>
      <div><label>Pack price (₹)</label><input id="new-price" type="number" value="0"></div></div>
      <div class="btn-row"><button class="btn" id="add-mat">Add item</button></div>
    </div>
    ${nav("more")}`;
}

function viewInventory() {
  const rows = store.inventoryRows();
  const value = rows.reduce((s, r) => s + r.stockValue, 0);
  const low = rows.filter((r) => r.status === "REORDER").length;
  const items = rows.map((r) => `<div class="inv-item">
    <div style="display:flex;justify-content:space-between;gap:8px"><h3>${escapeHtml(r.item)} ${r.estimated?'<span class="tag-est">ESTIMATED</span>':""}</h3><span class="pill ${r.status}">${r.status}</span></div>
    <div class="meta">${escapeHtml(r.category)} · on hand <strong>${r.current}</strong> ${r.unit} · value ${rupees(r.stockValue)}</div>
    <div class="row-2" style="margin-top:8px">
      <div><label>Opening</label><input data-inv="${escapeHtml(r.item)}" data-field="opening" type="number" step="any" value="${r.opening}"></div>
      <div><label>Purchased</label><input data-inv="${escapeHtml(r.item)}" data-field="purchased" type="number" step="any" value="${r.purchased}"></div>
      <div><label>Used</label><input data-inv="${escapeHtml(r.item)}" data-field="used" type="number" step="any" value="${r.used}"></div>
      <div><label>Low-stock alert at (${escapeHtml(r.unit)})</label><input data-inv="${escapeHtml(r.item)}" data-field="reorderLevel" type="number" step="any" min="0" value="${r.reorderLevel}"></div>
    </div>
  </div>`).join("");
  const mats = store.state.rawMaterials;
  const pre = mats.find((m) => m.name === route.item) || null;
  const recent = (store.state.purchases || []).slice(-8).reverse();
  return `${topbar()}<h2 class="page-title">Inventory</h2>
    <div class="stats"><div class="card"><div class="stat-lbl">Stock value</div><div class="stat-val">${rupeesWhole(value)}</div></div>
    <div class="card"><div class="stat-lbl">Items to reorder</div><div class="stat-val">${low}</div></div></div>
    <div class="card" style="margin-top:12px">
      <h2>Log a purchase</h2>
      <label>Item</label><select id="lp-item"><option value="">Choose…</option>${mats.map((m) => `<option ${pre && m.name === pre.name ? "selected" : ""}>${escapeHtml(m.name)}</option>`).join("")}</select>
      <p class="tiny" id="lp-info">${pre ? `Pack: ${pre.packSize} ${escapeHtml(pre.unit)} at ${rupees(pre.packPrice)}` : "Adds to “Purchased” and, if the price changed, updates the pack price everywhere."}</p>
      <div class="row-3">
        <div><label>Packs bought</label><input id="lp-packs" type="number" min="0" step="any" value="1"></div>
        <div><label>Price per pack (₹)</label><input id="lp-price" type="number" min="0" step="any" value="${pre ? pre.packPrice : ""}"></div>
        <div><label>Date</label><input id="lp-date" type="date" value="${new Date().toISOString().slice(0, 10)}"></div>
      </div>
      <label>Note</label><input id="lp-note" placeholder="Shop, bill no. (optional)">
      <div class="btn-row"><button class="btn" id="lp-save">Add to stock</button></div>
      <p id="lp-msg" class="ok-msg"></p>
      ${recent.length ? `<h3>Recent purchases</h3><table class="breakdown"><tbody>${recent.map((p) => `<tr><td>${escapeHtml(p.item)}<div class="tiny">${fmtDate(p.date)} · ${p.packs} × ${rupees(p.pricePerPack)}</div></td><td class="num">+${p.qty} ${escapeHtml(p.unit || "")}<div class="tiny">${rupees(p.total)}</div></td></tr>`).join("")}</tbody></table>` : ""}
    </div>
    <div class="list" style="margin-top:12px">${items}</div>
    ${nav("more")}`;
}

function viewProfit() {
  const months = store.profitByMonth();
  const all = store.visibleOrders.filter((o) => o.status !== "Cancelled");
  const tot = all.reduce((s,o)=>({r:s.r+(o.quoteSnapshot?.finalQuote||0),c:s.c+(o.quoteSnapshot?.totalCost||0),p:s.p+(o.quoteSnapshot?.profit||0)}),{r:0,c:0,p:0});
  return `${topbar()}<h2 class="page-title">Profit</h2>
    <div class="stats">
      <div class="card"><div class="stat-lbl">All-time revenue</div><div class="stat-val">${rupeesWhole(tot.r)}</div></div>
      <div class="card"><div class="stat-lbl">All-time cost</div><div class="stat-val">${rupeesWhole(tot.c)}</div></div>
      <div class="card"><div class="stat-lbl">All-time profit</div><div class="stat-val">${rupeesWhole(tot.p)}</div></div>
      <div class="card"><div class="stat-lbl">Orders</div><div class="stat-val">${all.length}</div></div>
    </div>
    ${(() => { const cats = store.profitByCategory(); return cats.length > 1 || (cats[0] && cats[0].category !== "Cake") ? `<div class="card" style="margin-top:12px"><h2>By category</h2><table class="breakdown"><thead><tr><th>Category</th><th class="num">Orders</th><th class="num">Revenue</th><th class="num">Cost</th><th class="num">Profit</th></tr></thead><tbody>
        ${cats.map((m) => `<tr><td>${escapeHtml(m.category)}${m.qty ? `<div class="tiny">${m.qty} sold</div>` : ""}</td><td class="num">${m.orders}</td><td class="num">${rupees(m.revenue)}</td><td class="num">${rupees(m.cost)}</td><td class="num">${rupees(m.profit)}</td></tr>`).join("")}
      </tbody></table></div>` : ""; })()}
    <div class="card" style="margin-top:12px"><h2>By month</h2>
      ${months.length ? `<table class="breakdown"><thead><tr><th>Month</th><th class="num">Orders</th><th class="num">Revenue</th><th class="num">Cost</th><th class="num">Profit</th></tr></thead><tbody>
        ${months.map(m=>`<tr><td>${m.month}</td><td class="num">${m.orders}</td><td class="num">${rupees(m.revenue)}</td><td class="num">${rupees(m.cost)}</td><td class="num">${rupees(m.profit)}</td></tr>`).join("")}
      </tbody></table>` : `<p class="empty">Profit appears here once you save orders.</p>`}
    </div>
    <div class="card" style="margin-top:12px"><h2>Per order</h2>
      ${all.length ? `<table class="breakdown"><thead><tr><th>Order</th><th>Status</th><th class="num">Quote</th><th class="num">Cost</th><th class="num">Profit</th></tr></thead><tbody>
        ${all.map(o=>`<tr data-go="#/order/${o.id}" style="cursor:pointer"><td>${escapeHtml(o.customer||o.id)}<div class="tiny">${escapeHtml(cakeLabel(o))} · ${fmtDate(o.deliveryDate||o.createdAt)}</div></td><td><span class="pill ${o.status}">${o.status}</span></td><td class="num">${rupees(o.quoteSnapshot?.finalQuote||0)}</td><td class="num">${rupees(o.quoteSnapshot?.totalCost||0)}</td><td class="num">${rupees(o.quoteSnapshot?.profit||0)}</td></tr>`).join("")}
      </tbody></table>` : `<p class="empty">No orders yet.</p>`}
    </div>
    ${nav("more")}`;
}

function libraryPerKg(r) {
  return (r.ingredients || []).reduce((sum, l) => {
    const mat = store.state.rawMaterials.find((m) => m.name === l.ingredient);
    return sum + (mat ? (Number(l.qtyPerKg) || 0) * (Number(mat.packPrice) / Number(mat.packSize) || 0) : 0);
  }, 0);
}

function viewLibrary() {
  const cards = store.state.library.map((r) => `<div class="lib-item" data-go="#/library/${r.id}">
      <span class="pill">${escapeHtml(r.label)}</span>
      <h3 style="margin-top:6px">${escapeHtml(r.name)}</h3>
      <div class="meta">${escapeHtml(r.type)} · ${rupees(libraryPerKg(r))} / kg</div>
      <div class="tiny">${r.sourceUrl ? "Credit: "+escapeHtml(r.sourceName) : escapeHtml(r.sourceName)}</div>
    </div>`).join("");
  return `${topbar()}<h2 class="page-title">Recipe library</h2>
    <p class="muted">Adapted versions of the sources listed in her workbook. Methods are short paraphrases — not copies of the original text. Always credit the source when you share or teach a recipe.</p>
    <div class="list">${cards}</div>
    ${nav("more")}`;
}

function viewLibraryItem(id) {
  const r = store.state.library.find((x) => x.id === id);
  if (!r) return `${topbar()}<p>Not found.</p>${nav("more")}`;
  let lines = (r.ingredients || []).map((l) => {
    const mat = store.state.rawMaterials.find((m) => m.name === l.ingredient);
    const rate = mat ? Number(mat.packPrice)/Number(mat.packSize)||0 : 0;
    return { ...l, rate, cost: (l.qtyPerKg||0)*rate, estimated: !!(mat&&mat.estimated), missing: !mat };
  });
  const perKg = lines.reduce((s,l)=>s+l.cost,0);
  return `${topbar()}<p class="tiny"><a href="#/library">← Library</a></p>
    <h2 class="page-title">${escapeHtml(r.name)}</h2>
    <p><span class="pill">${escapeHtml(r.label)}</span> <span class="pill">${escapeHtml(r.type)}</span></p>
    <p>${/^https?:\/\//.test(r.sourceUrl||"") ? `Credit and original method: <a href="${escapeHtml(r.sourceUrl)}" target="_blank" rel="noopener">${escapeHtml(r.sourceName)}</a>` : escapeHtml(r.sourceName)}</p>
    <p class="tiny">${escapeHtml(r.sourceBatch||"")} ${escapeHtml(r.scaledHow||"")}</p>
    ${r.notes ? `<p class="tiny">${escapeHtml(r.notes)}</p>` : ""}
    <div class="card"><h2>Ingredients per 1 kg · ${rupees(perKg)}</h2>
      <table class="breakdown"><thead><tr><th>Ingredient</th><th class="num">Qty</th><th class="num">Cost</th></tr></thead><tbody>
        ${lines.map(l=>`<tr><td>${escapeHtml(l.ingredient)}${l.estimated?'<span class="tag-est">ESTIMATED — confirm</span>':""}${l.missing?'<span class="tag-est">MISSING</span>':""}</td><td class="num">${l.qtyPerKg} ${l.unit||""}</td><td class="num">${rupees(l.cost)}</td></tr>`).join("")}
      </tbody></table>
    </div>
    <div class="card" style="margin-top:12px"><h2>Method (short, adapted)</h2>
      <ol class="method">${(r.method||[]).map(s=>`<li>${escapeHtml(s)}</li>`).join("")}</ol>
    </div>
    ${nav("more")}`;
}

function viewMore() {
  const low = store.lowStock().length;
  return `${topbar()}<h2 class="page-title">More</h2>
    <div class="more-grid">
      <button class="more-tile" data-go="#/menu"><strong>Menu &amp; prices</strong><span>Categories, sizes, flavours, toppings, boxes, add-ons</span></button>
      <button class="more-tile" data-go="#/price-up"><strong>Price went up?</strong><span>Update an item's cost and see the effect</span></button>
      <button class="more-tile" data-go="#/customers"><strong>Customers</strong><span>Phone numbers and past orders</span></button>
      <button class="more-tile" data-go="#/pricing"><strong>Rules &amp; delivery</strong><span>Rush charges, discounts, delivery zones</span></button>
      <button class="more-tile" data-go="#/materials"><strong>Raw materials</strong><span>Pack prices that drive every quote</span></button>
      <button class="more-tile" data-go="#/inventory"><strong>Inventory ${low?`<span class="badge-count">${low}</span>`:""}</strong><span>Stock, used, reorder alerts</span></button>
      <button class="more-tile" data-go="#/profit"><strong>Profit</strong><span>Per order and per month</span></button>
      <button class="more-tile" data-go="#/library"><strong>Recipe library</strong><span>Adapted sources with methods</span></button>
      <button class="more-tile" data-go="#/settings"><strong>Settings & export</strong><span>PIN, backup, Google Sheet</span></button>
    </div>
    ${nav("more")}`;
}

function viewSettings() {
  const s = store.state.settings;
  const o = s.overheads || {};
  const st = engine?.status || { state: "local", conflicts: [] };
  const num = (id, label, v, step = "any") => `<div><label>${label}</label><input id="${id}" type="number" step="${step}" value="${v ?? 0}"></div>`;
  const conflicts = (st.conflicts || []).slice(-8).reverse();
  return `${topbar()}<h2 class="page-title">Settings</h2>
    <div class="card">
      <h2>Google Sheet sync</h2>
      <p class="tiny">Status: ${syncPill()}</p>
      <label>How should the app talk to the sheet?</label>
      <select id="sync-mode">
        ${[["local","Off: keep data on this phone only"],["appsscript","Apps Script web app (URL + token)"],["google","Sign in with Google (direct Sheets API)"]].map(([v,l])=>`<option value="${v}" ${syncCfg.mode===v?"selected":""}>${l}</option>`).join("")}
      </select>
      <div id="cfg-appsscript" ${syncCfg.mode==="appsscript"?"":"hidden"}>
        <label>Apps Script web app URL</label><input id="sync-url" placeholder="https://script.google.com/macros/s/…/exec" value="${escapeHtml(syncCfg.url)}">
        <label>Sync token (APP_TOKEN in the script)</label><input id="sync-token" type="password" value="${escapeHtml(syncCfg.token)}">
      </div>
      <div id="cfg-google" ${syncCfg.mode==="google"?"":"hidden"}>
        <label>Google OAuth client ID</label><input id="sync-client" placeholder="…apps.googleusercontent.com" value="${escapeHtml(syncCfg.clientId)}">
        <label>Spreadsheet ID</label><input id="sync-sheet" value="${escapeHtml(syncCfg.spreadsheetId || DEFAULT_SHEET_ID)}">
        <p class="tiny">${cachedGoogleToken() ? "Signed in to Google on this phone." : "Not signed in yet."}</p>
      </div>
      <div class="btn-row">
        <button class="btn" id="sync-save">Save &amp; connect</button>
        ${syncCfg.mode==="google" ? `<button class="btn secondary" id="sync-signin">Sign in with Google</button>` : ""}
        ${engine?.transport ? `<button class="btn secondary" id="sync-now">Reload from sheet</button>` : ""}
        ${syncCfg.mode==="google" && cachedGoogleToken() ? `<button class="btn ghost" id="sync-signout">Sign out</button>` : ""}
      </div>
      <p id="sync-msg" class="ok-msg">${escapeHtml(st.message || "")}</p>
      <p class="tiny">Data is read fresh from the sheet each time the app opens, and when you tap reload (↻). Every change you make here is written to the matching cells within a few seconds. Offline changes wait on this phone and are sent when you're back online.</p>
      ${conflicts.length ? `<h3>Sync notes</h3><ul class="tiny">${conflicts.map(c=>`<li>${escapeHtml(fmtTime(c.at))} · ${c.note ? "Note — " : "Not applied — "}${escapeHtml(c.what)}: ${escapeHtml(c.reason)}</li>`).join("")}</ul>` : ""}
      <p class="tiny"><a href="https://docs.google.com/spreadsheets/d/${escapeHtml(syncCfg.spreadsheetId || DEFAULT_SHEET_ID)}/edit" target="_blank" rel="noopener">Open Bakery Workbook</a> · <a href="./apps-script/SETUP.md.txt" target="_blank" rel="noopener">Setup steps</a> · <a href="./apps-script/Code.gs.txt" target="_blank" rel="noopener">Script to paste (Code.gs)</a></p>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Pricing rules</h2>
      <p class="tiny">Same as the Settings tab in the workbook${store.sheetMode ? "; saving writes the matching cells in column B" : ""}.</p>
      <div class="row-2">
        ${num("set-wastage","Wastage (%)", +(s.wastage*100).toFixed(4))}
        ${num("set-markup","Markup (%)", +(s.markup*100).toFixed(4))}
        ${num("set-gst","GST (%)", +(s.gst*100).toFixed(4))}
        ${num("set-round","Round up to (₹)", s.rounding)}
        ${num("set-labour","Labour ₹ / hour", s.labourRate)}
        ${num("set-basehrs","Base hours per cake", s.baseHours)}
        ${num("set-hrskg","Extra hours per kg", s.hoursPerKg)}
        ${num("set-hrslayer","Extra hours per layer", s.hoursPerLayer)}
        ${num("set-hrstop","Extra hours per topping", s.hoursPerTopping)}
        ${num("set-fill","Filling g / kg / gap", s.fillingGPerKg)}
        ${num("set-crumb","Crumb coat g / kg / layer", s.crumbCoatGPerKgPerLayer)}
        ${num("set-frost","Frosting g / kg / layer", s.finalFrostingGPerKgPerLayer)}
        ${num("set-delrate","Delivery ₹ / km", s.deliveryRatePerKm)}
        ${num("set-trip","Trip multiplier", s.tripMultiplier)}
        ${num("set-mindel","Min delivery (₹)", s.minDelivery)}
      </div>
      <h3>Fixed overheads per cake (₹)</h3>
      <div class="row-2">
        ${num("oh-gas","Gas / LPG", o.gas)}${num("oh-electricity","Electricity", o.electricity)}
        ${num("oh-water","Water & cleaning", o.water)}${num("oh-rent","Rent / kitchen share", o.rent)}
        ${num("oh-equipment","Equipment & misc.", o.equipment)}
      </div>
      <p class="tiny">Cake sizes, layers, flavours, toppings, boxes, boards and add-ons: <a href="#/menu">Menu &amp; prices</a>. Rush charges, discounts and delivery zones: <a href="#/pricing">Rules &amp; delivery</a>.</p>
      <div class="btn-row"><button class="btn" id="save-settings">Save rules</button></div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>PIN</h2>
      <p class="tiny">Default PIN is <code>2468</code>. Change it here. It is stored only on this phone (hashed).</p>
      <label>New PIN</label><input id="new-pin" type="password" inputmode="numeric" maxlength="8">
      <div class="btn-row"><button class="btn secondary" id="change-pin">Change PIN</button><button class="btn ghost" id="lock-now">Lock now</button></div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Backup</h2>
      <div class="btn-row">
        <button class="btn" id="export-json">Download backup (JSON)</button>
        <button class="btn secondary" id="export-orders">Download orders CSV</button>
        ${store.sheetMode ? "" : `<button class="btn secondary" id="import-json">Import backup</button>`}
        <input id="import-file" type="file" accept="application/json,.json" hidden>
      </div>
      <p id="io-msg" class="ok-msg"></p>
      ${store.sheetMode ? "" : `<div class="btn-row"><button class="btn danger" id="reset-data">Reset data to the original workbook</button></div>`}
    </div>
    ${nav("more")}`;
}

// ================= v3 screens: menu & prices, price went up, new recipe, rules & delivery, customers =================
function roomNote(section, type) {
  const c = store.roomLeft(section, type);
  if (!c) return "";
  if (c.free > 0) return `<p class="tiny">${c.free} empty row${c.free === 1 ? "" : "s"} left for this list in the sheet (${escapeHtml(c.range)}).</p>`;
  return `<p class="tiny warn-msg">This list is full in the sheet (${escapeHtml(c.range)}). <button class="mini-btn" data-room-sec="${section}" data-room-type="${escapeHtml(type || "")}">Make room for 5 more</button></p>`;
}
function sellInput(kind, item) {
  const v = store.sellingPrice(kind, item);
  return `<div><label>Selling price (₹)</label><input type="number" min="0" step="1" data-sell-kind="${kind}" data-sell-item="${escapeHtml(item)}" value="${v ?? ""}" placeholder="—"></div>`;
}
function matCost(name) {
  const m = store.state.rawMaterials.find((x) => x.name === name);
  if (!m) return `<span class="tag-est">no Raw Materials row</span>`;
  const r = Number(m.packPrice) / Number(m.packSize) || 0;
  return `${rupees(r)} / ${escapeHtml(m.unit)} <a href="#/price-up?item=${encodeURIComponent(name)}">change cost</a>`;
}
function delBtn(section, key) {
  return `<button class="x-btn" title="Remove" aria-label="Remove" data-menu-del="${section}" data-key="${escapeHtml(String(key))}">✕</button>`;
}
function matOptions(sel, withNew = true) {
  const names = store.state.rawMaterials.map((m) => m.name);
  return `${withNew ? `<option value="__new">+ New item…</option>` : ""}${names.map((n) => `<option ${n === sel ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}`;
}
function newMatFields(prefix, unit = "g") {
  return `<div class="row-3 newmat" id="${prefix}-newmat" hidden>
    <div><label>Pack size</label><input id="${prefix}-psize" type="number" min="0" step="any" value="1"></div>
    <div><label>Unit</label><input id="${prefix}-punit" value="${unit}"></div>
    <div><label>Pack price (₹)</label><input id="${prefix}-pprice" type="number" min="0" step="any" value="0"></div>
  </div>`;
}

function viewMenu() {
  const s = store.state.settings;
  const costs = store.recipeCosts();
  const card = (title, body, section, type) => `<div class="card" style="margin-top:12px"><h2>${title}</h2>${roomNote(section, type)}${body}</div>`;
  const sizes = s.sizes.map((z) => `<div class="mrow">
      <div class="row-3">
        <div><label>Size name</label><input data-menu-sec="sizes" data-key="${escapeHtml(z.name)}" data-field="name" value="${escapeHtml(z.name)}"></div>
        <div><label>Weight (kg)</label><input type="number" min="0" step="any" data-menu-sec="sizes" data-key="${escapeHtml(z.name)}" data-field="kg" value="${z.kg}"></div>
        <div><label>Servings</label><input data-menu-sec="sizes" data-key="${escapeHtml(z.name)}" data-field="servings" value="${escapeHtml(z.servings)}"></div>
        ${sellInput("Size", z.name)}
      </div>${delBtn("sizes", z.name)}</div>`).join("");
  const addSize = `<div class="addbox"><div class="row-3">
      <div><label>New size</label><input id="add-sizes-name" placeholder="e.g. 2.5 kg"></div>
      <div><label>Weight (kg)</label><input id="add-sizes-kg" type="number" min="0" step="any"></div>
      <div><label>Servings</label><input id="add-sizes-serv" placeholder="e.g. 20-25"></div>
      <div><label>Selling price (₹)</label><input id="add-sizes-sell" type="number" min="0" step="1" placeholder="optional"></div>
    </div><div class="btn-row"><button class="btn secondary" data-menu-add="sizes">Add size</button></div></div>`;
  const layers = `<div class="chips">${s.layers.map((l) => `<span class="chip on">${l} layers <button class="x-btn sm" data-menu-del="layers" data-key="${l}" aria-label="Remove">✕</button></span>`).join("")}</div>
    <div class="row-2"><div><label>Add layer option</label><input id="add-layers-value" type="number" min="1" step="1" placeholder="e.g. 6"></div><div><label>&nbsp;</label><button class="btn secondary full" data-menu-add="layers">Add</button></div></div>`;
  const onMenu = new Set(s.recipeMenu.map((m) => m.name));
  const recipeCard = (type, plural) => {
    const rows = s.recipeMenu.filter((m) => m.type === type).map((m) => `<div class="mrow">
        <div class="row-2">
          <div><label>Name</label><input data-menu-sec="recipeMenu" data-key="${escapeHtml(m.name)}" data-field="name" value="${escapeHtml(m.name)}"></div>
          <div><label>Cost per kg</label><div class="meta" style="padding:10px 0">${rupees(costs[m.name] || 0)} · <a href="#/recipes?recipe=${encodeURIComponent(m.name)}">recipe</a></div></div>
        </div>${delBtn("recipeMenu", m.name)}</div>`).join("");
    const orphan = [...new Set(store.state.recipes.filter((r) => r.type === type && !onMenu.has(r.name)).map((r) => r.name))];
    return card(plural, `${rows || `<p class="empty">None yet.</p>`}
      ${orphan.length ? `<p class="tiny">In Recipes but not on the menu: ${orphan.map((n) => `<button class="mini-btn" data-menu-orphan="${escapeHtml(n)}" data-type="${type}">+ ${escapeHtml(n)}</button>`).join(" ")}</p>` : ""}
      <div class="btn-row"><button class="btn secondary" data-go="#/recipe-new?type=${type}">New ${type.toLowerCase()} recipe</button></div>`, "recipeMenu", type);
  };
  const toppings = s.toppings.filter((t) => t.name !== "None").map((t) => `<div class="mrow">
      <div class="row-3">
        <div><label>Topping</label><input data-menu-sec="toppings" data-key="${escapeHtml(t.name)}" data-field="name" value="${escapeHtml(t.name)}"></div>
        <div><label>Uses item</label><select data-menu-sec="toppings" data-key="${escapeHtml(t.name)}" data-field="material">${matOptions(t.material, false)}</select></div>
        <div><label>Fixed qty / cake</label><input type="number" min="0" step="any" data-menu-sec="toppings" data-key="${escapeHtml(t.name)}" data-field="fixedQty" value="${t.fixedQty}"></div>
        <div><label>Qty per kg</label><input type="number" min="0" step="any" data-menu-sec="toppings" data-key="${escapeHtml(t.name)}" data-field="qtyPerKg" value="${t.qtyPerKg}"></div>
        ${sellInput("Topping", t.name)}
        <div><label>Cost</label><div class="meta" style="padding:10px 0">${matCost(t.material)}</div></div>
      </div>${delBtn("toppings", t.name)}</div>`).join("");
  const addTopping = `<div class="addbox"><div class="row-3">
      <div><label>New topping</label><input id="add-toppings-name" placeholder="e.g. Gold sprinkles"></div>
      <div><label>Uses item</label><select id="add-toppings-mat" data-newmat="add-toppings">${matOptions("")}</select></div>
      <div><label>Fixed qty / cake</label><input id="add-toppings-fixed" type="number" min="0" step="any" value="0"></div>
      <div><label>Qty per kg</label><input id="add-toppings-perkg" type="number" min="0" step="any" value="0"></div>
      <div><label>Selling price (₹)</label><input id="add-toppings-sell" type="number" min="0" step="1" placeholder="optional"></div>
    </div>${newMatFields("add-toppings", "g")}<div class="btn-row"><button class="btn secondary" data-menu-add="toppings">Add topping</button></div></div>`;
  const pack = (section, kind, label) => {
    const rows = s[section].map((b) => `<div class="mrow"><div class="row-3">
        <div><label>${label}</label><input data-menu-sec="${section}" data-key="${escapeHtml(b)}" data-field="name" value="${escapeHtml(b)}"></div>
        <div><label>Cost each</label><div class="meta" style="padding:10px 0">${matCost(b)}</div></div>
        ${sellInput(kind, b)}
      </div>${delBtn(section, b)}</div>`).join("");
    const add = `<div class="addbox"><div class="row-3">
        <div><label>New ${label.toLowerCase()}</label><input id="add-${section}-name"></div>
        <div><label>Pack size (nos)</label><input id="add-${section}-psize" type="number" min="1" step="1" value="10"></div>
        <div><label>Pack price (₹)</label><input id="add-${section}-pprice" type="number" min="0" step="any" value="0"></div>
        <div><label>Selling price (₹)</label><input id="add-${section}-sell" type="number" min="0" step="1" placeholder="optional"></div>
      </div><p class="tiny">Its cost comes from a Raw Materials row with the same name (created for you if it doesn't exist).</p>
      <div class="btn-row"><button class="btn secondary" data-menu-add="${section}">Add ${label.toLowerCase()}</button></div></div>`;
    return card(label === "Cake box" ? "Cake boxes" : "Cake boards", rows + add, section);
  };
  const addons = s.addons.map((a) => `<div class="mrow"><div class="row-3">
      <div><label>Add-on</label><input data-menu-sec="addons" data-key="${escapeHtml(a.name)}" data-field="name" value="${escapeHtml(a.name)}"></div>
      <div><label>Uses item</label><select data-menu-sec="addons" data-key="${escapeHtml(a.name)}" data-field="material">${matOptions(a.material, false)}</select></div>
      <div><label>Qty per cake</label><input type="number" min="0" step="any" data-menu-sec="addons" data-key="${escapeHtml(a.name)}" data-field="qty" value="${a.qty}"></div>
      ${sellInput("Add-on", a.name)}
      <div><label>Cost</label><div class="meta" style="padding:10px 0">${matCost(a.material)}</div></div>
    </div>${delBtn("addons", a.name)}</div>`).join("");
  const addAddon = `<div class="addbox"><div class="row-3">
      <div><label>New add-on</label><input id="add-addons-name" placeholder="e.g. Sparkler candle"></div>
      <div><label>Uses item</label><select id="add-addons-mat" data-newmat="add-addons">${matOptions("")}</select></div>
      <div><label>Qty per cake</label><input id="add-addons-qty" type="number" min="0" step="any" value="1"></div>
      <div><label>Selling price (₹)</label><input id="add-addons-sell" type="number" min="0" step="1" placeholder="optional"></div>
    </div>${newMatFields("add-addons", "nos")}<div class="btn-row"><button class="btn secondary" data-menu-add="addons">Add add-on</button></div></div>`;
  const catRows = store.categories.map((c) => `<div class="mrow cat-row">
      <div class="row-3">
        <div><label>Category</label><input data-cat="${escapeHtml(c.name)}" data-field="name" value="${escapeHtml(c.name)}"></div>
        <div><label>Pricing</label><select data-cat="${escapeHtml(c.name)}" data-field="pricing">${["Calculator", "Per unit"].map((x) => `<option value="${x}" ${c.pricing === x ? "selected" : ""}>${x === "Calculator" ? "Cake calculator" : "Per unit"}</option>`).join("")}</select></div>
        ${c.pricing === "Calculator" ? `<div><label>Price</label><div class="meta" style="padding:10px 0">From sizes &amp; recipes below</div></div>` : `
        <div><label>Price per unit (₹)</label><input type="number" min="0" step="1" data-cat="${escapeHtml(c.name)}" data-field="unitPrice" value="${c.unitPrice}"></div>
        <div><label>Unit</label><input data-cat="${escapeHtml(c.name)}" data-field="unit" value="${escapeHtml(c.unit)}" placeholder="bar, cookie, box"></div>
        <div><label>Cost per unit (₹)</label><input type="number" min="0" step="any" data-cat="${escapeHtml(c.name)}" data-field="costPerUnit" value="${c.costPerUnit || ""}" placeholder="optional"></div>`}
        <div><label>On the order screen</label><label class="check" style="margin:8px 0 0"><input type="checkbox" data-cat="${escapeHtml(c.name)}" data-field="active" ${c.active !== false ? "checked" : ""}> Show</label></div>
      </div><button class="x-btn" data-cat-del="${escapeHtml(c.name)}" aria-label="Remove ${escapeHtml(c.name)}">✕</button></div>`).join("");
  const addCat = `<div class="addbox"><div class="row-3">
      <div><label>New category</label><input id="add-cat-name" placeholder="e.g. Brownies"></div>
      <div><label>Price per unit (₹)</label><input id="add-cat-price" type="number" min="0" step="1" placeholder="e.g. 90"></div>
      <div><label>Unit</label><input id="add-cat-unit" placeholder="e.g. brownie"></div>
      <div><label>Cost per unit (₹)</label><input id="add-cat-cost" type="number" min="0" step="any" placeholder="optional"></div>
    </div><div class="btn-row"><button class="btn secondary" id="add-cat">Add category</button></div></div>`;
  return `${topbar()}<h2 class="page-title">Menu &amp; prices</h2>
    <p class="muted">Everything the order form offers. Changes are saved${store.sheetMode ? " to the Settings and Price List tabs of her sheet" : " on this phone"} and show up in the order form straight away.</p>
    <p class="tiny"><strong>Selling price</strong> is optional. When a size has one, quotes use her menu price (size price + each extra’s selling price, or its cost + markup when it has none). Leave sizes blank to keep the cost-plus calculator.</p>
    <p id="menu-msg" class="ok-msg"></p>
    <div class="card" style="margin-top:12px"><h2>Categories</h2>
      <p class="tiny">What she sells. Each one gets its own screen under New order. “Per unit” = quantity × price per unit; “Cake calculator” uses the sizes and recipes below.${store.sheetMode ? " Saved to the Categories tab." : ""}</p>
      ${catRows}${addCat}</div>
    ${card("Cake sizes", sizes + addSize, "sizes")}
    ${card("Layers", layers, "layers")}
    ${recipeCard("Sponge", "Sponges")}
    ${recipeCard("Filling", "Fillings")}
    ${recipeCard("Frosting", "Frostings")}
    ${card("Toppings", toppings + addTopping, "toppings")}
    ${pack("boxes", "Box", "Cake box")}
    ${pack("boards", "Board", "Cake board")}
    ${card("Add-ons", addons + addAddon, "addons")}
    ${nav("more")}`;
}

function viewPriceUp() {
  const mats = store.state.rawMaterials;
  const name = route.item && mats.some((m) => m.name === route.item) ? route.item : "";
  const m = mats.find((x) => x.name === name);
  return `${topbar()}<h2 class="page-title">Price went up?</h2>
    <p class="muted">Pick the item, type the new pack price, check what changes, then save. Every recipe cost and new quote uses the new price.</p>
    <div class="card">
      <label>Item</label>
      <select id="pu-item"><option value="">Choose…</option>${mats.map((x) => `<option ${x.name === name ? "selected" : ""}>${escapeHtml(x.name)}</option>`).join("")}</select>
      ${m ? `<p class="meta">Now: ${rupees(m.packPrice)} for ${m.packSize} ${escapeHtml(m.unit)} (${rupees(Number(m.packPrice) / Number(m.packSize) || 0, 4)} / ${escapeHtml(m.unit)})${m.estimated ? ' <span class="tag-est">ESTIMATED</span>' : ""}</p>
      <div class="row-2">
        <div><label>New pack price (₹)</label><input id="pu-price" type="number" min="0" step="any" value="${m.packPrice}"></div>
        <div><label>Pack size (${escapeHtml(m.unit)})</label><input id="pu-size" type="number" min="0" step="any" value="${m.packSize}"></div>
      </div>
      <div id="pu-impact"></div>
      <div class="btn-row"><button class="btn" id="pu-save">Save new price</button></div>
      <p id="pu-msg" class="ok-msg"></p>` : ""}
    </div>
    ${nav("more")}`;
}
function renderImpact() {
  const name = $("#pu-item")?.value;
  const box = $("#pu-impact");
  if (!name || !box) return;
  const im = store.priceImpact(name, $("#pu-price").value, $("#pu-size").value);
  if (!im) return;
  const arrow = (a, b) => `${rupees(a)} → <strong>${rupees(b)}</strong>`;
  box.innerHTML = `<div class="impact">
    <p>Cost per ${escapeHtml(im.item.unit)}: ${arrow(im.costPerUnitBefore, im.costPerUnitAfter)}</p>
    ${im.recipes.length ? `<table class="breakdown"><thead><tr><th>Recipe</th><th class="num">Cost per kg</th></tr></thead><tbody>${im.recipes.map((r) => `<tr><td>${escapeHtml(r.name)}</td><td class="num">${arrow(r.before, r.after)}</td></tr>`).join("")}</tbody></table>` : `<p class="tiny">Not used in any recipe.</p>`}
    ${im.uses.length ? `<p class="tiny">Also used by: ${im.uses.map(escapeHtml).join(", ")}</p>` : ""}
    ${im.sampleQuote ? `<p class="tiny">Sample 0.5 kg chocolate cake quote: ${rupeesWhole(im.sampleQuote.before)} → <strong>${rupeesWhole(im.sampleQuote.after)}</strong></p>` : ""}
    ${im.openOrders ? `<p class="tiny warn-msg">${im.openOrders} open order${im.openOrders > 1 ? "s" : ""} keep their saved quote. Open one and tap “Update order” to re-price it.</p>` : ""}
  </div>`;
}

function viewRecipeNew() {
  const type = ["Sponge", "Filling", "Frosting"].includes(route.type) ? route.type : "Sponge";
  const mats = store.state.rawMaterials.map((m) => m.name);
  const line = (i) => `<div class="row-2 rn-line"><div><select class="rn-ing"><option value="">Ingredient ${i + 1}…</option>${mats.map((n) => `<option>${escapeHtml(n)}</option>`).join("")}</select></div><div><input class="rn-qty" type="number" min="0" step="any" placeholder="Qty for 1 kg"></div></div>`;
  return `${topbar()}<p class="tiny"><a href="#/recipes">← Recipes</a></p><h2 class="page-title">New recipe</h2>
    <div class="card">
      <div class="row-2">
        <div><label>Recipe name</label><input id="rn-name" placeholder="e.g. Butterscotch Sponge"></div>
        <div><label>Type</label><select id="rn-type">${options(["Sponge", "Filling", "Frosting"], type)}</select></div>
      </div>
      ${roomNote("recipeMenu", type)}
      <label>Ingredients for 1 kg of this recipe (in each item’s own unit: g, ml or nos)</label>
      <div id="rn-lines">${[0, 1, 2, 3, 4].map(line).join("")}</div>
      <div class="btn-row"><button class="btn ghost" id="rn-more">+ another ingredient</button></div>
      <label>Method notes</label><textarea id="rn-method" placeholder="Short steps, oven temperature, baking time…"></textarea>
      <div class="btn-row"><button class="btn" id="rn-save">Create recipe</button></div>
      <p id="rn-msg" class="error"></p>
      <p class="tiny">Creates the recipe lines in the Recipes tab, adds it to the ${type.toLowerCase()} menu in Settings, and saves the notes in a “Recipe Notes” tab.</p>
    </div>
    ${nav("recipes")}`;
}

function viewPricing() {
  const rules = store.state.rules || [];
  const zones = store.state.zones || [];
  const s = store.state.settings;
  const ruleRow = (r) => `<div class="mrow" data-rule-row="${r.id}">
      <div class="row-3">
        <div><label>Rule</label><input data-rule="${r.id}" data-field="name" value="${escapeHtml(r.name)}"></div>
        <div><label>Effect</label><select data-rule="${r.id}" data-field="effect">${options(["Charge", "Discount"], r.effect)}</select></div>
        <div><label>Amount</label><div class="amt"><input type="number" min="0" step="any" data-rule="${r.id}" data-field="amount" value="${r.amount}"><select data-rule="${r.id}" data-field="type">${options(["%", "₹"], r.type)}</select></div></div>
        <div><label>Apply</label><select data-rule="${r.id}" data-field="applyTo">${options(["When ticked", "Every order"], r.applyTo)}</select></div>
      </div>
      <label class="check"><input type="checkbox" data-rule="${r.id}" data-field="active" ${r.active ? "checked" : ""}> On</label>
      <button class="x-btn" data-rule-del="${r.id}" aria-label="Remove">✕</button></div>`;
  const zoneRow = (z) => `<div class="mrow"><div class="row-2">
      <div><label>Zone</label><input data-zone="${z.id}" data-field="name" value="${escapeHtml(z.name)}"></div>
      <div><label>Flat charge (₹)</label><input type="number" min="0" step="1" data-zone="${z.id}" data-field="charge" value="${z.charge}"></div>
    </div><button class="x-btn" data-zone-del="${z.id}" aria-label="Remove">✕</button></div>`;
  return `${topbar()}<h2 class="page-title">Pricing rules &amp; delivery</h2>
    <div class="card">
      <h2>Rush charges &amp; discounts</h2>
      <p class="tiny">Switch a rule On to offer it. “When ticked” rules show as tick boxes on the order form (e.g. express orders); “Every order” rules are ticked automatically on new orders (e.g. a Diwali discount). % is worked out on the cake price before delivery.</p>
      ${rules.map(ruleRow).join("") || `<p class="empty">No rules yet.</p>`}
      <div class="btn-row">
        <button class="btn secondary" id="rule-add-rush">+ Rush / express charge</button>
        <button class="btn secondary" id="rule-add-fest">+ Festival discount</button>
      </div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Delivery zones</h2>
      <p class="tiny">A flat delivery charge per area. On the order form pick a zone, or leave it on “By distance” to use ₹${s.deliveryRatePerKm}/km × ${s.tripMultiplier} (minimum ₹${s.minDelivery}). Per-km rates are under Settings.</p>
      ${zones.map(zoneRow).join("") || `<p class="empty">No zones yet.</p>`}
      <div class="row-2"><div><label>New zone</label><input id="zone-name" placeholder="e.g. Within 5 km"></div><div><label>Flat charge (₹)</label><input id="zone-charge" type="number" min="0" step="1"></div></div>
      <div class="btn-row"><button class="btn secondary" id="zone-add">Add zone</button></div>
    </div>
    <p class="tiny">Saved${store.sheetMode ? " in the “Pricing Rules” and “Delivery Zones” tabs of her sheet (created the first time you add one)" : " on this phone"}.</p>
    ${nav("more")}`;
}

function custKey(c) {
  return c.id || "n:" + encodeURIComponent((c.name || "") + "|" + (c.phone || ""));
}
function viewCustomers() {
  const q = (route.q || "").toLowerCase();
  let rows = store.customerRows();
  if (q) rows = rows.filter((c) => (c.name + " " + c.phone).toLowerCase().includes(q));
  return `${topbar()}<h2 class="page-title">Customers</h2>
    <input class="search" id="cust-search" placeholder="Search name or phone" value="${escapeHtml(route.q || "")}">
    <div class="list">${rows.map((c) => `<div class="order-item" data-go="#/customer/${encodeURIComponent(custKey(c))}">
        <div><h3>${escapeHtml(c.name)}</h3><div class="meta">${escapeHtml(c.phone || "no phone")}${c.last ? " · last " + fmtDate(c.last) : ""}</div></div>
        <div style="text-align:right"><span class="pill">${c.orders} order${c.orders === 1 ? "" : "s"}</span><div class="meta" style="margin-top:4px">${rupeesWhole(c.spent)}</div></div>
      </div>`).join("") || `<p class="empty">No customers yet. They're added automatically when you save an order.</p>`}</div>
    <div class="card" style="margin-top:12px"><h2>Add customer</h2>
      <div class="row-2"><div><label>Name</label><input id="nc-name"></div><div><label>Phone</label><input id="nc-phone" type="tel"></div></div>
      <label>Notes</label><input id="nc-notes" placeholder="Allergies, favourite flavour…">
      <div class="btn-row"><button class="btn" id="nc-save">Save customer</button></div>
    </div>
    ${nav("more")}`;
}
function viewCustomer(key) {
  const c = store.customerRows().find((x) => custKey(x) === key);
  if (!c) return `${topbar()}<p>Customer not found.</p>${nav("more")}`;
  const orders = store.ordersOf(c);
  return `${topbar()}<p class="tiny"><a href="#/customers">← Customers</a></p>
    <h2 class="page-title">${escapeHtml(c.name)}</h2>
    <div class="card">
      <div class="row-2"><div><label>Name</label><input id="cu-name" value="${escapeHtml(c.name)}"></div><div><label>Phone</label><input id="cu-phone" type="tel" value="${escapeHtml(c.phone || "")}"></div></div>
      <label>Notes</label><input id="cu-notes" value="${escapeHtml(c.notes || "")}">
      <div class="btn-row">
        <button class="btn" id="cu-save" data-id="${escapeHtml(c.id || "")}">${c.saved ? "Save" : "Save to customer list"}</button>
        <button class="btn secondary" data-go="#/order/new?cust=${encodeURIComponent(key)}">New order for ${escapeHtml(c.name.split(" ")[0])}</button>
        ${c.phone ? `<a class="btn ghost" href="https://wa.me/${escapeHtml(String(c.phone).replace(/\D/g, "").replace(/^(\d{10})$/, "91$1"))}" target="_blank" rel="noopener">WhatsApp</a>` : ""}
        ${c.saved ? `<button class="btn ghost" id="cu-del" data-id="${escapeHtml(c.id)}">Remove</button>` : ""}
      </div>
    </div>
    <div class="card" style="margin-top:12px"><h2>Past orders · ${orders.length}</h2>
      <div class="list">${orders.map((o) => `<div class="order-item" data-go="#/order/${o.id}"><div><h3>${escapeHtml(cakeLabel(o))}</h3><div class="meta">${fmtDate(o.deliveryDate || o.createdAt)} · ${o.id}</div></div><div style="text-align:right"><span class="pill ${o.status}">${o.status}</span><div class="meta" style="margin-top:4px">${rupeesWhole(o.quoteSnapshot?.finalQuote || 0)}</div></div></div>`).join("") || `<p class="empty">No orders yet.</p>`}</div>
    </div>
    ${nav("more")}`;
}

function viewLogin() {
  return `<div class="login-wrap"><div class="login-card">
    <img class="logo" src="./icons/icon-192.png" alt="">
    <h1>Rise & Whisk</h1>
    <p class="sub">Her bakery book — quotes, orders and stock</p>
    <div class="pin-row">
      <input class="pin-d" inputmode="numeric" maxlength="1" data-i="0" autocomplete="off">
      <input class="pin-d" inputmode="numeric" maxlength="1" data-i="1" autocomplete="off">
      <input class="pin-d" inputmode="numeric" maxlength="1" data-i="2" autocomplete="off">
      <input class="pin-d" inputmode="numeric" maxlength="1" data-i="3" autocomplete="off">
    </div>
    <p class="error" id="pin-err"></p>
    <button class="btn full" id="pin-go">Unlock</button>
    <p class="hint">Default PIN is 2468. Change it in Settings after you sign in.</p>
  </div></div>`;
}

function parseHash() {
  const h = (location.hash || "#/dashboard").replace(/^#/, "");
  const [path, qs] = h.split("?");
  const parts = path.split("/").filter(Boolean);
  const qp = new URLSearchParams(qs || "");
  if (parts[0] === "order" && parts[1] === "new") return { name: "order-new", cat: parts[2] ? decodeURIComponent(parts.slice(2).join("/")) : "", cust: qp.get("cust") || "" };
  if (parts[0] === "customer" && parts[1]) return { name: "customer", id: decodeURIComponent(parts.slice(1).join("/")) };
  if (["price-up", "recipe-new", "inventory", "customers"].includes(parts[0])) return { name: parts[0], item: qp.get("item") || "", type: qp.get("type") || "", q: qp.get("q") || "" };
  if (parts[0] === "order" && parts[1]) return { name: "order", id: parts[1] };
  if (parts[0] === "library" && parts[1]) return { name: "library-item", id: parts[1] };
  return { name: parts[0] || "dashboard", q: new URLSearchParams(qs || "").get("q") || "", status: new URLSearchParams(qs || "").get("status") || "All", recipe: new URLSearchParams(qs || "").get("recipe") || "", kg: new URLSearchParams(qs || "").get("kg") || "" };
}

function render() {
  const root = $("#app");
  if (!isUnlocked()) {
    root.innerHTML = viewLogin();
    bindLogin();
    return;
  }
  route = parseHash();
  const views = {
    dashboard: viewDashboard,
    orders: viewOrders,
    "order-new": viewOrderNew,
    order: () => viewOrder(route.id),
    calculator: viewCalculator,
    recipes: viewRecipes,
    materials: viewMaterials,
    inventory: viewInventory,
    profit: viewProfit,
    library: viewLibrary,
    "library-item": () => viewLibraryItem(route.id),
    more: viewMore,
    settings: viewSettings,
    menu: viewMenu,
    "price-up": viewPriceUp,
    "recipe-new": viewRecipeNew,
    pricing: viewPricing,
    customers: viewCustomers,
    customer: () => viewCustomer(route.id),
  };
  const fn = views[route.name] || viewDashboard;
  root.innerHTML = fn();
  bindPage();
}

function bindLogin() {
  const digits = [...document.querySelectorAll(".pin-d")];
  digits[0]?.focus();
  digits.forEach((el, i) => {
    el.addEventListener("input", () => {
      el.value = el.value.replace(/\D/g, "").slice(-1);
      if (el.value && digits[i + 1]) digits[i + 1].focus();
      if (i === 3 && el.value) tryPin();
    });
    el.addEventListener("keydown", (e) => {
      if (e.key === "Backspace" && !el.value && digits[i - 1]) digits[i - 1].focus();
    });
  });
  $("#pin-go")?.addEventListener("click", tryPin);
  async function tryPin() {
    const pin = digits.map((d) => d.value).join("");
    const ok = await checkPin(pin);
    if (!ok) {
      $("#pin-err").textContent = "That PIN doesn’t match.";
      digits.forEach((d) => (d.value = ""));
      digits[0].focus();
      return;
    }
    unlockSession();
    render();
  }
}

function hookQuoteForm() {
  const uform = $("#unit-form");
  if (uform) {
    const upd = () => renderUnitQuote(uform);
    uform.addEventListener("input", upd);
    uform.addEventListener("change", upd);
    upd();
  }
  const form = $("#cake-form");
  if (!form) return;
  const update = () => renderLiveQuote(form);
  form.addEventListener("input", update);
  form.addEventListener("change", update);
  // size: presets or a manual size, shown in grams or kg
  const sizeSel = form.elements.size, box = $("#custom-size"), amt = form.elements.customSize;
  sizeSel?.addEventListener("change", () => {
    box.hidden = sizeSel.value !== "__custom";
    if (!box.hidden) {
      if (!amt.value) { const prev = sizeSel.dataset.prevKg; if (prev) amt.value = form.elements.sizeUnit.value === "g" ? Math.round(prev * 1000) : prev; }
      amt.focus();
    } else { sizeSel.dataset.prevKg = sizeSel.selectedOptions[0]?.dataset.kg || ""; amt.value = ""; }
    update();
  });
  sizeSel && (sizeSel.dataset.prevKg = sizeSel.selectedOptions[0]?.dataset.kg || "");
  $("#size-unit")?.addEventListener("click", (e) => {
    const b = e.target.closest("[data-unit]");
    if (!b) return;
    e.preventDefault();
    const to = b.dataset.unit, from = form.elements.sizeUnit.value;
    if (to === from) return;
    localStorage.setItem(SIZE_UNIT_KEY, to);
    form.elements.sizeUnit.value = to;
    document.querySelectorAll("#size-unit [data-unit]").forEach((x) => x.classList.toggle("on", x.dataset.unit === to));
    for (const o of sizeSel.options) {
      if (o.value === "__custom") continue;
      const kg = Number(o.dataset.kg);
      o.textContent = to === "g" && kg ? sizeLabel(kg, "g") : o.value;
    }
    if (amt.value) amt.value = String(Math.round((to === "g" ? Number(amt.value) * 1000 : Number(amt.value) / 1000) * 1000) / 1000);
    amt.placeholder = to === "g" ? "e.g. 750" : "e.g. 1.2";
    $("#custom-size .unit-tag").textContent = to;
    update();
  });
  update();
}

function bindPage() {
  document.querySelectorAll("[data-go]").forEach((el) => {
    el.addEventListener("click", (e) => {
      const go = el.getAttribute("data-go");
      if (go) { e.preventDefault(); location.hash = go; }
    });
  });
  hookQuoteForm();

  $("#order-search")?.addEventListener("input", (e) => {
    const status = route.status || "All";
    location.hash = `#/orders?q=${encodeURIComponent(e.target.value)}&status=${encodeURIComponent(status)}`;
  });
  $("#status-filters")?.addEventListener("click", (e) => {
    const s = e.target.dataset.status;
    if (!s) return;
    location.hash = `#/orders?q=${encodeURIComponent(route.q||"")}&status=${encodeURIComponent(s)}`;
  });

  $("#save-order")?.addEventListener("click", () => {
    const form = $("#cake-form");
    const { input, q } = renderLiveQuote(form);
    const id = $("#save-order").dataset.id;
    const payload = { ...input, quoteSnapshot: q };
    if (id) store.updateOrder(id, { ...payload, recalculate: true });
    else store.addOrder(payload);
    location.hash = "#/orders";
  });
  $("#save-unit-order")?.addEventListener("click", () => {
    const form = $("#unit-form");
    const { input, q } = renderUnitQuote(form);
    if (!(input.quantity > 0)) { alert("Enter how many."); return; }
    if (!String(input.customer || "").trim()) { alert("Enter the customer's name."); return; }
    const id = $("#save-unit-order").dataset.id;
    try {
      if (id) store.updateOrder(id, { ...input, quoteSnapshot: q });
      else store.addOrder(input);
    } catch (err) { alert(err.message || String(err)); return; }
    location.hash = "#/orders";
  });
  $("#save-quote-order")?.addEventListener("click", () => {
    const form = $("#cake-form");
    const { input, q } = renderLiveQuote(form);
    if (!input.customer) input.customer = "Quote";
    store.addOrder({ ...input, status: input.status || "Enquiry", quoteSnapshot: q });
    location.hash = "#/orders";
  });
  $("#load-sample")?.addEventListener("click", () => {
    const s = store.seed.sampleOrder;
    const form = $("#cake-form");
    for (const [k, v] of Object.entries({ size: s.size, layers: s.layers, sponge: s.sponge, filling: s.filling, frosting: s.frosting, topping1: "None", topping2: "None", topping3: "None", box: s.box, board: s.board, delivery: s.delivery, distanceKm: s.distanceKm })) {
      if (form.elements[k]) form.elements[k].value = v;
    }
    form.elements.candles.checked = !!s.candles;
    form.elements.carryBag.checked = !!s.carryBag;
    renderLiveQuote(form);
  });
  $("#mark-paid")?.addEventListener("click", () => {
    const id = $("#mark-paid").dataset.id;
    const o = store.state.orders.find((x) => x.id === id);
    store.updateOrder(id, { status: "Paid", advancePaid: o?.quoteSnapshot?.finalQuote || o?.advancePaid });
    render();
  });
  $("#delete-order")?.addEventListener("click", () => {
    if (confirm("Delete this order?")) {
      store.deleteOrder($("#delete-order").dataset.id);
      location.hash = "#/orders";
    }
  });
  $("#use-stock")?.addEventListener("click", () => {
    const id = $("#use-stock").dataset.id;
    const o = store.state.orders.find((x) => x.id === id);
    if (!o) return;
    const q = o.quoteSnapshot || store.quote(o);
    const usage = [];
    const addRecipe = (name, kg) => {
      store.state.recipes.filter((r) => r.name === name).forEach((line) => {
        usage.push({ item: line.ingredient, qty: line.qtyPerKg * kg * (1 + store.state.settings.wastage) });
      });
    };
    addRecipe(o.sponge, q.lines.sponge.qty);
    addRecipe(o.filling, q.lines.filling.qty);
    addRecipe(o.frosting, q.lines.crumbCoat.qty + q.lines.frosting.qty);
    for (const t of q.lines.toppings) if (t.material && t.qty) usage.push({ item: t.material, qty: t.qty });
    if (o.box) usage.push({ item: o.box, qty: 1 });
    if (o.board) usage.push({ item: o.board, qty: 1 });
    if (o.candles && q.lines.candles.qty) { const a = findAddon(store.state.settings.addons, "candles"); if (a) usage.push({ item: a.material, qty: q.lines.candles.qty }); }
    if (o.carryBag && q.lines.carryBag.qty) { const a = findAddon(store.state.settings.addons, "carryBag"); if (a) usage.push({ item: a.material, qty: q.lines.carryBag.qty }); }
    for (const a of q.lines.extraAddons || []) if (a.material && a.qty) usage.push({ item: a.material, qty: a.qty });
    store.deductStock(usage);
    alert("Stock updated from this cake’s ingredients, toppings and packaging.");
    render();
  });

  $("#recipe-pick")?.addEventListener("change", (e) => {
    location.hash = `#/recipes?recipe=${encodeURIComponent(e.target.value)}&kg=${encodeURIComponent(route.kg||1)}`;
  });
  $("#scale-kg")?.addEventListener("change", (e) => {
    location.hash = `#/recipes?recipe=${encodeURIComponent(route.recipe||$("#recipe-pick").value)}&kg=${encodeURIComponent(e.target.value)}`;
  });

  document.querySelectorAll("[data-mat]").forEach((el) => {
    el.addEventListener("change", () => {
      const name = el.dataset.mat;
      const field = el.dataset.field;
      let val = el.value;
      if (field !== "unit") val = Number(val);
      const patch = { [field]: val };
      if (field === "packPrice" || field === "packSize") {
        const m = store.state.rawMaterials.find((x) => x.name === name);
        if (m && m.estimated) patch.estimated = true;
      }
      store.updateMaterial(name, patch);
      render();
    });
  });
  $("#add-mat")?.addEventListener("click", () => {
    try {
      store.addMaterial({
        category: $("#new-cat").value,
        name: $("#new-name").value.trim(),
        packSize: Number($("#new-size").value),
        unit: $("#new-unit").value.trim(),
        packPrice: Number($("#new-price").value),
        notes: "",
        estimated: false,
      });
      render();
    } catch (err) { alert(err.message); }
  });
  document.querySelectorAll("[data-inv]").forEach((el) => {
    el.addEventListener("change", () => {
      store.updateInventory(el.dataset.inv, { [el.dataset.field]: Number(el.value) });
      render();
    });
  });

  $("#save-settings")?.addEventListener("click", () => {
    const v = (id) => Number($(id).value);
    store.updateSettings({
      wastage: v("#set-wastage") / 100,
      markup: v("#set-markup") / 100,
      gst: v("#set-gst") / 100,
      rounding: v("#set-round"),
      labourRate: v("#set-labour"),
      baseHours: v("#set-basehrs"),
      hoursPerKg: v("#set-hrskg"),
      hoursPerLayer: v("#set-hrslayer"),
      hoursPerTopping: v("#set-hrstop"),
      fillingGPerKg: v("#set-fill"),
      crumbCoatGPerKgPerLayer: v("#set-crumb"),
      finalFrostingGPerKgPerLayer: v("#set-frost"),
      deliveryRatePerKm: v("#set-delrate"),
      tripMultiplier: v("#set-trip"),
      minDelivery: v("#set-mindel"),
      overheads: { gas: v("#oh-gas"), electricity: v("#oh-electricity"), water: v("#oh-water"), rent: v("#oh-rent"), equipment: v("#oh-equipment") },
    });
    $("#save-settings").textContent = "Saved";
  });
  $("#sync-mode")?.addEventListener("change", (e) => {
    $("#cfg-appsscript").hidden = e.target.value !== "appsscript";
    $("#cfg-google").hidden = e.target.value !== "google";
  });
  $("#sync-save")?.addEventListener("click", async () => {
    syncCfg = {
      ...syncCfg,
      mode: $("#sync-mode").value,
      url: $("#sync-url")?.value.trim() || syncCfg.url,
      token: $("#sync-token")?.value || syncCfg.token,
      clientId: $("#sync-client")?.value.trim() || syncCfg.clientId,
      spreadsheetId: $("#sync-sheet")?.value.trim() || syncCfg.spreadsheetId || DEFAULT_SHEET_ID,
    };
    saveSyncConfig(syncCfg);
    if (syncCfg.mode === "google" && !cachedGoogleToken() && syncCfg.clientId) {
      try { await signInWithGoogle(syncCfg.clientId); } catch (err) { alert("Google sign-in: " + err.message); }
    }
    setupData();
    await refreshFromSheet();
  });
  $("#sync-signin")?.addEventListener("click", async () => {
    const id = $("#sync-client")?.value.trim() || syncCfg.clientId;
    if (!id) return alert("Paste the OAuth client ID first.");
    syncCfg = { ...syncCfg, clientId: id };
    saveSyncConfig(syncCfg);
    try { await signInWithGoogle(id); } catch (err) { return alert("Google sign-in: " + err.message); }
    setupData();
    await refreshFromSheet();
  });
  $("#sync-signout")?.addEventListener("click", () => { clearGoogleToken(); render(); });
  $("#sync-now")?.addEventListener("click", refreshFromSheet);
  $("#sync-refresh")?.addEventListener("click", refreshFromSheet);
  document.querySelectorAll("[data-rline]").forEach((el) => {
    el.addEventListener("change", () => {
      const val = el.dataset.field === "qtyPerKg" ? Number(el.value) : el.value;
      store.updateRecipeLine(el.dataset.rline, { [el.dataset.field]: val });
      render();
    });
  });
  $("#rl-add")?.addEventListener("click", (e) => {
    const qty = Number($("#rl-qty").value);
    if (!(qty > 0)) return alert("Enter the quantity for 1 kg.");
    store.addRecipeLine({ name: e.target.dataset.recipe, type: e.target.dataset.type, ingredient: $("#rl-ing").value, qtyPerKg: qty });
    render();
  });
  $("#change-pin")?.addEventListener("click", async () => {
    const pin = $("#new-pin").value.trim();
    if (pin.length < 4) return alert("Use at least 4 digits.");
    await setPin(pin);
    alert("PIN updated on this device.");
  });
  $("#lock-now")?.addEventListener("click", () => { lockSession(); render(); });
  $("#export-json")?.addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(store.exportState(), null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "bakery-workbook-backup.json";
    a.click();
  });
  $("#export-orders")?.addEventListener("click", () => {
    const csv = toCsv(ordersToSheetRows(store.state.orders));
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    a.download = "bakery-orders.csv";
    a.click();
  });
  $("#import-json")?.addEventListener("click", () => $("#import-file").click());
  $("#import-file")?.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      store.importState(JSON.parse(await file.text()));
      $("#io-msg").textContent = "Backup imported.";
      render();
    } catch { alert("Could not read that file."); }
  });
  $("#reset-data")?.addEventListener("click", () => {
    if (confirm("Reset recipes, prices, stock and orders to the original workbook?")) {
      store.resetToSeed();
      render();
    }
  });
  bindV3();
}

function handleErr(err, section, type, retry) {
  if (err && err.full && store.sheetMode && confirm(err.message.replace(/ Tap “Make room”.*$/, "") + "\n\nInsert 5 empty rows inside that list in her sheet now? (Sheets widens the Calculator's dropdowns to include them.)")) {
    store.makeRoom(section, type);
    if (retry) {
      try { retry(); return; } catch (e2) { alert(e2 && e2.message ? e2.message : String(e2)); }
    }
    render();
    return;
  }
  alert(err && err.message ? err.message : String(err));
}

function bindV3() {
  const on = (sel, ev, fn) => document.querySelectorAll(sel).forEach((el) => el.addEventListener(ev, (e) => fn(el, e)));
  // ----- order form: customer quick-fill -----
  const form = $("#cake-form") || $("#unit-form");
  if (form) {
    const nameEl = form.elements.customer, phoneEl = form.elements.phone, hint = $("#cust-hint");
    const fill = (c) => {
      if (!c) { hint.textContent = ""; return; }
      const n = store.ordersOf(c).length;
      hint.textContent = `Saved customer · ${n} past order${n === 1 ? "" : "s"}${c.notes ? " · " + c.notes : ""}`;
    };
    nameEl?.addEventListener("change", () => {
      const c = store.findCustomer({ name: nameEl.value });
      if (c && !phoneEl.value) phoneEl.value = c.phone || "";
      fill(c);
    });
    phoneEl?.addEventListener("change", () => {
      const c = store.findCustomer({ phone: phoneEl.value });
      if (c && !nameEl.value) nameEl.value = c.name;
      fill(c);
    });
  }
  // ----- orders: one-tap status + payments -----
  on("[data-next]", "click", (el, e) => { e.stopPropagation(); store.setStatus(el.dataset.next, el.dataset.status); render(); });
  on("[data-set-status]", "click", (el, e) => { e.preventDefault(); store.setStatus(el.dataset.id, el.dataset.setStatus); render(); });
  $("#pay-add")?.addEventListener("click", (e) => {
    try { store.recordPayment(e.target.dataset.id, $("#pay-amt").value); render(); } catch (err) { handleErr(err); }
  });
  $("#pay-full")?.addEventListener("click", (e) => { store.markPaidInFull(e.target.dataset.id); render(); });

  // ----- menu & prices -----
  on("[data-menu-sec]", "change", (el) => {
    const { menuSec: sec, key, field } = el.dataset;
    let val = el.value;
    try {
      if (field === "name") {
        if (!val.trim()) { el.value = key; return; }
        if (sec === "layers") val = Number(val);
      }
      store.updateMenuItem(sec, key, { [field]: val });
      render();
    } catch (err) { el.value = key; handleErr(err); }
  });
  on("[data-sell-kind]", "change", (el) => { store.setSellingPrice(el.dataset.sellKind, el.dataset.sellItem, el.value); $("#menu-msg") && ($("#menu-msg").textContent = `Selling price saved for ${el.dataset.sellItem}.`); });
  on("[data-menu-del]", "click", (el, e) => {
    e.preventDefault();
    if (!confirm(`Remove “${el.dataset.key}” from the list? Old orders keep it.`)) return;
    store.removeMenuItem(el.dataset.menuDel, el.dataset.key);
    render();
  });
  on("[data-room-sec]", "click", (el, e) => {
    e.preventDefault();
    if (!confirm("Insert 5 empty rows inside this list in her sheet? Rows below move down; the Calculator's ranges widen to include them.")) return;
    store.makeRoom(el.dataset.roomSec, el.dataset.roomType);
    render();
  });
  on("[data-newmat]", "change", (el) => { const box = $("#" + el.dataset.newmat + "-newmat"); if (box) box.hidden = el.value !== "__new"; });
  on("[data-menu-orphan]", "click", (el) => {
    const doAdd = () => { store.addMenuItem("recipeMenu", { type: el.dataset.type, name: el.dataset.menuOrphan }); render(); };
    try { doAdd(); } catch (err) { handleErr(err, "recipeMenu", el.dataset.type, doAdd); }
  });
  const v = (id) => $("#" + id)?.value ?? "";
  const newMat = (prefix, name) => (v(prefix + "-mat") === "__new" ? { name, spec: { packSize: v(prefix + "-psize"), unit: v(prefix + "-punit"), packPrice: v(prefix + "-pprice") } } : { name: v(prefix + "-mat") });
  on("[data-menu-add]", "click", (el) => {
    const sec = el.dataset.menuAdd;
    const doAdd = () => {
      if (sec === "sizes") store.addMenuItem("sizes", { name: v("add-sizes-name").trim(), kg: v("add-sizes-kg"), servings: v("add-sizes-serv"), sellingPrice: v("add-sizes-sell") });
      else if (sec === "layers") store.addMenuItem("layers", { value: v("add-layers-value") });
      else if (sec === "toppings") {
        const name = v("add-toppings-name").trim();
        const m = newMat("add-toppings", name);
        store.addMenuItem("toppings", { name, material: m.name, newMaterial: m.spec, fixedQty: v("add-toppings-fixed"), qtyPerKg: v("add-toppings-perkg"), sellingPrice: v("add-toppings-sell") });
      } else if (sec === "boxes" || sec === "boards") {
        store.addMenuItem(sec, { name: v(`add-${sec}-name`).trim(), newMaterial: { packSize: v(`add-${sec}-psize`), unit: "nos", packPrice: v(`add-${sec}-pprice`) }, sellingPrice: v(`add-${sec}-sell`) });
      } else if (sec === "addons") {
        const name = v("add-addons-name").trim();
        const m = newMat("add-addons", name);
        store.addMenuItem("addons", { name, material: m.name, newMaterial: m.spec, qty: v("add-addons-qty"), sellingPrice: v("add-addons-sell") });
      }
      render();
    };
    try { doAdd(); } catch (err) { handleErr(err, sec, undefined, doAdd); }
  });

  // ----- product categories -----
  on("[data-cat]", "change", (el) => {
    const name = el.dataset.cat, f = el.dataset.field;
    const c = store.category(name);
    if (!c) return;
    const val = el.type === "checkbox" ? el.checked : el.value;
    if (f === "name" && !String(val).trim()) { el.value = name; return; }
    try {
      const next = { ...c, [f]: val };
      if (f === "pricing" && val === "Per unit" && !(Number(c.unitPrice) > 0)) next.unitPrice = 90;
      store.saveCategory(next, name);
      render();
      $("#menu-msg") && ($("#menu-msg").textContent = `Saved ${next.name}.`);
    } catch (err) { handleErr(err); render(); }
  });
  on("[data-cat-del]", "click", (el, e) => {
    e.preventDefault();
    if (!confirm(`Remove the “${el.dataset.catDel}” category? Old orders keep it.`)) return;
    try { store.removeCategory(el.dataset.catDel); render(); } catch (err) { handleErr(err); }
  });
  $("#add-cat")?.addEventListener("click", () => {
    try {
      store.saveCategory({ name: v("add-cat-name"), pricing: "Per unit", unitPrice: v("add-cat-price"), unit: v("add-cat-unit").trim() || "piece", costPerUnit: v("add-cat-cost"), active: true });
      render();
    } catch (err) { handleErr(err); }
  });

  // ----- price went up -----
  $("#pu-item")?.addEventListener("change", (e) => { location.hash = "#/price-up?item=" + encodeURIComponent(e.target.value); });
  ["#pu-price", "#pu-size"].forEach((id) => $(id)?.addEventListener("input", renderImpact));
  if ($("#pu-impact")) renderImpact();
  $("#pu-save")?.addEventListener("click", () => {
    const name = $("#pu-item").value;
    const price = Number($("#pu-price").value), size = Number($("#pu-size").value);
    if (!(price >= 0) || !(size > 0)) return alert("Enter the new pack price and pack size.");
    store.priceWentUp(name, price, size);
    const im = store.priceImpact(name, price, size);
    $("#pu-msg").textContent = `Saved. ${im.recipes.length} recipe cost${im.recipes.length === 1 ? "" : "s"} and all new quotes now use ${rupees(price)} per pack.`;
  });

  // ----- recipes -----
  on("[data-rdel]", "click", (el) => {
    if (!confirm("Remove this ingredient line from the recipe?")) return;
    store.removeRecipeLine(el.dataset.rdel);
    render();
  });
  $("#r-notes")?.addEventListener("change", (e) => { store.setRecipeNotes(e.target.dataset.recipe, e.target.value, e.target.dataset.type); });
  $("#rn-more")?.addEventListener("click", (e) => {
    e.preventDefault();
    const first = $(".rn-line");
    const c = first.cloneNode(true);
    c.querySelector("select").value = ""; c.querySelector("input").value = "";
    $("#rn-lines").appendChild(c);
  });
  $("#rn-save")?.addEventListener("click", () => {
    const lines = [...document.querySelectorAll(".rn-line")].map((r) => ({ ingredient: r.querySelector("select").value, qtyPerKg: Number(r.querySelector("input").value) }));
    const name = $("#rn-name").value.trim(), type = $("#rn-type").value;
    const doCreate = () => {
      store.createRecipe({ name, type, lines, method: $("#rn-method").value.trim() });
      location.hash = `#/recipes?recipe=${encodeURIComponent(name)}`;
    };
    try { doCreate(); } catch (err) {
      if (err.full) handleErr(err, "recipeMenu", type, doCreate);
      else $("#rn-msg").textContent = err.message;
    }
  });

  // ----- inventory: purchases -----
  $("#lp-item")?.addEventListener("change", (e) => { location.hash = "#/inventory?item=" + encodeURIComponent(e.target.value); });
  $("#lp-save")?.addEventListener("click", () => {
    try {
      const rec = store.logPurchase({ item: $("#lp-item").value, packs: $("#lp-packs").value, pricePerPack: $("#lp-price").value, date: $("#lp-date").value, notes: $("#lp-note").value.trim() });
      render();
      $("#lp-msg") && ($("#lp-msg").textContent = `Added ${rec.qty} ${rec.unit} of ${rec.item}.`);
    } catch (err) { handleErr(err); }
  });

  // ----- pricing rules + delivery zones -----
  on("[data-rule]", "change", (el) => {
    const r = (store.state.rules || []).find((x) => x.id === el.dataset.rule);
    if (!r) return;
    const f = el.dataset.field;
    try { store.saveRule({ ...r, [f]: el.type === "checkbox" ? el.checked : el.value }); } catch (err) { handleErr(err); render(); }
  });
  on("[data-rule-del]", "click", (el) => { if (confirm("Remove this rule?")) { store.removeRule(el.dataset.ruleDel); render(); } });
  $("#rule-add-rush")?.addEventListener("click", () => { try { store.saveRule({ name: uniqueName("Express / rush order", (store.state.rules || []).map((r) => r.name)), effect: "Charge", type: "%", amount: 20, active: true, applyTo: "When ticked" }); render(); } catch (err) { handleErr(err); } });
  $("#rule-add-fest")?.addEventListener("click", () => { try { store.saveRule({ name: uniqueName("Festival discount", (store.state.rules || []).map((r) => r.name)), effect: "Discount", type: "%", amount: 10, active: false, applyTo: "Every order" }); render(); } catch (err) { handleErr(err); } });
  on("[data-zone]", "change", (el) => {
    const z = (store.state.zones || []).find((x) => x.id === el.dataset.zone);
    if (!z) return;
    try { store.saveZone({ ...z, [el.dataset.field]: el.value }); } catch (err) { handleErr(err); render(); }
  });
  on("[data-zone-del]", "click", (el) => { if (confirm("Remove this delivery zone?")) { store.removeZone(el.dataset.zoneDel); render(); } });
  $("#zone-add")?.addEventListener("click", () => { try { store.saveZone({ name: $("#zone-name").value, charge: $("#zone-charge").value }); render(); } catch (err) { handleErr(err); } });

  // ----- customers -----
  $("#cust-search")?.addEventListener("input", (e) => { location.hash = "#/customers?q=" + encodeURIComponent(e.target.value); });
  $("#nc-save")?.addEventListener("click", () => { try { store.saveCustomer({ name: $("#nc-name").value, phone: $("#nc-phone").value, notes: $("#nc-notes").value }); render(); } catch (err) { handleErr(err); } });
  $("#cu-save")?.addEventListener("click", (e) => {
    try {
      const c = store.saveCustomer({ id: e.target.dataset.id || undefined, name: $("#cu-name").value, phone: $("#cu-phone").value, notes: $("#cu-notes").value });
      location.hash = "#/customer/" + encodeURIComponent(c.id);
      render();
    } catch (err) { handleErr(err); }
  });
  $("#cu-del")?.addEventListener("click", (e) => { if (confirm("Remove from the customer list? Their orders stay.")) { store.removeCustomer(e.target.dataset.id); location.hash = "#/customers"; } });
}
function uniqueName(base, taken) {
  let n = base, i = 2;
  while (taken.includes(n)) n = `${base} ${i++}`;
  return n;
}

let seedData;
function setupData() {
  engine?.stop();
  syncCfg = loadSyncConfig();
  const transport = makeTransport(syncCfg, { getToken: () => cachedGoogleToken() });
  engine = new SyncEngine({
    transport,
    onState: (grid) => store.applySheetGrid(grid),
    onFlushed: (done) => store.markFlushed(done),
    onStatus: updateSyncPill,
  });
  store = new Store(seedData, { sync: engine });
}

async function main() {
  seedData = await fetch("./js/seed-data.json").then((r) => r.json());
  setupData();
  window.addEventListener("hashchange", render);
  // Writes waiting offline go out as soon as the connection is back (no re-read of the sheet).
  window.addEventListener("online", () => engine.flush().catch(() => {}));
  render(); // last loaded copy first, so the app opens instantly
  if (engine.transport) {
    await engine.pull(); // fresh read of every tab on every page load
    const typing = document.activeElement?.matches?.("input, select, textarea");
    if (!typing) render(); // don't wipe a form she has started filling in
    else updateSyncPill();
  }
}

main();
