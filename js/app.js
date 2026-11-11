import { Store, checkPin, setPin, isUnlocked, unlockSession, lockSession, ordersToSheetRows, toCsv } from "./store.js";
import { rupees, rupeesWhole, scaleRecipe } from "./calculator.js";
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
  return `${o.size} ${o.sponge || ""}`.replace("Sponge", "").trim();
}

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
  return `<div class="quote-box">
    <div class="lbl">Final quote</div>
    <div class="quote-num">${rupeesWhole(q.finalQuote)}</div>
    <div class="quote-grid">
      <div><span>Raw materials</span><br><strong>${rupees(q.rawMaterialCost)}</strong></div>
      <div><span>Labour</span><br><strong>${rupees(q.labour)}</strong></div>
      <div><span>Overheads</span><br><strong>${rupees(q.overheads)}</strong></div>
      <div><span>Profit</span><br><strong>${rupees(q.profit)}</strong></div>
      <div><span>Delivery</span><br><strong>${rupees(q.delivery)}</strong></div>
      <div><span>Per kg (ex-del.)</span><br><strong>${rupees(q.pricePerKgExclDelivery)}</strong></div>
    </div>
  </div>`;
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
  ];
  return `<table class="breakdown"><thead><tr><th>Component</th><th>Selection</th><th class="num">Qty</th><th class="num">Cost</th></tr></thead><tbody>
    ${rows.map(([l,x]) => `<tr><td>${l}</td><td>${x.name || "—"}</td><td class="num">${x.qty ? Number(x.qty).toLocaleString("en-IN",{maximumFractionDigits:3}) + " " + (x.unit||"") : "—"}</td><td class="num">${rupees(x.cost)}</td></tr>`).join("")}
    <tr class="total-row"><td colspan="3">Ingredients (incl. wastage)</td><td class="num">${rupees(q.ingredientCost)}</td></tr>
    <tr><td colspan="3">Packaging</td><td class="num">${rupees(q.packagingCost)}</td></tr>
    <tr><td colspan="3">Total cost of making</td><td class="num">${rupees(q.totalCost)}</td></tr>
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
  const size = g("size");
  const sz = settings.sizes.find((s) => s.name === size);
  return {
    customer: g("customer"),
    phone: g("phone"),
    size,
    weightKg: sz ? sz.kg : 0,
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
  const p = { size: "0.5 kg", layers: 2, sponge: "Chocolate Sponge", filling: "Chocolate Ganache", frosting: "Chocolate Ganache Frosting", topping1: "None", topping2: "None", topping3: "None", box: "Cake box 8x8x5 in (up to 1 kg)", board: "Cake board 8 in", delivery: "Home delivery", distanceKm: 8, carryBag: true, candles: false, status: "Enquiry", ...pre };
  return `<form id="cake-form">
    <div class="card">
      <h2>Order details</h2>
      <div class="row-2">
        <div><label>Customer</label><input name="customer" value="${escapeHtml(p.customer||"")}" placeholder="Name" required></div>
        <div><label>Phone</label><input name="phone" type="tel" value="${escapeHtml(p.phone||"")}" placeholder="WhatsApp number"></div>
      </div>
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
        <div><label>Size</label><select name="size">${options(settings.sizes.map(s=>s.name), p.size)}</select></div>
        <div><label>Layers</label><select name="layers">${options(settings.layers.map(String), String(p.layers))}</select></div>
      </div>
      <label>Sponge</label><select name="sponge">${options(sponges, p.sponge)}</select>
      <label>Filling</label><select name="filling">${options(fillings, p.filling)}</select>
      <label>Frosting</label><select name="frosting">${options(frostings, p.frosting)}</select>
      <div class="row-3">
        <div><label>Topping 1</label><select name="topping1">${options(tops, p.topping1||"None")}</select></div>
        <div><label>Topping 2</label><select name="topping2">${options(tops, p.topping2||"None")}</select></div>
        <div><label>Topping 3</label><select name="topping3">${options(tops, p.topping3||"None")}</select></div>
      </div>
      <div class="row-2">
        <div><label>Cake box</label><select name="box">${options(settings.boxes, p.box)}</select></div>
        <div><label>Cake board</label><select name="board">${options(settings.boards, p.board)}</select></div>
      </div>
      <label class="check"><input type="checkbox" name="candles" ${p.candles?"checked":""}> Candles & knife set</label>
      <label class="check"><input type="checkbox" name="carryBag" ${p.carryBag?"checked":""}> Carry bag</label>
      <div class="row-2">
        <div><label>Delivery</label><select name="delivery">${options(["Home delivery","Pickup"], p.delivery)}</select></div>
        <div><label>Distance, one way (km)</label><input name="distanceKm" type="number" min="0" step="0.1" value="${p.distanceKm||0}"></div>
      </div>
      <label>Hours needed (leave blank for suggested)</label>
      <input name="hoursOverride" type="number" min="0" step="0.25" value="${p.hoursOverride ?? ""}" placeholder="Auto from size, layers, toppings">
    </div>
    <div id="live-quote" style="margin-top:12px"></div>
  </form>`;
}

function renderLiveQuote(form) {
  const input = calcInputsFromForm(form, store.state.settings);
  const q = store.quote(input);
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
  return `${topbar()}<h2 class="page-title">New order</h2>
    ${cakeForm({}, store.state.settings)}
    <div class="btn-row"><button class="btn" id="save-order">Save order</button><button class="btn secondary" data-go="#/orders">Cancel</button></div>
    ${nav("orders")}`;
}

function viewOrder(id) {
  const o = store.state.orders.find((x) => x.id === id);
  if (!o) return `${topbar()}<p>Order not found.</p>${nav("orders")}`;
  const q = o.quoteSnapshot || store.quote(o);
  const bal = Math.max((q.finalQuote || 0) - (o.advancePaid || 0), 0);
  return `${topbar()}<h2 class="page-title">${escapeHtml(o.customer||"Order")}</h2>
    <p class="meta">${o.id} · Created ${fmtDate(o.createdAt)}</p>
    ${quoteBox(q)}
    <p class="tiny">Advance paid ${rupees(o.advancePaid||0)} · Balance ${rupees(bal)} · Profit ${rupees(q.profit)}</p>
    ${(() => { const live = store.quote(o); return q.fromSheet && Math.round(live.finalQuote) !== Math.round(q.finalQuote) ? `<p class="tiny warn-msg">Saved quote is ${rupeesWhole(q.finalQuote)} (from the sheet). With today’s prices it would be ${rupeesWhole(live.finalQuote)}. Tap “Update order” to re-price.</p>` : ""; })()}
    ${cakeForm(o, store.state.settings)}
    <div class="btn-row">
      <button class="btn" id="save-order" data-id="${o.id}">Update order</button>
      <button class="btn sage" id="mark-paid" data-id="${o.id}">Mark paid</button>
      <button class="btn secondary" id="use-stock" data-id="${o.id}">Deduct ingredients from stock</button>
      <button class="btn danger" id="delete-order" data-id="${o.id}">Delete</button>
    </div>
    <div class="card" style="margin-top:12px">${breakdownTable(q)}</div>
    ${nav("orders")}`;
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
      <table class="breakdown"><thead><tr><th>Ingredient</th><th class="num">Qty / 1 kg</th><th class="num">Cost</th></tr></thead><tbody>
        ${lines.map((l) => {
          const mat = store.state.rawMaterials.find((m) => m.name === l.ingredient);
          const rate = mat ? (Number(mat.packPrice)/Number(mat.packSize)||0) : 0;
          return `<tr><td><select class="mini" data-rline="${escapeHtml(l.key)}" data-field="ingredient">${options(matNames, l.ingredient)}</select>${mat?.estimated?'<span class="tag-est">ESTIMATED</span>':""}${!mat?'<span class="tag-est">MISSING</span>':""}</td><td class="num"><input class="mini num" data-rline="${escapeHtml(l.key)}" data-field="qtyPerKg" type="number" step="any" min="0" value="${l.qtyPerKg}"> ${l.unit||""}</td><td class="num">${rupees(l.qtyPerKg*rate)}</td></tr>`;
        }).join("")}
        <tr class="total-row"><td colspan="2">Per 1 kg</td><td class="num">${rupees(costs[chosen]||0)}</td></tr>
      </tbody></table>
      <p class="tiny">Edit a quantity or ingredient and it is saved${store.sheetMode ? " to the Recipes tab of the sheet" : ""}. Quantities are for 1 kg of this recipe, in the item’s own unit.</p>
      <div class="row-3">
        <div><label>Add ingredient</label><select id="rl-ing">${options(matNames, matNames[0])}</select></div>
        <div><label>Qty / 1 kg</label><input id="rl-qty" type="number" step="any" min="0" value="0"></div>
        <div><label>&nbsp;</label><button class="btn secondary" id="rl-add" data-recipe="${escapeHtml(chosen)}" data-type="${escapeHtml(type)}">Add line</button></div>
      </div>
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
      <div><label>Reorder level</label><input data-inv="${escapeHtml(r.item)}" data-field="reorderLevel" type="number" step="any" value="${r.reorderLevel}"></div>
    </div>
  </div>`).join("");
  return `${topbar()}<h2 class="page-title">Inventory</h2>
    <div class="stats"><div class="card"><div class="stat-lbl">Stock value</div><div class="stat-val">${rupeesWhole(value)}</div></div>
    <div class="card"><div class="stat-lbl">Items to reorder</div><div class="stat-val">${low}</div></div></div>
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
    <div class="card" style="margin-top:12px"><h2>By month</h2>
      ${months.length ? `<table class="breakdown"><thead><tr><th>Month</th><th class="num">Orders</th><th class="num">Revenue</th><th class="num">Cost</th><th class="num">Profit</th></tr></thead><tbody>
        ${months.map(m=>`<tr><td>${m.month}</td><td class="num">${m.orders}</td><td class="num">${rupees(m.revenue)}</td><td class="num">${rupees(m.cost)}</td><td class="num">${rupees(m.profit)}</td></tr>`).join("")}
      </tbody></table>` : `<p class="empty">Profit appears here once you save orders.</p>`}
    </div>
    <div class="card" style="margin-top:12px"><h2>Per order</h2>
      ${all.length ? `<table class="breakdown"><thead><tr><th>Order</th><th>Status</th><th class="num">Quote</th><th class="num">Cost</th><th class="num">Profit</th></tr></thead><tbody>
        ${all.map(o=>`<tr data-go="#/order/${o.id}" style="cursor:pointer"><td>${escapeHtml(o.customer||o.id)}<div class="tiny">${fmtDate(o.deliveryDate||o.createdAt)}</div></td><td><span class="pill ${o.status}">${o.status}</span></td><td class="num">${rupees(o.quoteSnapshot?.finalQuote||0)}</td><td class="num">${rupees(o.quoteSnapshot?.totalCost||0)}</td><td class="num">${rupees(o.quoteSnapshot?.profit||0)}</td></tr>`).join("")}
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
      ${conflicts.length ? `<h3>Changes not applied</h3><ul class="tiny">${conflicts.map(c=>`<li>${escapeHtml(fmtTime(c.at))} · ${escapeHtml(c.what)}: ${escapeHtml(c.reason)}</li>`).join("")}</ul>` : ""}
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
      <p class="tiny">Cake sizes, layers, recipe menu, toppings, boxes, boards and add-ons are edited in the sheet; the app picks them up on the next load.</p>
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
  if (parts[0] === "order" && parts[1] === "new") return { name: "order-new" };
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
  const form = $("#cake-form");
  if (!form) return;
  const update = () => renderLiveQuote(form);
  form.addEventListener("input", update);
  form.addEventListener("change", update);
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
    if (o.candles) usage.push({ item: "Candles & knife set", qty: 1 });
    if (o.carryBag) usage.push({ item: "Carry bag (printed)", qty: 1 });
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
