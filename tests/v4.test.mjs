// v4: product categories (Categories tab), per-unit orders, manual cake sizes.
// Runs against the local mock of the Google Sheets REST API (never the real sheet).
import { readFileSync } from "fs";
import { startMock } from "./mock-sheets.mjs";

const mem = new Map();
globalThis.localStorage = { getItem: (k) => (mem.has(k) ? mem.get(k) : null), setItem: (k, v) => mem.set(k, String(v)), removeItem: (k) => mem.delete(k), clear: () => mem.clear() };
globalThis.sessionStorage = { getItem() { return null; }, setItem() {}, removeItem() {} };

const { Store } = await import("../js/store.js");
const { SyncEngine } = await import("../js/sync.js");
const { GoogleSheetsTransport } = await import("../js/transports.js");
const { parseSizeKg, sizeInfo, calculateQuote, estimateServings, resolveSizeInput } = await import("../js/calculator.js");
const seed = JSON.parse(readFileSync(new URL("../js/seed-data.json", import.meta.url)));

let failed = 0, passed = 0;
function check(cond, label, extra) {
  if (cond) { passed++; console.log("  ok  " + label); }
  else { failed++; console.error("  FAIL " + label, extra !== undefined ? JSON.stringify(extra) : ""); }
}

console.log("\n=== pure: manual cake sizes ===");
check(parseSizeKg("750 g") === 0.75 && parseSizeKg("1.2 kg") === 1.2 && parseSizeKg("750g") === 0.75 && parseSizeKg("Tier cake") === 0, "parseSizeKg grams / kg");
const S = seed.settings;
check(sizeInfo("0.5 kg", S).servings === "4-6", "preset size keeps its servings");
const custom = sizeInfo("750 g", S);
check(custom.kg === 0.75 && /^~\d+-\d+$/.test(custom.servings), "manual 750 g -> 0.75 kg, estimated servings", custom);
check(estimateServings(1, S.sizes) === "8-10", "manual size equal to a preset uses its servings");
const data = { rawMaterials: seed.rawMaterials, recipes: seed.recipes, settings: S };
const q500 = calculateQuote({ ...seed.sampleOrder, size: "500 g" }, data);
const qPreset = calculateQuote(seed.sampleOrder, data);
check(q500.finalQuote === qPreset.finalQuote && qPreset.finalQuote === 780, "500 g manual prices the same as the 0.5 kg preset (₹780)", [q500.finalQuote, qPreset.finalQuote]);
const q12 = calculateQuote({ ...seed.sampleOrder, size: "1.2 kg" }, data);
check(q12.weightKg === 1.2 && q12.finalQuote > qPreset.finalQuote && q12.weightOk, "1.2 kg manual: calculator scales, weight check OK", [q12.weightKg, q12.finalQuote]);
check(sizeInfo("500 g", S).name === "0.5 kg" && !sizeInfo("500 g", S).custom, "sizeInfo maps 500 g to the 0.5 kg preset");

console.log("\n=== pure: resolveSizeInput (combobox) ===");
const R = (text, unit="kg") => resolveSizeInput(text, unit, S.sizes);
check(R("0.5 kg").size === "0.5 kg" && !R("0.5 kg").custom, "preset name stays the preset");
check(R("500 g").size === "0.5 kg" && !R("500 g").custom && R("500 g").servings === "4-6", "500 g matches the 0.5 kg preset");
check(R("500", "g").size === "0.5 kg" && !R("500", "g").custom, "bare 500 in g mode matches 0.5 kg");
check(R("0.5", "kg").size === "0.5 kg" && !R("0.5", "kg").custom, "bare 0.5 in kg mode matches 0.5 kg");
check(R("500 g", "kg").size === "0.5 kg", "explicit unit wins over the toggle");
const r750g = R("750", "g");
check(r750g.custom && r750g.size === "750 g" && r750g.weightKg === 0.75 && /^~\d+-\d+$/.test(r750g.servings), "bare 750 in g mode is a custom 750 g", r750g);
const r12 = R("1.2", "kg");
check(r12.custom && r12.size === "1.2 kg" && r12.weightKg === 1.2, "bare 1.2 in kg mode is a custom 1.2 kg", r12);
check(R("1.2 kg", "g").size === "1.2 kg" && R("1.2 kg", "g").custom, "1.2 kg stays custom even in g mode");
check(!R("").valid && !R("abc").valid && !R("0", "g").valid, "empty, text and zero are invalid");
check(R("750g").size === "750 g" && R("750g").custom, "750g normalises to 750 g");

// ---------------- sheet mode ----------------
const mock = await startMock();
const wb = mock.wb;
const transport = new GoogleSheetsTransport({ spreadsheetId: wb.id, getToken: () => mock.googleToken, base: mock.base });
let store;
const engine = new SyncEngine({ transport, debounceMs: 100000, onState: (g) => store.applySheetGrid(g), onFlushed: (d) => store.markFlushed(d) });
store = new Store(seed, { sync: engine });
const hdr = (tab) => wb.tabs[tab].values[0];
const rowOf = (tab, col, val) => wb.tabs[tab].values.findIndex((r) => r[hdr(tab).indexOf(col)] === val);
const val = (tab, r, col) => wb.tabs[tab].values[r][hdr(tab).indexOf(col)];

console.log("\n=== first sync: Categories tab ===");
check(!wb.tabs.Categories, "fixture: no Categories tab yet");
const formulasBefore = Object.values(wb.tabs).reduce((n, t) => n + t.formulas.flat().filter(Boolean).length, 0);
await engine.pull();
check(engine.status.state === "synced" && engine.status.pending === 0, "first pull ends synced (seeding sent straight away)", engine.status);
check((mock.stats.addedSheets || []).includes("Categories"), "Categories tab created via addSheet");
check(JSON.stringify(hdr("Categories").slice(0, 7)) === JSON.stringify(["Category", "Pricing", "Unit price (₹)", "Unit", "Cost per unit (₹)", "Active", "Notes"]), "Categories headers", hdr("Categories"));
check(mock.cell("Categories", "A2") === "Cake" && mock.cell("Categories", "B2") === "Calculator", "row 2: Cake on the calculator");
check(mock.cell("Categories", "A3") === "Granola bar" && mock.cell("Categories", "C3") === 90 && mock.cell("Categories", "D3") === "bar" && mock.cell("Categories", "F3") === "Yes", "row 3: Granola bar ₹90 per bar");
check(mock.cell("Categories", "A4") === "Cookies" && mock.cell("Categories", "C4") === 90 && mock.cell("Categories", "D4") === "cookie", "row 4: Cookies ₹90 per cookie");
const oh = hdr("Orders");
check(oh.includes("Category") && oh.includes("Quantity") && oh.includes("Unit price (₹)") && oh.indexOf("Category") > oh.indexOf("Delivery charge (₹)"), "Orders: Category / Quantity / Unit price appended to the right", oh);
check(oh[0] === "ID" && oh[2] === "Customer", "Orders: existing columns did not move");

console.log("\n=== reload reads categories from the sheet ===");
const added = (mock.stats.addedSheets || []).length;
await engine.pull();
check((mock.stats.addedSheets || []).length === added && engine.status.pending === 0, "second load: nothing re-seeded");
check(store.activeCategories().map((c) => c.name).join("|") === "Cake|Granola bar|Cookies", "app shows 3 categories from the sheet");
// she edits the sheet by hand: cookies now ₹100, granola hidden
wb.tabs.Categories.values[3][2] = 100;
wb.tabs.Categories.values[2][5] = "No";
await engine.pull();
check(store.category("Cookies").unitPrice === 100 && !store.activeCategories().some((c) => c.name === "Granola bar"), "sheet edits (price, Active = No) show up on reload");
wb.tabs.Categories.values[3][2] = 90;
wb.tabs.Categories.values[2][5] = "Yes";
await engine.pull();

console.log("\n=== per-unit orders ===");
const g = store.addOrder({ category: "Granola bar", quantity: 12, customer: "Meena", phone: "9000011111", deliveryDate: "2026-11-20", delivery: "Home delivery", deliveryCharge: 40, advancePaid: 500, notes: "Honey oats" });
check(g.quoteSnapshot.finalQuote === 12 * 90 + 40 && g.quoteSnapshot.subtotal === 1080, "12 granola bars = ₹1,080 + ₹40 delivery", g.quoteSnapshot);
const ck = store.addOrder({ category: "Cookies", quantity: 20, customer: "Ravi" });
check(ck.quoteSnapshot.finalQuote === 1800, "20 cookies = ₹1,800");
const cake = store.addOrder({ ...seed.sampleOrder, category: "Cake", customer: "Asha", size: "750 g", weightKg: 0.75 });
check(cake.quoteSnapshot.weightKg === 0.75 && cake.quoteSnapshot.finalQuote > 0, "cake order with a manual 750 g size");
await engine.flush();
let r = rowOf("Orders", "ID", g.id);
check(r > 0 && val("Orders", r, "Category") === "Granola bar" && val("Orders", r, "Quantity") === 12 && val("Orders", r, "Unit price (₹)") === 90, "granola order row: Category, Quantity, Unit price", wb.tabs.Orders.values[r]);
check(val("Orders", r, "Quote (₹)") === 1120 && val("Orders", r, "Delivery charge (₹)") === 40 && val("Orders", r, "Balance (₹)") === 620 && val("Orders", r, "Payment status") === "Advance paid", "granola order: total, delivery, balance, payment status");
check(["Size", "Sponge", "Filling", "Frosting", "Box", "Board", "Topping 1"].every((h) => (val("Orders", r, h) ?? "") === ""), "granola order: cake-only columns left blank", ["Size", "Sponge", "Topping 1"].map((h) => val("Orders", r, h)));
const rc = rowOf("Orders", "ID", cake.id);
check(val("Orders", rc, "Category") === "Cake" && val("Orders", rc, "Size") === "750 g" && (val("Orders", rc, "Quantity") ?? "") === "", "cake order row: Category Cake, Size 750 g, Quantity blank");

console.log("\n=== edits write back + reload round trip ===");
store.updateOrder(g.id, { quantity: 15 });
store.recordPayment(g.id, 200);
store.setStatus(ck.id, "Confirmed");
await engine.flush();
r = rowOf("Orders", "ID", g.id);
check(val("Orders", r, "Quantity") === 15 && val("Orders", r, "Quote (₹)") === 15 * 90 + 40 && val("Orders", r, "Advance paid (₹)") === 700, "quantity edit re-prices, payment added", [val("Orders", r, "Quantity"), val("Orders", r, "Quote (₹)"), val("Orders", r, "Advance paid (₹)")]);
check(val("Orders", rowOf("Orders", "ID", ck.id), "Status") === "Confirmed", "one-tap status on a cookie order");
await engine.pull();
const g2 = store.state.orders.find((o) => o.id === g.id);
check(store.isUnitOrder(g2) && g2.quantity === 15 && g2.unitPrice === 90 && g2.quoteSnapshot.finalQuote === 1390, "reloaded granola order is a per-unit order", g2 && g2.quoteSnapshot);
const c2 = store.state.orders.find((o) => o.id === cake.id);
check(!store.isUnitOrder(c2) && c2.category === "Cake" && Math.abs(store.quote(c2).weightKg - 0.75) < 1e-9, "reloaded cake order re-prices from its 750 g size", [c2.category, c2.size, c2.quantity, store.quote(c2).weightKg]);
// a price change in Categories doesn't re-price old orders (they keep their unit price)
store.saveCategory({ ...store.category("Granola bar"), unitPrice: 95 }, "Granola bar");
check(store.quote(g2).finalQuote === 1390, "old order keeps the unit price it was sold at");

console.log("\n=== categories from the Menu screen ===");
store.saveCategory({ name: "Brownies", pricing: "Per unit", unitPrice: 120, unit: "brownie", costPerUnit: 45, active: true });
await engine.flush();
const br = rowOf("Categories", "Category", "Brownies");
check(br > 0 && val("Categories", br, "Unit price (₹)") === 120 && val("Categories", br, "Cost per unit (₹)") === 45, "new category appended to the Categories tab");
check(val("Categories", rowOf("Categories", "Category", "Granola bar"), "Unit price (₹)") === 95, "price change written");
store.saveCategory({ ...store.category("Brownies"), name: "Fudge brownies" }, "Brownies");
await engine.flush();
check(rowOf("Categories", "Category", "Fudge brownies") === br, "rename keeps the same row");
const b = store.addOrder({ category: "Fudge brownies", quantity: 4, customer: "Kiran" });
check(b.quoteSnapshot.totalCost === 180 && b.quoteSnapshot.profit === 300, "cost per unit gives real profit (4 × ₹45 cost)");
store.removeCategory("Fudge brownies");
await engine.flush();
check(rowOf("Categories", "Category", "Fudge brownies") < 0, "removed category cleared from the sheet");
let err = null; try { store.removeCategory("Cake"); } catch (e) { err = e; }
check(err && /calculator/.test(err.message), "can't remove the only calculator category");
err = null; try { store.saveCategory({ name: "Cookies", pricing: "Per unit", unitPrice: 90 }); } catch (e) { err = e; }
check(err && /already/.test(err.message), "duplicate category refused");

console.log("\n=== dashboard + profit ===");
const d = store.dashboard(new Date("2026-11-30T12:00:00"));
check(d.month.byCategory.some((c) => c.category === "Granola bar") && d.month.byCategory.some((c) => c.category === "Cake"), "dashboard splits the month by category", d.month.byCategory);
const pc = store.profitByCategory();
check(pc.find((c) => c.category === "Cookies")?.qty === 20, "profit by category counts units sold", pc);
const formulasAfter = Object.values(wb.tabs).reduce((n, t) => n + t.formulas.flat().filter(Boolean).length, 0);
check(formulasAfter >= formulasBefore, "no formulas lost", [formulasBefore, formulasAfter]);

console.log("\n=== local mode (no sheet) ===");
mem.clear();
const local = new Store(seed, {});
check(local.activeCategories().length === 3, "local mode falls back to the default categories");
const lo = local.addOrder({ category: "Cookies", quantity: 3, customer: "X" });
check(lo.quoteSnapshot.finalQuote === 270 && !mem.get("bakery.v2.queue"), "cookie order priced locally, nothing queued");

await mock.close();
console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
