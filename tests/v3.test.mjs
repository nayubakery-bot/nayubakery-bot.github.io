// v3: everything editable from the app, written back to the sheet.
// Runs against the local mock of the Google Sheets REST API (never the real sheet),
// on a copy of the workbook reshaped like her sheet: a Calculator tab whose
// dropdowns/formulas read Settings ranges, with empty reserved rows inside the sizes list.
import { readFileSync } from "fs";
import { startMock, colIndex } from "./mock-sheets.mjs";

const mem = new Map();
globalThis.localStorage = { getItem: (k) => (mem.has(k) ? mem.get(k) : null), setItem: (k, v) => mem.set(k, String(v)), removeItem: (k) => mem.delete(k), clear: () => mem.clear() };
globalThis.sessionStorage = { getItem() { return null; }, setItem() {}, removeItem() {} };

const { Store } = await import("../js/store.js");
const { SyncEngine } = await import("../js/sync.js");
const { GoogleSheetsTransport, AppsScriptTransport } = await import("../js/transports.js");
const { planWrites, parseWorkbook, sectionLayout, settingsRefs } = await import("../js/sheet-model.js");
const seed = JSON.parse(readFileSync(new URL("../js/seed-data.json", import.meta.url)));

let failed = 0, passed = 0;
function check(cond, label, extra) {
  if (cond) { passed++; console.log("  ok  " + label); }
  else { failed++; console.error("  FAIL " + label, extra !== undefined ? JSON.stringify(extra) : ""); }
}
async function throws(fn) { try { await fn(); return null; } catch (e) { return e; } }

// ---------------- build "her" layout on the mock ----------------
const mock = await startMock();
const wb = mock.wb;
wb.tabs.Calculator = {
  name: "Calculator", sheetId: 777, maxRows: 100, maxCols: 26, formats: {},
  values: [["Custom Cake Pricing Calculator"], [], [], [], [], ["Cake size", "0.5 kg"], ["Weight (kg)", 0.5], ["Servings", "4-6"], ["Layers", 2]],
  formulas: [[], [], [], [], [], [], ["", "=INDEX(Settings!$B$35:$B$41,MATCH(B6,Settings!$A$35:$A$41,0))"], ["", "=INDEX(Settings!$C$35:$C$41,MATCH(B6,Settings!$A$35:$A$41,0))"], []],
};
wb.validationRefs = ["=Settings!$A$35:$A$41", "=Settings!$A$45:$A$48"];
// Reserve 3 empty rows inside the sizes list (Sheets-style insert: the Calculator ranges widen to A35:A44)
{
  const { default: _ } = { default: null };
  const t = wb.tabs.Settings;
  // use the mock's own insert via the REST endpoint below
}
async function rest(path, body) {
  const r = await fetch(mock.base + path, { method: body ? "POST" : "GET", headers: { Authorization: "Bearer " + mock.googleToken, "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
  return r.json();
}
await rest(`/v4/spreadsheets/${wb.id}:batchUpdate`, { requests: [{ insertDimension: { range: { sheetId: wb.tabs.Settings.sheetId, dimension: "ROWS", startIndex: 40, endIndex: 43 } } }] });
check(mock.formula("Calculator", "B7") === "=INDEX(Settings!$B$35:$B$44,MATCH(B6,Settings!$A$35:$A$44,0))", "fixture: Calculator B7 now reads Settings A35:A44 (3 reserved rows inside)", mock.formula("Calculator", "B7"));
check(mock.cell("Settings", "A44") === "5 kg" && mock.cell("Settings", "A41") === "", "fixture: rows 41–43 empty, 5 kg moved to row 44");
check(wb.validationRefs[1] === "=Settings!$A$48:$A$51", "fixture: layers dropdown range moved to A48:A51", wb.validationRefs);
const formulaCount = () => Object.values(wb.tabs).reduce((n, t) => n + t.formulas.flat().filter(Boolean).length, 0);
const formulasAtStart = formulaCount();

const transport = new GoogleSheetsTransport({ spreadsheetId: wb.id, getToken: () => mock.googleToken, base: mock.base });
let store;
const engine = new SyncEngine({ transport, debounceMs: 100000, onState: (g) => store.applySheetGrid(g), onFlushed: (d) => store.markFlushed(d) });
store = new Store(seed, { sync: engine });
const flush = async () => { const r = await engine.flush(); if (process.env.DBG) console.log("    queue after flush:", JSON.stringify(JSON.parse(localStorage.getItem("bakery.v2.queue")).map(q=>[q.kind,q.section||q.tab,q.action,q.key])), (r.cells||[]).filter(c=>c.keyCell).length, JSON.stringify((r.results||[]).filter(x=>x.skipped))); return r; };

console.log("\n=== load ===");
await engine.pull();
check(store.state.settings.sizes.length === 7 && store.state.settings.sizes.at(-1).name === "5 kg", "sizes parsed across the reserved blank rows (7, last is 5 kg)");
check(store.state.capacity.sizes.free === 3 && /A\$35:\$A\$44/.test(store.state.capacity.sizes.range), "capacity: 3 empty rows inside the Calculator range", store.state.capacity.sizes);
check(store.state.capacity.layers.free === 0, "layers list full (range A48:A51 is all used)", store.state.capacity.layers);
check(store.quote(seed.sampleOrder).finalQuote === 780, "sample quote still ₹780");

console.log("\n=== 1. menu categories ===");
store.addMenuItem("sizes", { name: "2.5 kg", kg: 2.5, servings: "20-25", sellingPrice: 2600 });
check(store.state.settings.sizes.some((z) => z.name === "2.5 kg"), "new size in the order form immediately (before sync)");
let r = await flush();
check(mock.cell("Settings", "A41") === "2.5 kg" && mock.cell("Settings", "B41") === 2.5 && mock.cell("Settings", "C41") === "20-25", "size written into first reserved row 41 (A41:C41)", [mock.cell("Settings", "A41"), mock.cell("Settings", "B41"), mock.cell("Settings", "C41")]);
check((mock.stats.addedSheets || []).includes("Price List") && mock.cell("Price List", "A1") === "Kind" && mock.cell("Price List", "C1") === "Selling price (₹)", "Price List tab created with headers");
check(mock.cell("Price List", "A2") === "Size" && mock.cell("Price List", "B2") === "2.5 kg" && mock.cell("Price List", "C2") === 2600, "selling price row written");
store.addMenuItem("sizes", { name: "3.5 kg", kg: 3.5, servings: "28-35" });
store.addMenuItem("sizes", { name: "6 kg", kg: 6, servings: "48-60" });
let e = await throws(() => store.addMenuItem("sizes", { name: "7 kg", kg: 7 }));
check(e && e.full, "4th new size refused: list full in the sheet", e && e.message);
await flush();
check(mock.cell("Settings", "A42") === "3.5 kg" && mock.cell("Settings", "A43") === "6 kg" && mock.cell("Settings", "A44") === "5 kg", "rows 42, 43 used; 5 kg untouched at 44");
check(mock.cell("Settings", "A45") === "", "nothing written below the range");

store.makeRoom("sizes");
store.addMenuItem("sizes", { name: "7 kg", kg: 7, servings: "56-70" });
await flush();
check((mock.stats.insertedRows || []).some((x) => x.tab === "Settings" && x.at === 43 && x.n === 5), "Make room inserted 5 rows before the range's last row", mock.stats.insertedRows);
check(mock.formula("Calculator", "B7") === "=INDEX(Settings!$B$35:$B$49,MATCH(B6,Settings!$A$35:$A$49,0))", "Calculator ranges widened to A35:A49", mock.formula("Calculator", "B7"));
check(mock.cell("Settings", "A44") === "7 kg" && mock.cell("Settings", "A49") === "5 kg", "7 kg landed inside the widened range");
await engine.pull();
check(store.state.settings.sizes.length === 11 && store.state.capacity.sizes.free === 4, "reload: 11 sizes, 4 free rows", [store.state.settings.sizes.length, store.state.capacity.sizes]);

e = await throws(() => store.addMenuItem("layers", { value: 6 }));
check(e && e.full, "layers full -> clear message", e && e.message);

// rename + remove
store.updateMenuItem("sizes", "6 kg", { name: "6 kg (party)", servings: "50-60" });
store.updateMenuItem("sizes", "2.5 kg", { name: "2.5 kg tall" });
await flush();
check(mock.cell("Settings", "A43") === "6 kg (party)" && mock.cell("Settings", "C43") === "50-60", "rename + servings edit -> A43, C43");
check(mock.cell("Price List", "B2") === "2.5 kg tall", "price list follows the rename");
store.removeMenuItem("sizes", "3.5 kg");
await flush();
check(mock.cell("Settings", "A42") === "" && mock.cell("Settings", "B42") === "" && mock.cell("Settings", "C42") === "", "remove clears A42:C42 (row stays, no shifting)");
await engine.pull();
check(!store.state.settings.sizes.some((z) => z.name === "3.5 kg") && store.state.capacity.sizes.free === 5, "removed size gone after reload; its row is free again");

// toppings (no Calculator range, no spare rows) -> make room, new raw material
e = await throws(() => store.addMenuItem("toppings", { name: "Gold sprinkles", material: "Gold sprinkles", newMaterial: { packSize: 50, unit: "g", packPrice: 300 }, fixedQty: 5, qtyPerKg: 0 }));
check(e && e.full, "toppings list full (no blank rows before the next section)");
store.makeRoom("toppings");
store.addMenuItem("toppings", { name: "Gold sprinkles", material: "Gold sprinkles", newMaterial: { packSize: 50, unit: "g", packPrice: 300 }, fixedQty: 5, qtyPerKg: 0, sellingPrice: 80 });
await flush();
const tRow = wb.tabs.Settings.values.findIndex((row) => row[0] === "Gold sprinkles") + 1;
check(tRow > 0 && mock.cell("Settings", "B" + tRow) === "Gold sprinkles" && mock.cell("Settings", "D" + tRow) === 5, "topping row written", tRow);
check(mock.formula("Settings", "F" + tRow).includes(`MATCH(B${tRow},`) && mock.formula("Settings", "G" + tRow) === `=N(D${tRow})*F${tRow}`, "topping cost formulas copied to its own row", [mock.formula("Settings", "F" + tRow), mock.formula("Settings", "G" + tRow)]);
const rmRow = wb.tabs["Raw Materials"].values.findIndex((row) => row[1] === "Gold sprinkles") + 1;
check(rmRow > 0 && mock.cell("Raw Materials", "E" + rmRow) === 300, "new raw material created for the topping", rmRow);

// boxes + add-ons
store.makeRoom("boxes");
store.addMenuItem("boxes", { name: "Cake box 6x6x4 in", newMaterial: { packSize: 10, unit: "nos", packPrice: 180 }, sellingPrice: 30 });
store.makeRoom("addons");
store.addMenuItem("addons", { name: "Sparkler candle", material: "Sparkler candle", newMaterial: { packSize: 10, unit: "nos", packPrice: 100 }, qty: 1, sellingPrice: 25 });
await flush();
const bRow = wb.tabs.Settings.values.findIndex((row) => row[0] === "Cake box 6x6x4 in") + 1;
check(bRow > 0 && mock.formula("Settings", "B" + bRow).includes(`MATCH(A${bRow},`), "new box row + its cost formula", [bRow, mock.formula("Settings", "B" + bRow)]);
await engine.pull();
check(store.state.settings.boxes.includes("Cake box 6x6x4 in") && store.state.settings.addons.some((a) => a.name === "Sparkler candle"), "box + add-on back from the sheet");
const withExtra = store.quote({ ...seed.sampleOrder, extraAddons: ["Sparkler candle"] });
check(Math.abs(withExtra.lines.extraAddons[0].cost - 10) < 1e-9, "extra add-on costed from its raw material (₹10)", withExtra.lines.extraAddons);

console.log("\n=== 2. selling prices + override ===");
store.setSellingPrice("Size", "0.5 kg", 700);
await flush();
const mq = store.quote(seed.sampleOrder);
check(mq.menu && mq.menu.parts[0].price === 700, "quote uses her menu price for 0.5 kg");
// 700 + box 25*1.35 + board 18*1.35 + bag 15*1.35 = 778.3 ; + delivery 128 = 906.3 -> 910
check(mq.finalQuote === 910, "menu quote = 700 + extras at cost+markup + delivery, rounded (₹910)", mq.finalQuote);
store.setSellingPrice("Size", "0.5 kg", "");
await flush();
check(store.quote(seed.sampleOrder).finalQuote === 780, "clearing the price -> back to cost-plus ₹780");
const priceRows = wb.tabs["Price List"].values.filter((row) => row[1] === "0.5 kg" && row[2] !== "");
check(priceRows.length === 0, "cleared price removed from Price List tab");
const ov = store.addOrder({ ...seed.sampleOrder, customer: "Override Meera", phone: "98450 11111", priceOverride: 900 });
await flush();
const ovRow = wb.tabs.Orders.values.findIndex((row) => row[0] === ov.id) + 1;
const hdr = wb.tabs.Orders.values[0];
const col = (h) => String.fromCharCode(65 + hdr.indexOf(h)).replace(/^/, "") ;
const colA1 = (h) => { let i = hdr.indexOf(h) + 1, s = ""; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
check(hdr.includes("Price override (₹)") && hdr.includes("Payment status") && hdr.includes("Balance (₹)") && hdr.includes("Delivery zone") && hdr.includes("Rules applied") && hdr.includes("Other add-ons"), "Orders got the new columns at the end", hdr.slice(29));
check(mock.cell("Orders", colA1("Price override (₹)") + ovRow) === 900 && mock.cell("Orders", colA1("Quote (₹)") + ovRow) === 900, "override written; Quote = ₹900");

console.log("\n=== 3. price went up ===");
const im = store.priceImpact("Cocoa powder", 600);
check(im && im.recipes.some((x) => x.name === "Chocolate Sponge" && x.after > x.before), "impact lists Chocolate Sponge cost rising", im && im.recipes);
const before = store.costOf("Chocolate Sponge");
store.priceWentUp("Cocoa powder", 600);
await flush();
const cRow = wb.tabs["Raw Materials"].values.findIndex((row) => row[1] === "Cocoa powder") + 1;
check(mock.cell("Raw Materials", "E" + cRow) === 600 && store.costOf("Chocolate Sponge") > before, "new pack price in sheet; recipe cost updated");

console.log("\n=== 4. recipes ===");
e = await throws(() => store.createRecipe({ name: "Butterscotch Sponge", type: "Sponge", lines: [{ ingredient: "Maida (refined flour)", qtyPerKg: 250 }] }));
check(e && e.full, "recipe menu full -> asks to make room");
store.makeRoom("recipeMenu", "Sponge");
store.createRecipe({ name: "Butterscotch Sponge", type: "Sponge", lines: [{ ingredient: "Maida (refined flour)", qtyPerKg: 250 }, { ingredient: "Sugar (granulated)", qtyPerKg: 200 }], method: "Cream butter and sugar. Bake 170°C 35 min." });
await flush();
const mRow = wb.tabs.Settings.values.findIndex((row) => row[1] === "Butterscotch Sponge") + 1;
check(mRow > 0 && mock.cell("Settings", "A" + mRow) === "Sponge" && mock.formula("Settings", "C" + mRow).includes(`B${mRow}`), "menu row + cost-per-kg formula", [mRow, mock.formula("Settings", "C" + mRow)]);
const lines = wb.tabs.Recipes.values.filter((row) => row[0] === "Butterscotch Sponge");
check(lines.length === 2, "2 recipe lines appended");
check(mock.cell("Recipe Notes", "A2") === "Butterscotch Sponge" && /170°C/.test(mock.cell("Recipe Notes", "C2")), "Recipe Notes tab created with the method");
await engine.pull();
check(store.state.settings.recipeMenu.some((m) => m.name === "Butterscotch Sponge" && m.type === "Sponge"), "new sponge in the order form after reload");
const line = store.state.recipes.find((l) => l.name === "Butterscotch Sponge" && l.ingredient === "Sugar (granulated)");
store.removeRecipeLine(line.key);
await flush();
check(wb.tabs.Recipes.values.filter((row) => row[0] === "Butterscotch Sponge").length === 1, "recipe line removed (cells cleared)");
await engine.pull();
check(store.state.recipes.filter((l) => l.name === "Butterscotch Sponge").length === 1, "removed line stays gone after reload");
store.setRecipeNotes("Butterscotch Sponge", "Updated notes", "Sponge");
await flush();
check(mock.cell("Recipe Notes", "C2") === "Updated notes" && wb.tabs["Recipe Notes"].values.length === 2, "notes updated in place");

console.log("\n=== 5. inventory ===");
const eggs = store.inventoryRows().find((i) => i.item === "Eggs");
store.logPurchase({ item: "Eggs", packs: 2, pricePerPack: 220, notes: "Ratnadeep" });
await flush();
check(mock.cell("Inventory", "E10") === Number(eggs.purchased) + 60, "purchase adds 2×30 eggs to Purchased", mock.cell("Inventory", "E10"));
check(mock.cell("Raw Materials", "E10") === 220, "pack price updated to ₹220");
check(mock.cell("Purchases", "B2") === "Eggs" && mock.cell("Purchases", "D2") === 60 && mock.cell("Purchases", "G2") === 440, "Purchases log row");
store.updateInventory("Eggs", { reorderLevel: 24 });
await flush();
check(mock.cell("Inventory", "H10") === 24, "low-stock threshold -> Reorder level column");

console.log("\n=== 6–7. rules + zones ===");
const rush = store.saveRule({ name: "Express (24 h)", effect: "Charge", type: "%", amount: 20, active: true, applyTo: "When ticked" });
const diwali = store.saveRule({ name: "Diwali", effect: "Discount", type: "₹", amount: 50, active: true, applyTo: "Every order" });
const zone = store.saveZone({ name: "Within 5 km", charge: 60 });
await flush();
check(mock.cell("Pricing Rules", "B2") === "Express (24 h)" && mock.cell("Pricing Rules", "F2") === "Yes" && mock.cell("Pricing Rules", "B3") === "Diwali", "Pricing Rules tab rows");
check(mock.cell("Delivery Zones", "B2") === "Within 5 km" && mock.cell("Delivery Zones", "C2") === 60, "Delivery Zones tab row");
check(JSON.stringify(store.defaultRules()) === JSON.stringify(["Diwali"]), "Every-order rule pre-ticked on new orders");
const rq = store.quote({ ...seed.sampleOrder, rules: ["Express (24 h)", "Diwali"], zone: "Within 5 km" });
const base = store.quote(seed.sampleOrder).priceBeforeDelivery;
const want = Math.ceil((base * 1.2 - 50 + 60) / 10 - 1e-12) * 10;
check(rq.delivery === 60 && rq.finalQuote === want && rq.rules.length === 2 && Math.abs(rq.profit - (base * 1.2 - 50 - rq.totalCost)) < 1e-6, `rush +20%, Diwali −₹50, zone ₹60 -> ₹${want}; extra goes to profit`, [rq.delivery, rq.finalQuote, want]);
store.saveRule({ ...rush, active: false });
await flush();
check(mock.cell("Pricing Rules", "F2") === "No", "toggle off -> Active No");

console.log("\n=== 8–9. customers, payments, status ===");
const o = store.addOrder({ ...seed.sampleOrder, customer: "Asha Rao", phone: "+91 98480 22222", rules: ["Diwali"], zone: "Within 5 km" });
await flush();
check(wb.tabs.Customers.values.some((row) => row[1] === "Asha Rao" && row[2] === "+91 98480 22222"), "customer saved automatically from the order");
const asha = store.findCustomer({ phone: "9848022222" });
check(asha && asha.name === "Asha Rao", "quick-fill finds her by phone (any format)");
const o2 = store.addOrder({ ...seed.sampleOrder, customer: "Asha Rao", phone: "9848022222" });
await flush();
check(wb.tabs.Customers.values.filter((row) => row[1] === "Asha Rao").length === 1, "no duplicate customer");
check(store.customerRows().find((c) => c.name === "Asha Rao").orders === 2, "customer shows 2 past orders");
const oRow = wb.tabs.Orders.values.findIndex((row) => row[0] === o.id) + 1;
check(mock.cell("Orders", colA1("Rules applied") + oRow) === "Diwali" && mock.cell("Orders", colA1("Delivery zone") + oRow) === "Within 5 km", "rules + zone stored on the order");
store.recordPayment(o.id, 300);
await flush();
check(mock.cell("Orders", colA1("Advance paid (₹)") + oRow) === 300 && mock.cell("Orders", colA1("Payment status") + oRow) === "Advance paid", "advance recorded", [mock.cell("Orders", colA1("Advance paid (₹)") + oRow), mock.cell("Orders", colA1("Payment status") + oRow)]);
const total = store.state.orders.find((x) => x.id === o.id).quoteSnapshot.finalQuote;
check(mock.cell("Orders", colA1("Balance (₹)") + oRow) === total - 300, "balance written", [mock.cell("Orders", colA1("Balance (₹)") + oRow), total]);
store.setStatus(o.id, "Delivered");
store.markPaidInFull(o.id);
await flush();
check(mock.cell("Orders", colA1("Status") + oRow) === "Paid" && mock.cell("Orders", colA1("Payment status") + oRow) === "Paid in full" && mock.cell("Orders", colA1("Balance (₹)") + oRow) === 0, "paid in full after delivery -> Status Paid, balance 0");
store.removeCustomer(store.findCustomer({ name: "Override Meera" }).id);
await flush();
check(!wb.tabs.Customers.values.some((row) => row[1] === "Override Meera"), "customer removed from Customers tab");

console.log("\n=== reload: everything comes back from the sheet ===");
const st0 = JSON.parse(JSON.stringify(store.state));
mem.delete("bakery.v2.sheetCache");
await engine.pull();
const st = store.state;
check(st.rules.length === 2 && st.rules.find((x) => x.name === "Diwali").type === "₹" && st.rules.find((x) => x.name === "Express (24 h)").active === false, "rules reloaded");
check(st.zones.length === 1 && st.zones[0].charge === 60, "zones reloaded");
check(st.customers.length === st0.customers.length && st.customers.some((c) => c.name === "Asha Rao"), "customers reloaded");
check(st.priceList.length === st0.priceList.length, "price list reloaded", [st.priceList, st0.priceList]);
check(st.recipeNotes.length === 1 && st.purchases.length === 1, "recipe notes + purchases reloaded");
const ro = st.orders.find((x) => x.id === o.id);
check(ro.zone === "Within 5 km" && ro.rules[0] === "Diwali" && ro.advancePaid === total, "order zone/rules/payment reloaded");
check(st.orders.find((x) => x.id === ov.id).priceOverride === 900, "price override reloaded");
check(store.quote(seed.sampleOrder).finalQuote > 780, "quote follows the higher cocoa price after reload");

console.log("\n=== safety ===");
check(formulaCount() >= formulasAtStart, "no formula cell lost", [formulasAtStart, formulaCount()]);
check(engine.status.conflicts.filter((c) => !c.note).length === 0, "no change refused", engine.status.conflicts);
check(JSON.parse(localStorage.getItem("bakery.v2.queue")).length === 0, "queue empty");

console.log("\n=== offline: new categories wait on the phone ===");
mock.setDown(true);
store.addMenuItem("sizes", { name: "8 kg", kg: 8, servings: "64-80" });
store.saveZone({ name: "Secunderabad", charge: 120 });
await engine.flush().catch(() => {});
check(engine.status.state === "offline" && JSON.parse(localStorage.getItem("bakery.v2.queue")).length === 2, "2 changes queued offline");
mock.setDown(false);
await engine.pull();
check(store.state.settings.sizes.some((z) => z.name === "8 kg") && wb.tabs["Delivery Zones"].values.some((row) => row[1] === "Secunderabad"), "written once back online");
await mock.close();

console.log("\n=== pure: per-type Calculator ranges for the recipe menu ===");
{
  const t = { values: [["Type", "Recipe name", "Cost per kg (₹)"], ["Sponge", "A"], ["Sponge", "B"], ["Filling", "C"], ["Filling", ""], ["Frosting", "D"]], formulas: [] };
  t.values[4] = ["", ""];
  const grid = { tabs: { Settings: t, Calculator: { values: [], formulas: [["=Settings!$B$2:$B$3", "=Settings!$B$4:$B$5", "=Settings!$B$6:$B$6"]] } } };
  const plan = planWrites(grid, [
    { kind: "menu", section: "recipeMenu", action: "add", key: "New sponge", fields: { type: "Sponge" }, queuedAt: "a" },
    { kind: "menu", section: "recipeMenu", action: "add", key: "New filling", fields: { type: "Filling" }, queuedAt: "b" },
  ]);
  check(plan.conflicts.length === 1 && /full/.test(plan.conflicts[0].reason) && plan.conflicts[0].op.key === "New sponge", "sponge range B2:B3 full -> refused", plan.conflicts);
  check(plan.cells.some((c) => c.row === 5 && c.col === 2 && c.value === "New filling"), "filling goes into the empty row 5 inside its own range B4:B5");
}

console.log("\n=== local mode (no sheet) ===");
{
  mem.clear();
  const ls = new Store(seed, { sync: new SyncEngine({ transport: null }) });
  ls.addMenuItem("sizes", { name: "2.5 kg", kg: 2.5, servings: "20" });
  ls.addMenuItem("toppings", { name: "Gold", material: "Gold leaf", newMaterial: { packSize: 10, unit: "nos", packPrice: 250 }, fixedQty: 1 });
  ls.saveRule({ name: "Rush", effect: "Charge", type: "₹", amount: 100, active: true, applyTo: "When ticked" });
  ls.saveZone({ name: "Near", charge: 40 });
  ls.setSellingPrice("Size", "2.5 kg", 2500);
  const o = ls.addOrder({ ...seed.sampleOrder, customer: "Local Lata", phone: "9000000001", rules: ["Rush"], zone: "Near" });
  ls.recordPayment(o.id, 200);
  const saved = JSON.parse(localStorage.getItem("bakery.v1.state"));
  check(saved.settings.sizes.some((z) => z.name === "2.5 kg") && saved.rules.length === 1 && saved.zones.length === 1 && saved.customers.length === 1 && saved.priceList.length === 1, "everything kept on the phone in local mode");
  check(ls.quote({ ...seed.sampleOrder, size: "2.5 kg" }).menu?.parts[0].price === 2500, "menu price used in local mode");
  check(!localStorage.getItem("bakery.v2.queue") || JSON.parse(localStorage.getItem("bakery.v2.queue")).length === 0, "nothing queued in local mode");
}

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
