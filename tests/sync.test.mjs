// End-to-end sync tests against the local mock (never the real Google Sheet).
// Runs every scenario twice: Apps Script transport and direct Sheets API transport.
import { readFileSync } from "fs";
import { startMock } from "./mock-sheets.mjs";

// --- browser shims ---
const mem = new Map();
globalThis.localStorage = {
  getItem: (k) => (mem.has(k) ? mem.get(k) : null),
  setItem: (k, v) => mem.set(k, String(v)),
  removeItem: (k) => mem.delete(k),
  clear: () => mem.clear(),
};
globalThis.sessionStorage = { getItem() { return null; }, setItem() {}, removeItem() {} };

const { Store } = await import("../js/store.js");
const { SyncEngine } = await import("../js/sync.js");
const { AppsScriptTransport, GoogleSheetsTransport } = await import("../js/transports.js");
const seed = JSON.parse(readFileSync(new URL("../js/seed-data.json", import.meta.url)));

let failed = 0, passed = 0;
function check(cond, label, extra) {
  if (cond) { passed++; console.log("  ok  " + label); }
  else { failed++; console.error("  FAIL " + label, extra !== undefined ? JSON.stringify(extra) : ""); }
}

function snapshotFormulas(wb) {
  const out = {};
  for (const [n, t] of Object.entries(wb.tabs)) t.formulas.forEach((row, r) => row.forEach((f, c) => { if (f) out[`${n}!${r},${c}`] = f; }));
  return out;
}

async function scenario(mode) {
  console.log(`\n=== transport: ${mode} ===`);
  mem.clear();
  const mock = await startMock();
  // an extra column she might add herself, to prove unknown columns survive
  mock.wb.tabs["Raw Materials"].values[3][8] = "My extra col";
  mock.wb.tabs["Raw Materials"].values[4][8] = "keep me";
  const formulasBefore = snapshotFormulas(mock.wb);

  const transport = mode === "appsscript"
    ? new AppsScriptTransport({ url: mock.execUrl, token: "2468" })
    : new GoogleSheetsTransport({ spreadsheetId: mock.wb.id, getToken: () => mock.googleToken, base: mock.base });

  let store;
  const engine = new SyncEngine({
    transport,
    debounceMs: 5,
    onState: (grid) => store.applySheetGrid(grid),
    onFlushed: (done) => store.markFlushed(done),
  });
  store = new Store(seed, { sync: engine });

  // ---------- 1. sheet -> app on load ----------
  await engine.pull();
  check(engine.status.state === "synced", "initial pull synced", engine.status);
  check(store.state.rawMaterials.length === 46 && store.state.recipes.length === 88 && store.state.inventory.length === 46, "all rows mapped (46 materials, 88 recipe lines, 46 stock rows)");
  check(store.state.library.length === 6 && store.state.orders.length === 0, "library 6, orders 0 (instruction row ignored)");
  const sample = seed.sampleOrder;
  check(store.quote(sample).finalQuote === 780, "0.5 kg chocolate birthday cake quotes ₹780 from live sheet values", store.quote(sample).finalQuote);
  check(store.state.settings.sizes[0].servings === "4-6", "servings 46118 (date-mangled) read back as 4-6");

  // ---------- 2. sheet -> app after hand edits + reload ----------
  mock.handEdit("Raw Materials", "E5", 60);        // Maida 52 -> 60
  mock.handEdit("Settings", "B29", 0.4);           // markup 35% -> 40%
  mock.handEdit("Inventory", "E5", 1000);          // purchased
  mock.handEdit("Recipes", "D20", 220);            // chocolate sponge maida
  await engine.pull();                             // = page reload / refresh button
  const maida = store.state.rawMaterials.find((m) => m.name === "Maida (refined flour)");
  check(maida.packPrice === 60, "sheet price edit visible after reload");
  check(store.state.settings.markup === 0.4, "sheet settings edit visible after reload");
  check(store.inventoryRows()[0].current === 6000, "sheet inventory edit visible (5000+1000)");
  check(store.state.recipes.find((r) => r.name === "Chocolate Sponge" && r.ingredient === "Maida (refined flour)").qtyPerKg === 220, "sheet recipe edit visible");
  check(store.quote(sample).finalQuote !== 780, "quote follows live sheet values", store.quote(sample).finalQuote);
  // put them back by hand
  mock.handEdit("Raw Materials", "E5", 52); mock.handEdit("Settings", "B29", 0.35);
  mock.handEdit("Inventory", "E5", 0); mock.handEdit("Recipes", "D20", 209.9);
  await engine.pull();
  check(store.quote(sample).finalQuote === 780, "back to ₹780 after reverting in sheet");

  // ---------- 3. app -> sheet ----------
  const o = store.addOrder({ ...sample, customer: "Test Priya", phone: "9999900000", deliveryDate: "2026-11-15", status: "Confirmed", advancePaid: 300 });
  await engine.flush();
  check(mock.cell("Orders", "A3") === o.id, "new order appended as row 3 (instruction row 2 kept)", mock.cell("Orders", "A3"));
  check(String(mock.cell("Orders", "A2")).startsWith("Paste or import"), "instruction row untouched");
  check(mock.cell("Orders", "C3") === "Test Priya" && mock.cell("Orders", "W3") === 780, "order customer + Quote (₹) 780 written", [mock.cell("Orders", "C3"), mock.cell("Orders", "W3")]);
  check(mock.cell("Orders", "S3") === "2026-11-15", "delivery date stored as typed");
  check(mock.cell("Orders", "AD1") === "Updated at" && mock.cell("Orders", "AE1") === "Hours override", "Updated at + Hours override columns added at the END (AD, AE)");
  check(String(mock.cell("Orders", "AD3")).length > 10, "updatedAt stamped");
  check(String(mock.cell("Recipes", "H4")) === "Line ID" && /^R0005-/.test(mock.cell("Recipes", "H5")), "Recipes got a Line ID column at the end (H)");

  store.updateOrder(o.id, { status: "Baking", notes: "Happy birthday Asha" });
  await engine.flush();
  check(mock.cell("Orders", "T3") === "Baking" && mock.cell("Orders", "V3") === "Happy birthday Asha", "order edit -> exact cells T3, V3");

  store.deleteOrder(o.id);
  await engine.flush();
  check(mock.cell("Orders", "T3") === "Deleted" && mock.cell("Orders", "A3") === o.id, "delete = soft (Status Deleted, row kept)");
  check(store.visibleOrders.length === 0, "deleted order hidden in app");

  store.updateMaterial("Butter (unsalted)", { packPrice: 300 });
  await engine.flush();
  check(mock.cell("Raw Materials", "E8") === 300, "price edit -> Raw Materials!E8");
  check(mock.formula("Raw Materials", "F8").startsWith("=IF(AND(ISNUMBER(C8)"), "Cost per unit formula F8 untouched");

  store.updateInventory("Eggs", { used: 12, reorderLevel: 15 });
  await engine.flush();
  check(mock.cell("Inventory", "F10") === 12 && mock.cell("Inventory", "H10") === 15, "inventory edit -> Inventory!F10, H10");

  store.updateSettings({ labourRate: 150, overheads: { gas: 25 } });
  await engine.flush();
  check(mock.cell("Settings", "B11") === 150 && mock.cell("Settings", "B18") === 25, "settings -> Settings!B11, B18");
  check(mock.formula("Settings", "B23") === "=SUM(B18:B22)", "Total overhead formula B23 untouched");

  const line = store.state.recipes.find((r) => r.name === "Whipped Cream" && r.ingredient === "Icing sugar");
  store.updateRecipeLine(line.key, { qtyPerKg: 110 });
  await engine.flush();
  check(mock.cell("Recipes", "D50") === 110, "recipe qty edit -> Recipes!D50");

  store.addRecipeLine({ name: "Whipped Cream", type: "Filling", ingredient: "Salt", qtyPerKg: 1 });
  await engine.flush();
  check(mock.cell("Recipes", "A93") === "Whipped Cream" && mock.cell("Recipes", "C93") === "Salt" && mock.cell("Recipes", "D93") === 1, "new recipe line appended at row 93");
  check(mock.formula("Recipes", "G93") === '=IF(OR(C93="",NOT(ISNUMBER(D93))),0,D93*F93)', "row 93 got the line-cost formula shifted to its own row", mock.formula("Recipes", "G93"));

  store.addMaterial({ category: "Toppings", name: "Gold leaf", packSize: 10, unit: "nos", packPrice: 250, notes: "", estimated: false });
  await engine.flush();
  check(mock.cell("Raw Materials", "B51") === "Gold leaf" && mock.cell("Raw Materials", "E51") === 250, "new material appended at Raw Materials row 51");
  check(mock.formula("Raw Materials", "F51").includes("E51/C51"), "row 51 cost-per-unit formula created");
  check(mock.formula("Inventory", "A51") === `=IF('Raw Materials'!B51="","",'Raw Materials'!B51)`, "Inventory row 51 created, mirroring Raw Materials row 51", mock.formula("Inventory", "A51"));
  check(mock.cell("Inventory", "D51") === 0, "Inventory row 51 opening stock 0");

  // batch scaler / calculator are computed views: nothing written by them
  const before = mock.stats.writeBodies.length;
  store.quote({ ...sample, size: "2 kg" });
  check(mock.stats.writeBodies.length === before, "calculator/batch scaler never write");

  // ---------- 4. safety ----------
  const formulasAfter = snapshotFormulas(mock.wb);
  const changed = Object.keys(formulasBefore).filter((k) => formulasAfter[k] !== formulasBefore[k]);
  check(changed.length === 0, "no pre-existing formula cell was changed", changed.slice(0, 5));
  check(mock.cell("Raw Materials", "I4") === "My extra col" && mock.cell("Raw Materials", "I5") === "keep me", "unknown extra column survived");
  // direct attempt to hit a formula cell is refused by the script / transport guard
  const r = await transport.write([{ tab: "Raw Materials", row: 5, col: 6, value: 999 }], await transport.read());
  check(r.results[0].skipped === "formula" && mock.formula("Raw Materials", "F5"), "write to formula cell refused");
  if (mode === "appsscript") {
    const bad = await new AppsScriptTransport({ url: mock.execUrl, token: "0000" }).read().catch((e) => e);
    check(bad.kind === "auth", "wrong token rejected");
    check(mock.stats.options === 0, "no CORS preflight sent (text/plain POST)");
  } else {
    const bad = await new GoogleSheetsTransport({ spreadsheetId: mock.wb.id, getToken: () => "nope", base: mock.base }).read().catch((e) => e);
    check(bad.kind === "auth", "bad Google token rejected");
  }

  // ---------- 5. offline queue ----------
  mock.setDown(true);
  store.updateMaterial("Sugar (granulated)", { packPrice: 85 });
  const o2 = store.addOrder({ ...sample, customer: "Offline Ravi" });
  await engine.flush().catch(() => {});
  check(engine.status.state === "offline", "status Offline when the sheet can't be reached", engine.status.state);
  const q = JSON.parse(localStorage.getItem("bakery.v2.queue"));
  check(q.length === 3 && q.filter((x) => x.kind === "row" && x.tab === "Customers").length === 1, "3 edits waiting in the localStorage queue (price, order, new customer)", q.map((x) => [x.kind, x.tab, x.key, x.field, x.id]));
  // simulate closing and reopening the app: new engine + store from storage
  engine.stop(); // the old page is closed
  mock.setDown(false);
  let store2;
  const engine2 = new SyncEngine({ transport, debounceMs: 5, onState: (g) => store2.applySheetGrid(g), onFlushed: (d) => store2.markFlushed(d) });
  store2 = new Store(seed, { sync: engine2 });
  check(store2.state.orders.some((x) => x.id === o2.id), "offline order still on the phone after restart");
  await engine2.pull();
  check(mock.cell("Raw Materials", "E6") === 85, "queued price written once back online");
  check(mock.cell("Orders", "A4") === o2.id, "queued order appended (row 4)");
  check(JSON.parse(localStorage.getItem("bakery.v2.queue")).length === 0 && engine2.status.state === "synced", "queue empty, status Synced", [engine2.status, JSON.parse(localStorage.getItem("bakery.v2.queue"))]);

  // ---------- 6. conflicts ----------
  // Orders: last write wins by Updated at. She edits the order in the sheet AFTER the app's offline edit.
  mock.setDown(true);
  store2.updateOrder(o2.id, { notes: "app edit (older)" });
  await engine2.flush().catch(() => {});
  await new Promise((r) => setTimeout(r, 5));
  mock.setDown(false);
  mock.handEdit("Orders", "V4", "sheet edit (newer)"); // onEdit stamps Updated at = now
  if (mode !== "appsscript") {
    // direct-API mode has no onEdit trigger; stamp by hand like the script would
    mock.wb.tabs.Orders.values[3][29] = new Date(Date.now() + 1000).toISOString();
  }
  await engine2.flush();
  check(mock.cell("Orders", "V4") === "sheet edit (newer)", "newer sheet edit wins over older app edit");
  check(engine2.status.conflicts.some((c) => /more recently/.test(c.reason)), "conflict reported to the app");
  // Other tabs: if the cell changed in the sheet since the app loaded it, the sheet value is kept.
  mock.handEdit("Raw Materials", "E7", 77);
  store2.updateMaterial("Icing sugar", { packPrice: 70 });
  await engine2.flush();
  check(mock.cell("Raw Materials", "E7") === 77, "price changed in sheet meanwhile -> sheet value kept");

  // ---------- 7. two phones add an order at the same moment ----------
  const mkPhone = () => {
    let st;
    const en = new SyncEngine({ transport, debounceMs: 100000, onState: (g) => st.applySheetGrid(g), onFlushed: (d) => st.markFlushed(d) });
    st = new Store(seed, { sync: en });
    return { st, en };
  };
  mem.delete("bakery.v2.queue");
  const A = mkPhone();
  const oa = A.st.addOrder({ ...sample, customer: "Phone A" });
  A.en.stop();
  mem.delete("bakery.v2.queue");
  const B = mkPhone();
  B.en.queue = [];
  const ob = B.st.addOrder({ ...sample, customer: "Phone B" });
  B.en.stop();
  await Promise.all([A.en.flush(), B.en.flush()]);
  await new Promise((r) => setTimeout(r, 1700)); // the loser retries
  const ids = mock.wb.tabs.Orders.values.map((r) => r[0]);
  check(ids.includes(oa.id) && ids.includes(ob.id), "both simultaneous orders landed in their own rows", ids.slice(2));
  const rowA = ids.indexOf(oa.id), rowB = ids.indexOf(ob.id);
  check(mock.wb.tabs.Orders.values[rowA][2] === "Phone A" && mock.wb.tabs.Orders.values[rowB][2] === "Phone B", "no mixing of the two rows");

  await mock.close();
}

// local mode (no sync configured) behaves like v1
{
  console.log("\n=== local mode ===");
  mem.clear();
  const st = new Store(seed, { sync: new SyncEngine({ transport: null }) });
  const o = st.addOrder({ ...seed.sampleOrder, customer: "Local" });
  check(st.quote(seed.sampleOrder).finalQuote === 780, "local mode quote ₹780 (seed)");
  check(JSON.parse(localStorage.getItem("bakery.v1.state")).orders.length === 1, "local order saved in localStorage");
  check(!localStorage.getItem("bakery.v2.queue") || JSON.parse(localStorage.getItem("bakery.v2.queue")).length === 0, "nothing queued in local mode");
  st.deleteOrder(o.id);
  check(st.state.orders.length === 0, "local delete removes the order");
}

await scenario("appsscript");
await scenario("google");
console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
