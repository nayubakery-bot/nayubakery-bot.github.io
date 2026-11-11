// Sheet model: maps the Bakery Workbook tabs (by header NAME, not position)
// to the app's data model, and turns app edits into cell-level writes.
// Pure functions, no DOM, shared by both sync transports (Apps Script and
// direct Google Sheets API) and by the Node tests.
//
// Grid shape (what a transport's read() returns):
//   { tabs: { [tabName]: { values: any[][], formulas: string[][] } },
//     meta: { [tabName]: { sheetId, rowCount, columnCount } } }
// Rows/cols in `values` are 0-based; cell writes use 1-based row/col.

export const TABS = ["Raw Materials", "Recipes", "Settings", "Inventory", "Orders", "Recipe Library"];

export function norm(h) {
  return String(h ?? "")
    .toLowerCase()
    .replace(/₹/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

const blank = (v) => v === "" || v === null || v === undefined;
const num = (v) => {
  if (blank(v)) return 0;
  const x = Number(String(v).replace(/[₹,\s]/g, ""));
  return isFinite(x) ? x : 0;
};
const str = (v) => (blank(v) ? "" : String(v));

// Google/Excel serial date -> yyyy-mm-dd
export function serialToISO(n) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(Number(n) * 86400000));
  return d.toISOString().slice(0, 10);
}
function toDate(v) {
  if (blank(v)) return "";
  if (typeof v === "number" && v > 20000 && v < 80000) return serialToISO(v);
  return String(v);
}
function toBool(v) {
  if (v === true) return true;
  return /^(yes|y|true|1|✓)$/i.test(String(v ?? "").trim());
}
// Servings like "4-6" that Sheets auto-converted to a date (46118 = 6 Apr).
function fixServings(v) {
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return `${d.getUTCMonth() + 1}-${d.getUTCDate()}`;
  }
  const iso = String(v ?? "").match(/^\d{4}-(\d{2})-(\d{2})/);
  if (iso) return `${Number(iso[1])}-${Number(iso[2])}`;
  return str(v);
}

// ---------- column specs (header aliases are normalized) ----------
// ro: computed in the sheet (usually a formula) – never written by the app.
export const SPECS = {
  "Raw Materials": {
    required: ["item name", "pack size"],
    key: "name",
    fields: {
      category: { h: ["category"], t: "str" },
      name: { h: ["item name", "item"], t: "str" },
      packSize: { h: ["pack size"], t: "num" },
      unit: { h: ["unit"], t: "str" },
      packPrice: { h: ["pack price", "price"], t: "num" },
      costPerUnit: { h: ["cost per unit"], t: "num", ro: true },
      notes: { h: ["supplier notes", "notes", "supplier"], t: "str" },
      priceStatus: { h: ["price status"], t: "str" },
    },
  },
  Recipes: {
    required: ["recipe name", "ingredient"],
    key: "lineId",
    addColumns: [{ field: "lineId", header: "Line ID" }],
    fields: {
      name: { h: ["recipe name"], t: "str" },
      type: { h: ["type"], t: "str" },
      ingredient: { h: ["ingredient"], t: "str" },
      qtyPerKg: { h: ["qty per 1 kg of recipe", "qty per 1 kg", "qty per kg"], t: "num" },
      unit: { h: ["unit"], t: "str", ro: true },
      costPerUnit: { h: ["cost per unit"], t: "num", ro: true },
      lineCost: { h: ["cost for this line"], t: "num", ro: true },
      lineId: { h: ["line id"], t: "str" },
    },
  },
  Inventory: {
    required: ["item", "opening stock"],
    key: "item",
    fields: {
      item: { h: ["item"], t: "str", ro: true },
      category: { h: ["category"], t: "str", ro: true },
      unit: { h: ["unit"], t: "str", ro: true },
      opening: { h: ["opening stock", "opening"], t: "num" },
      purchased: { h: ["purchased"], t: "num" },
      used: { h: ["used"], t: "num" },
      current: { h: ["current stock"], t: "num", ro: true },
      reorderLevel: { h: ["reorder level"], t: "num" },
      status: { h: ["status"], t: "str", ro: true },
      costPerUnit: { h: ["cost per unit"], t: "num", ro: true },
      stockValue: { h: ["stock value"], t: "num", ro: true },
    },
  },
  Orders: {
    required: ["id", "customer"],
    key: "id",
    addColumns: [
      { field: "updatedAt", header: "Updated at" },
      { field: "hoursOverride", header: "Hours override" },
    ],
    fields: {
      id: { h: ["id", "order id"], t: "str" },
      createdAt: { h: ["created", "created at"], t: "date" },
      customer: { h: ["customer"], t: "str" },
      phone: { h: ["phone"], t: "str" },
      size: { h: ["size"], t: "str" },
      layers: { h: ["layers"], t: "num" },
      sponge: { h: ["sponge"], t: "str" },
      filling: { h: ["filling"], t: "str" },
      frosting: { h: ["frosting"], t: "str" },
      topping1: { h: ["topping 1"], t: "str" },
      topping2: { h: ["topping 2"], t: "str" },
      topping3: { h: ["topping 3"], t: "str" },
      box: { h: ["box", "cake box"], t: "str" },
      board: { h: ["board", "cake board"], t: "str" },
      candles: { h: ["candles"], t: "bool" },
      carryBag: { h: ["carry bag"], t: "bool" },
      delivery: { h: ["delivery"], t: "str" },
      distanceKm: { h: ["distance km", "distance"], t: "num" },
      deliveryDate: { h: ["delivery date"], t: "date" },
      status: { h: ["status"], t: "str" },
      advancePaid: { h: ["advance paid"], t: "num" },
      notes: { h: ["notes"], t: "str" },
      finalQuote: { h: ["quote"], t: "num" },
      totalCost: { h: ["cost"], t: "num" },
      profit: { h: ["profit"], t: "num" },
      hoursUsed: { h: ["hours"], t: "num" },
      rawMaterialCost: { h: ["raw material"], t: "num" },
      labour: { h: ["labour"], t: "num" },
      deliveryCharge: { h: ["delivery charge"], t: "num" },
      updatedAt: { h: ["updated at"], t: "str" },
      hoursOverride: { h: ["hours override"], t: "str" },
    },
  },
  "Recipe Library": {
    required: ["name", "source"],
    key: "name",
    fields: {
      name: { h: ["name"], t: "str" },
      type: { h: ["type"], t: "str" },
      label: { h: ["label"], t: "str" },
      sourceName: { h: ["source"], t: "str" },
      sourceUrl: { h: ["source url"], t: "str" },
      sourceBatch: { h: ["source batch"], t: "str" },
      scaledHow: { h: ["how it was scaled"], t: "str" },
      method: { h: ["method"], t: "str" },
      ingredients: { h: ["ingredients per 1 kg", "ingredients"], t: "str" },
      notes: { h: ["notes"], t: "str" },
    },
  },
};

// Settings tab: label in column A, value in column B.
export const SETTINGS_LABELS = {
  fillingGPerKg: "filling per gap between layers",
  crumbCoatGPerKgPerLayer: "crumb coat",
  finalFrostingGPerKgPerLayer: "final frosting",
  wastage: "wastage allowance on ingredients",
  labourRate: "labour rate",
  baseHours: "base hours per cake",
  hoursPerKg: "extra hours per kg of cake",
  hoursPerLayer: "extra hours per layer",
  hoursPerTopping: "extra hours per topping used",
  "overheads.gas": "gas lpg",
  "overheads.electricity": "electricity",
  "overheads.water": "water cleaning",
  "overheads.rent": "rent kitchen share",
  "overheads.equipment": "equipment wear miscellaneous",
  deliveryRatePerKm: "delivery rate",
  tripMultiplier: "trip multiplier",
  minDelivery: "minimum delivery charge",
  markup: "profit markup on total cost",
  gst: "gst on final price",
  rounding: "round final quote up to nearest",
};

// ---------- helpers ----------
function tab(grid, name) {
  const t = grid.tabs?.[name];
  return t || { values: [], formulas: [] };
}
const cell = (t, r, c) => (t.values[r] && t.values[r][c] !== undefined ? t.values[r][c] : "");
const formulaAt = (t, r, c) => (t.formulas?.[r] && t.formulas[r][c]) || "";

/** Locate the header row and map each spec field to a 0-based column. */
export function locateTable(t, spec) {
  for (let r = 0; r < Math.min(t.values.length, 15); r++) {
    const row = (t.values[r] || []).map(norm);
    if (spec.required.every((h) => row.includes(h))) {
      const cols = {};
      for (const [f, def] of Object.entries(spec.fields)) {
        for (const alias of def.h) {
          const i = row.indexOf(alias);
          if (i >= 0 && !Object.values(cols).includes(i)) { cols[f] = i; break; }
        }
      }
      let lastCol = 0;
      row.forEach((h, i) => { if (h) lastCol = i; });
      return { headerRow: r, cols, lastCol, headers: t.values[r] };
    }
  }
  return null;
}

function convert(v, type) {
  if (type === "num") return num(v);
  if (type === "bool") return toBool(v);
  if (type === "date") return toDate(v);
  return str(v);
}

function rowsOf(t, loc, spec, isValid) {
  const out = [];
  for (let r = loc.headerRow + 1; r < t.values.length; r++) {
    const rec = { _row: r + 1 };
    let any = false;
    for (const [f, c] of Object.entries(loc.cols)) {
      const raw = cell(t, r, c);
      if (!blank(raw)) any = true;
      rec[f] = convert(raw, spec.fields[f].t);
    }
    if (any && isValid(rec)) out.push(rec);
  }
  return out;
}

export function compositeRecipeKey(r) {
  return `${r.name}||${r.ingredient}`;
}

function parseLibraryIngredients(text, recipeName) {
  return String(text || "")
    .split(/;\s*/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => {
      const m = s.match(/^(.+?)\s+([\d.]+)\s*([A-Za-z]+)?$/);
      if (!m) return { name: recipeName, ingredient: s, qtyPerKg: 0, unit: "" };
      return { name: recipeName, ingredient: m[1].trim(), qtyPerKg: Number(m[2]), unit: m[3] || "" };
    });
}

function slug(s) {
  return String(s).toLowerCase().replace(/\(adapted\)/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

function parseSettings(t) {
  const s = { overheads: {} };
  const labelRow = {};
  for (let r = 0; r < t.values.length; r++) {
    const a = norm(cell(t, r, 0));
    if (!a) continue;
    for (const [key, label] of Object.entries(SETTINGS_LABELS)) {
      if (a === label && labelRow[key] === undefined && !blank(cell(t, r, 1))) {
        labelRow[key] = r;
        const v = num(cell(t, r, 1));
        if (key.startsWith("overheads.")) s.overheads[key.split(".")[1]] = v;
        else s[key] = v;
      }
    }
  }
  // sub-tables, found by their header row
  const findHeader = (pred) => {
    for (let r = 0; r < t.values.length; r++) {
      const row = (t.values[r] || []).map(norm);
      if (pred(row)) return r;
    }
    return -1;
  };
  const tableRows = (hr) => {
    const out = [];
    if (hr < 0) return out;
    for (let r = hr + 1; r < t.values.length; r++) {
      if (blank(cell(t, r, 0))) break;
      out.push(r);
    }
    return out;
  };
  const sizesH = findHeader((row) => row[0] === "size name");
  s.sizes = tableRows(sizesH).map((r) => ({ name: str(cell(t, r, 0)), kg: num(cell(t, r, 1)), servings: fixServings(cell(t, r, 2)) }));
  const layersH = findHeader((row) => row[0] === "number of layers");
  s.layers = tableRows(layersH).map((r) => num(cell(t, r, 0))).filter((x) => x > 0);
  const menuH = findHeader((row) => row[0] === "type" && row[1] === "recipe name");
  s.recipeMenu = tableRows(menuH).map((r) => ({ type: str(cell(t, r, 0)), name: str(cell(t, r, 1)) })).filter((m) => m.name);
  const topH = findHeader((row) => row[0] === "topping name");
  s.toppings = tableRows(topH).map((r) => ({
    name: str(cell(t, r, 0)),
    material: str(cell(t, r, 1)),
    unit: str(cell(t, r, 2)),
    fixedQty: num(cell(t, r, 3)),
    qtyPerKg: num(cell(t, r, 4)),
  }));
  const boxH = findHeader((row) => row[0] === "cake box" && row[1] === "cost each");
  s.boxes = tableRows(boxH).map((r) => str(cell(t, r, 0)));
  const boardH = findHeader((row) => row[0] === "cake board" && row[1] === "cost each");
  s.boards = tableRows(boardH).map((r) => str(cell(t, r, 0)));
  const addH = findHeader((row) => row[0] === "add on");
  s.addons = tableRows(addH).map((r) => ({ name: str(cell(t, r, 0)), material: str(cell(t, r, 1)), qty: num(cell(t, r, 2)) }));
  return { settings: s, labelRow };
}

/** Whole workbook -> app state. Also returns `base` (sheet values as loaded). */
export function parseWorkbook(grid) {
  const state = {};
  const problems = [];

  const rm = tab(grid, "Raw Materials");
  const rmLoc = locateTable(rm, SPECS["Raw Materials"]);
  if (!rmLoc) problems.push("Raw Materials header row not found");
  state.rawMaterials = rmLoc
    ? rowsOf(rm, rmLoc, SPECS["Raw Materials"], (r) => r.name).map((r) => ({
        category: r.category, name: r.name, packSize: r.packSize, unit: r.unit, packPrice: r.packPrice,
        notes: r.notes || "", priceStatus: r.priceStatus || "",
        estimated: /ESTIMATED/i.test(r.priceStatus || "") || /^ESTIMATED/i.test(r.notes || ""),
      }))
    : [];

  const rc = tab(grid, "Recipes");
  const rcLoc = locateTable(rc, SPECS.Recipes);
  if (!rcLoc) problems.push("Recipes header row not found");
  const unitOf = new Map(state.rawMaterials.map((m) => [m.name, m.unit]));
  state.recipes = rcLoc
    ? rowsOf(rc, rcLoc, SPECS.Recipes, (r) => r.name).map((r) => ({
        lineId: r.lineId || "",
        name: r.name, type: r.type, ingredient: r.ingredient, qtyPerKg: r.qtyPerKg,
        unit: unitOf.get(r.ingredient) || r.unit || "",
      }))
    : [];

  const inv = tab(grid, "Inventory");
  const invLoc = locateTable(inv, SPECS.Inventory);
  if (!invLoc) problems.push("Inventory header row not found");
  state.inventory = invLoc
    ? rowsOf(inv, invLoc, SPECS.Inventory, (r) => r.item).map((r) => ({
        item: r.item, opening: r.opening, purchased: r.purchased, used: r.used, reorderLevel: r.reorderLevel,
      }))
    : [];

  const st = tab(grid, "Settings");
  state.settings = parseSettings(st).settings;

  const od = tab(grid, "Orders");
  const odLoc = locateTable(od, SPECS.Orders);
  if (!odLoc) problems.push("Orders header row not found");
  state.orders = odLoc
    ? rowsOf(od, odLoc, SPECS.Orders, (r) => r.id && (r.customer || r.createdAt || r.size)).map(orderFromRow)
    : [];

  const lb = tab(grid, "Recipe Library");
  const lbLoc = locateTable(lb, SPECS["Recipe Library"]);
  state.library = lbLoc
    ? rowsOf(lb, lbLoc, SPECS["Recipe Library"], (r) => r.name).map((r) => {
        const recipeName = r.name.replace(/\s*\(adapted\)\s*$/i, "").trim();
        return {
          id: slug(r.name), name: r.name, type: r.type, label: r.label, sourceName: r.sourceName,
          sourceUrl: r.sourceUrl, sourceBatch: r.sourceBatch, scaledHow: r.scaledHow, notes: r.notes,
          method: String(r.method || "").split(/\s*\|\s*/).map((x) => x.replace(/^\d+\.\s*/, "")).filter(Boolean),
          ingredients: parseLibraryIngredients(r.ingredients, recipeName),
        };
      })
    : [];

  return { state, problems };
}

export function orderFromRow(r) {
  const o = {
    id: r.id, createdAt: r.createdAt, customer: r.customer, phone: r.phone, size: r.size,
    layers: r.layers || 2, sponge: r.sponge, filling: r.filling, frosting: r.frosting,
    topping1: r.topping1 || "None", topping2: r.topping2 || "None", topping3: r.topping3 || "None",
    box: r.box, board: r.board, candles: !!r.candles, carryBag: !!r.carryBag,
    delivery: r.delivery || "Pickup", distanceKm: r.distanceKm || 0, deliveryDate: r.deliveryDate,
    status: r.status || "Enquiry", advancePaid: r.advancePaid || 0, notes: r.notes,
    updatedAt: r.updatedAt || "",
    hoursOverride: blank(r.hoursOverride) ? null : Number(r.hoursOverride),
    quoteSnapshot: null,
  };
  if (r.finalQuote) {
    o.quoteSnapshot = {
      finalQuote: r.finalQuote, totalCost: r.totalCost, profit: r.profit, hoursUsed: r.hoursUsed,
      rawMaterialCost: r.rawMaterialCost, labour: r.labour, delivery: r.deliveryCharge, fromSheet: true,
    };
  }
  return o;
}

/** App order -> sheet field values (what gets written). */
export function orderToFields(o) {
  const q = o.quoteSnapshot || {};
  const r2 = (x) => (isFinite(Number(x)) ? Math.round(Number(x) * 100) / 100 : "");
  return {
    id: o.id, createdAt: o.createdAt, customer: o.customer || "", phone: o.phone || "", size: o.size || "",
    layers: Number(o.layers) || "", sponge: o.sponge || "", filling: o.filling || "", frosting: o.frosting || "",
    topping1: o.topping1 || "None", topping2: o.topping2 || "None", topping3: o.topping3 || "None",
    box: o.box || "", board: o.board || "", candles: o.candles ? "Yes" : "No", carryBag: o.carryBag ? "Yes" : "No",
    delivery: o.delivery || "", distanceKm: Number(o.distanceKm) || 0, deliveryDate: o.deliveryDate || "",
    status: o.status || "Enquiry", advancePaid: Number(o.advancePaid) || 0, notes: o.notes || "",
    finalQuote: r2(q.finalQuote), totalCost: r2(q.totalCost), profit: r2(q.profit), hoursUsed: r2(q.hoursUsed),
    rawMaterialCost: r2(q.rawMaterialCost), labour: r2(q.labour), deliveryCharge: r2(q.delivery),
    updatedAt: o.updatedAt || new Date().toISOString(),
    hoursOverride: o.hoursOverride == null || o.hoursOverride === "" ? "" : Number(o.hoursOverride),
  };
}

// ---------- formulas for new rows ----------
/** Shift relative row numbers in A1 references by `delta` (skips $-anchored rows and "strings"). */
export function shiftFormula(f, delta) {
  if (!f || !delta) return f;
  return f
    .split('"')
    .map((part, i) =>
      i % 2 === 1
        ? part
        : part.replace(/(?<![A-Za-z_\d])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g, (m, d1, col, d2, row) =>
            d2 ? m : `${d1}${col}${Number(row) + delta}`
          )
    )
    .join('"');
}

function lastUsedRow(t, fromRow) {
  // 0-based index of the last row with any value or formula, at/after fromRow
  let last = fromRow;
  const n = Math.max(t.values.length, t.formulas?.length || 0);
  for (let r = fromRow; r < n; r++) {
    const vr = t.values[r] || [];
    const fr = t.formulas?.[r] || [];
    if (vr.some((v) => !blank(v)) || fr.some((v) => !blank(v))) last = r;
  }
  return last;
}

export function newId(prefix) {
  return prefix + "-" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase();
}

// ---------- planning writes ----------
/**
 * Turn queued app operations into cell writes against a FRESH grid.
 * ops:
 *  {kind:"set", tab, key, field, value, base}        Raw Materials / Inventory / Recipes / Settings
 *  {kind:"order", id, fields, updatedAt, isNew}      Orders upsert (last write wins by Updated at)
 *  {kind:"append", tab, record}                       Raw Materials (also adds the Inventory row) / Recipes
 * Returns { cells, conflicts, done } where cells are
 *  {tab,row,col,value} or {tab,row,col,formula} (1-based), optional ifEmpty / expectKey {col,value}.
 */
export function planWrites(grid, ops) {
  const cells = [];
  const conflicts = [];
  const done = [];
  const nextFree = {};
  const locs = {};
  const loc = (name) => {
    if (!(name in locs)) locs[name] = locateTable(tab(grid, name), SPECS[name]) || null;
    return locs[name];
  };

  // 1. bootstrap: add missing columns at the END of the header row + Line IDs
  for (const name of ["Orders", "Recipes"]) {
    const l = loc(name);
    if (!l) continue;
    const t = tab(grid, name);
    let end = l.lastCol;
    for (const add of SPECS[name].addColumns) {
      if (l.cols[add.field] === undefined) {
        end += 1;
        l.cols[add.field] = end;
        l.lastCol = end;
        cells.push({ tab: name, row: l.headerRow + 1, col: end + 1, value: add.header, ifEmpty: true, boot: true });
      }
    }
    if (name === "Recipes") {
      const nameCol = l.cols.name;
      const idCol = l.cols.lineId;
      for (let r = l.headerRow + 1; r < t.values.length; r++) {
        if (!blank(cell(t, r, nameCol)) && blank(cell(t, r, idCol))) {
          const id = "R" + String(r + 1).padStart(4, "0") + "-" + Math.random().toString(36).slice(2, 6).toUpperCase();
          setCell(t, r, idCol, id);
          cells.push({ tab: name, row: r + 1, col: idCol + 1, value: id, ifEmpty: true, boot: true });
        }
      }
    }
  }

  const findRow = (name, field, value) => {
    const l = loc(name);
    if (!l || l.cols[field] === undefined) return -1;
    const t = tab(grid, name);
    for (let r = l.headerRow + 1; r < t.values.length; r++) {
      if (str(cell(t, r, l.cols[field])) === String(value)) return r;
    }
    return -1;
  };
  const findRecipeRow = (key) => {
    let r = findRow("Recipes", "lineId", key);
    if (r >= 0) return r;
    const [n, ing] = String(key).split("||");
    const l = loc("Recipes");
    const t = tab(grid, "Recipes");
    if (!l) return -1;
    for (let i = l.headerRow + 1; i < t.values.length; i++) {
      if (str(cell(t, i, l.cols.name)) === n && str(cell(t, i, l.cols.ingredient)) === ing) return i;
    }
    return -1;
  };
  const freeRow = (name) => {
    if (nextFree[name] === undefined) {
      const l = loc(name);
      nextFree[name] = lastUsedRow(tab(grid, name), l ? l.headerRow : 0) + 1; // 0-based
    }
    return nextFree[name]++;
  };
  const same = (a, b) => String(a ?? "") === String(b ?? "") || (num(a) === num(b) && !blank(a) && !blank(b) && isFinite(Number(a)) && isFinite(Number(b)));

  const pushValue = (name, r, c, value, extra = {}) => {
    const t = tab(grid, name);
    if (formulaAt(t, r, c)) return false; // never overwrite a formula
    cells.push({ tab: name, row: r + 1, col: c + 1, value, ...extra });
    setCell(t, r, c, value);
    return true;
  };

  const appendWithTemplate = (name, r, record, opRef, extraValues = {}) => {
    // r: 0-based target row; copies formulas from the last data row, shifted.
    // The key cell goes first; every other cell of the new row is only written if
    // that key is still ours (protects against two phones appending at once).
    const t = tab(grid, name);
    const l = loc(name);
    const spec = SPECS[name];
    const tpl = lastUsedRow(t, l.headerRow + 1);
    const delta = r - tpl;
    const vals = { ...record, ...extraValues };
    const keyCol = l.cols[spec.key];
    const keyIsValue = keyCol !== undefined && !spec.fields[spec.key]?.ro && !formulaAt(t, tpl, keyCol) && !blank(vals[spec.key]);
    let guard = {};
    if (keyIsValue) {
      cells.push({ tab: name, row: r + 1, col: keyCol + 1, value: vals[spec.key], ifEmpty: true, keyCell: true, opRef });
      guard = { expectKey: { col: keyCol + 1, value: String(vals[spec.key]) } };
    }
    const width = Math.max((t.formulas?.[tpl] || []).length, l.lastCol + 1);
    for (let c = 0; c < width; c++) {
      const f = formulaAt(t, tpl, c);
      if (f && tpl > l.headerRow) {
        const nf = shiftFormula(f, delta);
        cells.push({ tab: name, row: r + 1, col: c + 1, formula: nf, ifEmpty: true, ...guard });
        setFormulaCell(t, r, c, nf);
      }
    }
    if (keyIsValue) setCell(t, r, keyCol, vals[spec.key]);
    for (const [field, v] of Object.entries(vals)) {
      const c = l.cols[field];
      if (c === undefined || spec.fields[field]?.ro || (keyIsValue && field === spec.key)) continue;
      if (formulaAt(t, tpl, c)) continue; // column is formula-driven
      if (blank(v)) continue;
      cells.push({ tab: name, row: r + 1, col: c + 1, value: v, ifEmpty: true, ...guard });
      setCell(t, r, c, v);
    }
  };

  for (const op of ops) {
    try {
      if (op.kind === "set" && op.tab === "Settings") {
        const t = tab(grid, "Settings");
        const { labelRow } = parseSettings(t);
        const r = labelRow[op.key];
        if (r === undefined) { conflicts.push({ op, reason: "setting label not found in sheet" }); continue; }
        const cur = cell(t, r, 1);
        if (op.base !== undefined && !same(cur, op.base) && !same(cur, op.value)) {
          conflicts.push({ op, reason: "changed in the sheet since the app loaded it; sheet kept" });
          continue;
        }
        if (!pushValue("Settings", r, 1, op.value, { expectKey: { col: 1, value: str(cell(t, r, 0)) } }))
          conflicts.push({ op, reason: "cell holds a formula; not overwritten" });
        else done.push(op);
      } else if (op.kind === "set") {
        const l = loc(op.tab);
        const spec = SPECS[op.tab];
        if (!l) { conflicts.push({ op, reason: "tab layout not found" }); continue; }
        const c = l.cols[op.field];
        if (c === undefined || spec.fields[op.field]?.ro) { conflicts.push({ op, reason: "column not writable" }); continue; }
        const r = op.tab === "Recipes" ? findRecipeRow(op.key) : findRow(op.tab, spec.key, op.key);
        if (r < 0) { conflicts.push({ op, reason: "row not found in sheet (renamed or removed?)" }); continue; }
        const t = tab(grid, op.tab);
        const cur = cell(t, r, c);
        if (op.base !== undefined && !same(cur, op.base) && !same(cur, op.value)) {
          conflicts.push({ op, reason: "changed in the sheet since the app loaded it; sheet kept" });
          continue;
        }
        const keyCol = op.tab === "Recipes" ? l.cols.lineId : l.cols[spec.key];
        if (!pushValue(op.tab, r, c, op.value, { expectKey: { col: keyCol + 1, value: str(cell(t, r, keyCol)) } }))
          conflicts.push({ op, reason: "cell holds a formula; not overwritten" });
        else done.push(op);
      } else if (op.kind === "order") {
        const l = loc("Orders");
        if (!l) { conflicts.push({ op, reason: "Orders layout not found" }); continue; }
        const t = tab(grid, "Orders");
        let r = findRow("Orders", "id", op.id);
        if (r >= 0) {
          const sheetUpdated = str(cell(t, r, l.cols.updatedAt));
          if (sheetUpdated && op.updatedAt && sheetUpdated > op.updatedAt) {
            conflicts.push({ op, reason: "order edited more recently in the sheet; sheet kept" });
            continue;
          }
          for (const [f, v] of Object.entries(op.fields)) {
            const c = l.cols[f];
            if (c === undefined || f === "id") continue;
            if (same(cell(t, r, c), v) && f !== "updatedAt") continue;
            pushValue("Orders", r, c, v, { expectKey: { col: l.cols.id + 1, value: op.id } });
          }
          done.push(op);
        } else {
          r = freeRow("Orders");
          appendWithTemplate("Orders", r, op.fields, op.queuedAt);
          done.push(op);
        }
      } else if (op.kind === "append" && op.tab === "Raw Materials") {
        if (findRow("Raw Materials", "name", op.record.name) >= 0) { conflicts.push({ op, reason: "item already exists" }); continue; }
        const r = freeRow("Raw Materials");
        appendWithTemplate("Raw Materials", r, op.record, op.queuedAt);
        // Inventory mirrors Raw Materials row-for-row via formulas: add the same row number.
        const it = tab(grid, "Inventory");
        const il = loc("Inventory");
        if (il) {
          const rowEmpty = !(it.values[r] || []).some((v) => !blank(v)) && !(it.formulas?.[r] || []).some((v) => !blank(v));
          if (rowEmpty) {
            appendWithTemplate("Inventory", r, { opening: 0, purchased: 0, used: 0, reorderLevel: 0, ...(op.inventory || {}) }, op.queuedAt);
            nextFree.Inventory = Math.max(nextFree.Inventory || 0, r + 1);
          } else conflicts.push({ op, reason: `Inventory row ${r + 1} already used; add the stock row by hand` });
        }
        done.push(op);
      } else if (op.kind === "append" && op.tab === "Recipes") {
        const r = freeRow("Recipes");
        appendWithTemplate("Recipes", r, op.record, op.queuedAt);
        done.push(op);
      } else {
        conflicts.push({ op, reason: "unknown operation" });
      }
    } catch (e) {
      conflicts.push({ op, reason: String(e && e.message) });
    }
  }
  return { cells, conflicts, done };
}

function setFormulaCell(t, r, c, f) {
  if (!t.formulas) t.formulas = [];
  while (t.formulas.length <= r) t.formulas.push([]);
  const row = t.formulas[r];
  while (row.length <= c) row.push("");
  row[c] = f;
  setCell(t, r, c, "");
}

function setCell(t, r, c, v) {
  while (t.values.length <= r) t.values.push([]);
  const row = t.values[r];
  while (row.length <= c) row.push("");
  row[c] = v;
}

export function colLetter(c) {
  // 1-based -> A1 letters
  let s = "";
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
