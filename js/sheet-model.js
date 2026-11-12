// Sheet model: maps the Bakery Workbook tabs (by header NAME, not position)
// to the app's data model, and turns app edits into cell-level writes.
// Pure functions, no DOM, shared by both sync transports (Apps Script and
// direct Google Sheets API) and by the Node tests.
//
// Grid shape (what a transport's read() returns):
//   { tabs: { [tabName]: { values: any[][], formulas: string[][] } },
//     meta: { [tabName]: { sheetId, rowCount, columnCount } } }
// Rows/cols in `values` are 0-based; cell writes use 1-based row/col.

export const CORE_TABS = ["Raw Materials", "Recipes", "Settings", "Inventory", "Orders", "Recipe Library"];

// Tabs the app creates itself (via the Sheets API addSheet) the first time she
// saves something that has no home in the original workbook. Header row 1.
// Columns are found by header name, so she can reorder them in the sheet.
export const APP_TABS = {
  // What she sells. "Calculator" = priced by the cake calculator (sizes, layers, sponge…);
  // "Per unit" = quantity × Unit price (granola bars, cookies…).
  Categories: {
    key: ["name"],
    cols: [
      ["name", "Category", "str"], ["pricing", "Pricing", "str"], ["unitPrice", "Unit price (₹)", "num"],
      ["unit", "Unit", "str"], ["costPerUnit", "Cost per unit (₹)", "num"], ["active", "Active", "str"],
      ["notes", "Notes", "str"],
    ],
  },
  Customers: {
    key: ["id"],
    cols: [
      ["id", "Customer ID", "str"], ["name", "Name", "str"], ["phone", "Phone", "str"],
      ["notes", "Notes", "str"], ["createdAt", "Created at", "str"],
    ],
  },
  "Price List": {
    key: ["kind", "item"],
    cols: [["kind", "Kind", "str"], ["item", "Item", "str"], ["price", "Selling price (₹)", "num"], ["updatedAt", "Updated at", "str"]],
  },
  "Pricing Rules": {
    key: ["id"],
    cols: [
      ["id", "Rule ID", "str"], ["name", "Rule", "str"], ["effect", "Effect", "str"], ["type", "Type", "str"],
      ["amount", "Amount", "num"], ["active", "Active", "bool"], ["applyTo", "Apply to", "str"], ["notes", "Notes", "str"],
    ],
  },
  "Delivery Zones": {
    key: ["id"],
    cols: [["id", "Zone ID", "str"], ["name", "Zone", "str"], ["charge", "Flat charge (₹)", "num"], ["notes", "Notes", "str"]],
  },
  "Recipe Notes": {
    key: ["recipe"],
    cols: [["recipe", "Recipe", "str"], ["type", "Type", "str"], ["method", "Method / notes", "str"], ["updatedAt", "Updated at", "str"]],
  },
  Purchases: {
    key: ["id"],
    cols: [
      ["date", "Date", "str"], ["item", "Item", "str"], ["packs", "Packs", "num"], ["qty", "Qty added", "num"],
      ["unit", "Unit", "str"], ["pricePerPack", "Price per pack (₹)", "num"], ["total", "Total (₹)", "num"],
      ["notes", "Notes", "str"], ["id", "Purchase ID", "str"],
    ],
  },
};
// Seeded into a new Categories tab on the first sync (and used offline / in local mode).
export const DEFAULT_CATEGORIES = [
  { name: "Cake", pricing: "Calculator", unitPrice: 0, unit: "cake", costPerUnit: 0, active: true, notes: "Priced by the cake calculator (size, layers, sponge, toppings…)" },
  { name: "Granola bar", pricing: "Per unit", unitPrice: 90, unit: "bar", costPerUnit: 0, active: true, notes: "" },
  { name: "Cookies", pricing: "Per unit", unitPrice: 90, unit: "cookie", costPerUnit: 0, active: true, notes: "" },
];
export function categoryToFields(c) {
  return {
    pricing: c.pricing === "Calculator" ? "Calculator" : "Per unit",
    unitPrice: c.pricing === "Calculator" ? "" : Number(c.unitPrice) || 0,
    unit: c.unit || "",
    costPerUnit: Number(c.costPerUnit) || 0,
    active: c.active === false ? "No" : "Yes",
    notes: c.notes || "",
  };
}
export function normalizeCategory(r) {
  return {
    name: String(r.name || "").trim(),
    pricing: /calc|cake/i.test(r.pricing || "") ? "Calculator" : "Per unit",
    unitPrice: Number(r.unitPrice) || 0,
    unit: r.unit || "",
    costPerUnit: Number(r.costPerUnit) || 0,
    active: !/^(no|n|false|0|off|inactive)$/i.test(String(r.active ?? "").trim()),
    notes: r.notes || "",
  };
}

// Read for layout only (the app never writes to it): its formulas and data
// validations tell us how many rows of each Settings list the Calculator sees.
export const READ_ONLY_TABS = ["Calculator"];
export const TABS = [...CORE_TABS, ...READ_ONLY_TABS, ...Object.keys(APP_TABS)];

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
      { field: "priceOverride", header: "Price override (₹)" },
      { field: "rulesApplied", header: "Rules applied" },
      { field: "deliveryZone", header: "Delivery zone" },
      { field: "extraAddons", header: "Other add-ons" },
      { field: "paymentStatus", header: "Payment status" },
      { field: "balance", header: "Balance (₹)" },
      { field: "category", header: "Category" },
      { field: "quantity", header: "Quantity" },
      { field: "unitPrice", header: "Unit price (₹)" },
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
      priceOverride: { h: ["price override"], t: "str" },
      rulesApplied: { h: ["rules applied"], t: "str" },
      deliveryZone: { h: ["delivery zone"], t: "str" },
      extraAddons: { h: ["other add ons"], t: "str" },
      paymentStatus: { h: ["payment status"], t: "str" },
      balance: { h: ["balance"], t: "num" },
      category: { h: ["category", "product"], t: "str" },
      quantity: { h: ["quantity", "qty"], t: "str" },
      unitPrice: { h: ["unit price", "price each"], t: "str" },
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

for (const [name, def] of Object.entries(APP_TABS)) {
  const fields = {};
  for (const [f, header, t] of def.cols) fields[f] = { h: [norm(header)], t };
  SPECS[name] = { required: def.key.map((k) => fields[k].h[0]), key: def.key[0], keys: def.key, fields, appTab: true };
}
export function appTabHeaders(name) {
  return APP_TABS[name].cols.map((c) => c[1]);
}

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

// ---------- Settings sub-tables (menus) ----------
// Each list sits under its own header row. Rows inside a list may be blank
// (reserved rows, or an item she removed): parsing skips them and stops at the
// next section title ("6. Layers"), the next list header or a setting label.
export const SECTIONS = {
  sizes: { label: "Cake sizes", header: (row) => row[0] === "size name", keyCol: 0, cols: { name: 0, kg: 1, servings: 2 } },
  layers: { label: "Layers", header: (row) => row[0] === "number of layers", keyCol: 0, cols: { value: 0 } },
  recipeMenu: { label: "Recipe menu", header: (row) => row[0] === "type" && row[1] === "recipe name", keyCol: 1, cols: { type: 0, name: 1 } },
  toppings: { label: "Toppings", header: (row) => row[0] === "topping name", keyCol: 0, cols: { name: 0, material: 1, unit: 2, fixedQty: 3, qtyPerKg: 4 } },
  boxes: { label: "Cake boxes", header: (row) => row[0] === "cake box" && row[1] === "cost each", keyCol: 0, cols: { name: 0 } },
  boards: { label: "Cake boards", header: (row) => row[0] === "cake board" && row[1] === "cost each", keyCol: 0, cols: { name: 0 } },
  addons: { label: "Add-ons", header: (row) => row[0] === "add on", keyCol: 0, cols: { name: 0, material: 1, qty: 2 } },
};
const SETTING_LABEL_SET = new Set(Object.values(SETTINGS_LABELS));

function findHeaderRow(t, pred) {
  for (let r = 0; r < t.values.length; r++) {
    const row = (t.values[r] || []).map(norm);
    if (pred(row)) return r;
  }
  return -1;
}
function isBoundary(t, r) {
  const raw = cell(t, r, 0);
  if (typeof raw === "string" && /^\s*\d+\.\s+\D/.test(raw)) return true; // "6. Layers"
  const row = (t.values[r] || []).map(norm);
  if (Object.values(SECTIONS).some((sec) => sec.header(row))) return true;
  return SETTING_LABEL_SET.has(row[0]);
}

/** A1 ranges on the Settings tab referenced by the Calculator (formulas + data validations). */
export function settingsRefs(grid) {
  const texts = [];
  const calc = grid.tabs?.Calculator;
  if (calc) for (const row of calc.formulas || []) for (const f of row || []) if (f) texts.push(f);
  for (const v of grid.validationRefs || []) texts.push(String(v));
  const out = [];
  const re = /(?:'Settings'|\bSettings)!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g;
  const ci = (L) => { let n = 0; for (const ch of L) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  for (const txt of texts) {
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(txt))) {
      const c1 = ci(m[1]), r1 = Number(m[2]) - 1;
      const c2 = m[3] ? ci(m[3]) : c1, r2 = m[4] ? Number(m[4]) - 1 : r1;
      if (r2 > r1) out.push({ c1: Math.min(c1, c2), c2: Math.max(c1, c2), r1, r2, a1: m[0] });
    }
  }
  return out;
}

/**
 * Where a Settings list lives and which rows the app may write a new item into.
 * Rows are 0-based. `allowed` = [first,last] rows a new item may use: inside every
 * Calculator range that covers this list (or, with no such range, up to one blank
 * row above the next section title). `free` = blank rows inside `allowed`.
 */
export function sectionLayout(t, key, refs = [], opts = {}) {
  const sec = SECTIONS[key];
  const hr = findHeaderRow(t, sec.header);
  if (hr < 0) return null;
  let boundary = null;
  for (let r = hr + 1; r < t.values.length; r++) if (isBoundary(t, r)) { boundary = r; break; }
  const stop = boundary ?? t.values.length;
  const width = Math.max(lastNonBlank(t.values[hr] || []) + 1, ...Object.values(sec.cols).map((c) => c + 1));
  const rows = [];
  for (let r = hr + 1; r < stop; r++) if (!blank(cell(t, r, sec.keyCol))) rows.push(r);
  const lastData = rows.length ? rows[rows.length - 1] : hr;
  // Calculator ranges that cover this list's key column and start at its first rows
  const firstData = rows.length ? rows[0] : hr + 1;
  let related = refs.filter((x) => x.c1 <= sec.keyCol && x.c2 >= sec.keyCol && x.r1 <= Math.max(lastData, firstData) && x.r2 >= firstData && x.r1 >= hr && x.r1 < stop);
  if (opts.matchRows) related = related.filter((x) => opts.matchRows.some((r) => r >= x.r1 && r <= x.r2));
  let lo = hr + 1;
  let hi = boundary != null ? boundary - 2 : Infinity;
  if (related.length) {
    lo = Math.max(lo, ...related.map((x) => x.r1));
    hi = Math.min(...related.map((x) => x.r2));
    if (boundary != null) hi = Math.min(hi, boundary - 1);
  }
  const isFreeRow = (r) => {
    if (!blank(cell(t, r, sec.keyCol))) return false;
    for (const c of Object.values(sec.cols)) if (!formulaAt(t, r, c) && !blank(cell(t, r, c))) return false;
    return true;
  };
  const free = [];
  const scanHi = hi === Infinity ? Math.max(lastData + 1, t.values.length) : hi;
  for (let r = lo; r <= scanHi; r++) {
    if (r >= stop && hi !== Infinity) break;
    if (isFreeRow(r)) free.push(r);
  }
  if (hi === Infinity && !free.some((r) => r > lastData)) free.push(lastData + 1);
  return {
    key, header: hr, rows, boundary, width, lo, hi, free, related,
    rangeText: related.length ? [...new Set(related.map((x) => x.a1.replace(/'/g, "")))].join(", ") : `Settings rows ${lo + 1}–${hi === Infinity ? "end" : hi + 1}`,
  };
}
function lastNonBlank(arr) { let l = -1; arr.forEach((v, i) => { if (!blank(v)) l = i; }); return l; }

function parseSettings(t, refs = []) {
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
  const capacity = {};
  const rowsOfSec = (key) => {
    const lay = sectionLayout(t, key, refs);
    if (!lay) return [];
    capacity[key] = { free: lay.free.length, range: lay.rangeText, found: true };
    return lay.rows;
  };
  s.sizes = rowsOfSec("sizes").map((r) => ({ name: str(cell(t, r, 0)), kg: num(cell(t, r, 1)), servings: fixServings(cell(t, r, 2)) }));
  s.layers = rowsOfSec("layers").map((r) => num(cell(t, r, 0))).filter((x) => x > 0);
  s.recipeMenu = rowsOfSec("recipeMenu").map((r) => ({ type: str(cell(t, r, 0)), name: str(cell(t, r, 1)) })).filter((m) => m.name);
  s.toppings = rowsOfSec("toppings").map((r) => ({
    name: str(cell(t, r, 0)),
    material: str(cell(t, r, 1)),
    unit: str(cell(t, r, 2)),
    fixedQty: num(cell(t, r, 3)),
    qtyPerKg: num(cell(t, r, 4)),
  }));
  s.boxes = rowsOfSec("boxes").map((r) => str(cell(t, r, 0)));
  s.boards = rowsOfSec("boards").map((r) => str(cell(t, r, 0)));
  s.addons = rowsOfSec("addons").map((r) => ({ name: str(cell(t, r, 0)), material: str(cell(t, r, 1)), qty: num(cell(t, r, 2)) }));
  // per recipe type (Sponge / Filling / Frosting): Calculator ranges may cover one type only
  for (const type of [...new Set(s.recipeMenu.map((m) => m.type))]) {
    const lay = sectionLayout(t, "recipeMenu", refs, { matchRows: menuRowsOfType(t, type) });
    if (lay) capacity["recipeMenu:" + type] = { free: lay.free.length, range: lay.rangeText, found: true };
  }
  return { settings: s, labelRow, capacity };
}
function menuRowsOfType(t, type) {
  const hr = findHeaderRow(t, SECTIONS.recipeMenu.header);
  const out = [];
  if (hr < 0) return out;
  for (let r = hr + 1; r < t.values.length; r++) {
    if (isBoundary(t, r)) break;
    if (str(cell(t, r, 0)) === type && !blank(cell(t, r, 1))) out.push(r);
  }
  return out;
}

function parseAppTab(grid, name) {
  const t = tab(grid, name);
  const spec = SPECS[name];
  const loc = locateTable(t, spec);
  if (!loc) return [];
  return rowsOf(t, loc, spec, (r) => spec.keys.every((k) => !blank(r[k]))).map(({ _row, ...rest }) => rest);
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
  const refs = settingsRefs(grid);
  const ps = parseSettings(st, refs);
  state.settings = ps.settings;
  state.capacity = ps.capacity;
  state.sheetTabs = Object.keys(grid.tabs || {});

  state.categoriesFromSheet = !!grid.tabs?.Categories;
  state.categories = parseAppTab(grid, "Categories").map(normalizeCategory).filter((c) => c.name);
  state.customers = parseAppTab(grid, "Customers");
  state.priceList = parseAppTab(grid, "Price List").filter((p) => p.item);
  state.rules = parseAppTab(grid, "Pricing Rules").map((r) => ({ ...r, effect: /disc/i.test(r.effect) ? "Discount" : "Charge", type: /₹|flat|rs|inr/i.test(r.type) ? "₹" : "%", applyTo: /tick|order form|choose/i.test(r.applyTo) ? "When ticked" : "Every order" }));
  state.zones = parseAppTab(grid, "Delivery Zones");
  state.recipeNotes = parseAppTab(grid, "Recipe Notes");
  state.purchases = parseAppTab(grid, "Purchases");

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

export function splitList(v) {
  return String(v ?? "").split(/\s*;\s*/).map((x) => x.trim()).filter(Boolean);
}
export function paymentStatusOf(paid, total) {
  const p = Number(paid) || 0, t = Number(total) || 0;
  if (t > 0 && p >= t - 0.5) return "Paid in full";
  if (p > 0) return "Advance paid";
  return "Unpaid";
}

export function orderFromRow(r) {
  const category = r.category || "Cake";
  const o = {
    category,
    quantity: blank(r.quantity) ? null : num(r.quantity),
    unitPrice: blank(r.unitPrice) ? null : num(r.unitPrice),
    id: r.id, createdAt: r.createdAt, customer: r.customer, phone: r.phone, size: r.size,
    layers: r.layers || 2, sponge: r.sponge, filling: r.filling, frosting: r.frosting,
    topping1: r.topping1 || "None", topping2: r.topping2 || "None", topping3: r.topping3 || "None",
    box: r.box, board: r.board, candles: !!r.candles, carryBag: !!r.carryBag,
    delivery: r.delivery || "Pickup", distanceKm: r.distanceKm || 0, deliveryDate: r.deliveryDate,
    status: r.status || "Enquiry", advancePaid: r.advancePaid || 0, notes: r.notes,
    updatedAt: r.updatedAt || "",
    hoursOverride: blank(r.hoursOverride) ? null : Number(r.hoursOverride),
    priceOverride: blank(r.priceOverride) || !isFinite(Number(r.priceOverride)) ? null : Number(r.priceOverride),
    rules: splitList(r.rulesApplied),
    zone: r.deliveryZone || "",
    extraAddons: splitList(r.extraAddons),
    quoteSnapshot: null,
  };
  if (!blank(r.quantity)) {
    // per-unit order (granola bars, cookies…): no cake fields
    Object.assign(o, { layers: "", topping1: "None", topping2: "None", topping3: "None", deliveryCharge: Number(r.deliveryCharge) || 0 });
  }
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
  if (o.quantity != null && o.quantity !== "") {
    // per-unit order: fill what applies, leave the cake-only columns blank
    const total = Number(q.finalQuote) || 0;
    return {
      id: o.id, createdAt: o.createdAt, customer: o.customer || "", phone: o.phone || "",
      size: "", layers: "", sponge: "", filling: "", frosting: "", topping1: "", topping2: "", topping3: "",
      box: "", board: "", candles: "", carryBag: "",
      delivery: o.delivery || "", distanceKm: "", deliveryDate: o.deliveryDate || "",
      status: o.status || "Enquiry", advancePaid: Number(o.advancePaid) || 0, notes: o.notes || "",
      finalQuote: r2(q.finalQuote), totalCost: r2(q.totalCost), profit: r2(q.profit), hoursUsed: "",
      rawMaterialCost: r2(q.rawMaterialCost), labour: "", deliveryCharge: r2(q.delivery),
      updatedAt: o.updatedAt || new Date().toISOString(),
      hoursOverride: "", priceOverride: "", rulesApplied: "", deliveryZone: "", extraAddons: "",
      paymentStatus: paymentStatusOf(o.advancePaid, total),
      balance: r2(Math.max(total - (Number(o.advancePaid) || 0), 0)),
      category: o.category || "", quantity: Number(o.quantity) || 0, unitPrice: r2(o.unitPrice),
    };
  }
  return {
    category: o.category || "Cake", quantity: "", unitPrice: "",
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
    priceOverride: o.priceOverride == null || o.priceOverride === "" ? "" : Number(o.priceOverride),
    rulesApplied: (o.rules || []).join("; "),
    deliveryZone: o.zone || "",
    extraAddons: (o.extraAddons || []).join("; "),
    paymentStatus: paymentStatusOf(o.advancePaid, q.finalQuote),
    balance: r2(Math.max((Number(q.finalQuote) || 0) - (Number(o.advancePaid) || 0), 0)),
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
export function planWrites(grid, ops, opts = {}) {
  const appTabsOk = opts.appTabs !== false;
  const cells = [];
  const conflicts = [];
  const warnings = [];
  const newTabs = [];
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


  // ----- app-owned tabs (Customers, Price List, ...): create on first use -----
  const ensureAppTab = (name) => {
    const headers = appTabHeaders(name);
    if (!grid.tabs) grid.tabs = {};
    let t = grid.tabs[name];
    if (!t) {
      t = grid.tabs[name] = { values: [], formulas: [] };
      newTabs.push({ title: name, headers });
    }
    const hasAny = t.values.some((row) => (row || []).some((v) => !blank(v)));
    if (!hasAny) {
      headers.forEach((h, i) => {
        cells.push({ tab: name, row: 1, col: i + 1, value: h, ifEmpty: true, boot: true });
        setCell(t, 0, i, h);
      });
    }
    delete locs[name];
    const l = loc(name);
    if (!l) return null;
    // a column she deleted by hand comes back at the end of the header row
    for (const [f, h] of APP_TABS[name].cols.map((c) => [c[0], c[1]])) {
      if (l.cols[f] === undefined) {
        l.lastCol += 1;
        l.cols[f] = l.lastCol;
        cells.push({ tab: name, row: l.headerRow + 1, col: l.lastCol + 1, value: h, ifEmpty: true, boot: true });
        setCell(t, l.headerRow, l.lastCol, h);
      }
    }
    return l;
  };
  const findAppRow = (name, key) => {
    const l = loc(name);
    const t = tab(grid, name);
    const keys = SPECS[name].keys;
    for (let r = l.headerRow + 1; r < t.values.length; r++) {
      if (keys.every((k) => str(cell(t, r, l.cols[k])) === String(key[k] ?? ""))) return r;
    }
    return -1;
  };
  const rowOp = (op) => {
    const name = op.tab;
    if (!appTabsOk) { conflicts.push({ op, reason: `the ${name} tab can only be written in “Sign in with Google” mode (Settings → Google Sheet sync)` }); return; }
    const l = ensureAppTab(name);
    if (!l) { conflicts.push({ op, reason: `${name} tab has no header row the app recognises` }); return; }
    const t = tab(grid, name);
    const spec = SPECS[name];
    const guardField = spec.keys[spec.keys.length - 1];
    const guardCol = l.cols[guardField];
    let r = findAppRow(name, op.key);
    if (op.remove) {
      if (r < 0) { done.push(op); return; }
      const guard = { expectKey: { col: guardCol + 1, value: str(cell(t, r, guardCol)) } };
      const order = Object.entries(l.cols).sort((a, b) => (a[0] === guardField) - (b[0] === guardField));
      for (const [, c] of order) {
        if (formulaAt(t, r, c) || blank(cell(t, r, c))) continue;
        pushValue(name, r, c, "", guard);
      }
      done.push(op);
      return;
    }
    const fields = { ...op.key, ...op.fields };
    if (r >= 0) {
      const guard = { expectKey: { col: guardCol + 1, value: str(cell(t, r, guardCol)) } };
      const order = Object.entries(fields).sort((a, b) => (spec.keys.includes(a[0]) ? 1 : 0) - (spec.keys.includes(b[0]) ? 1 : 0));
      for (const [f, v] of order) {
        const c = l.cols[f];
        if (c === undefined || same(cell(t, r, c), v)) continue;
        pushValue(name, r, c, v, guard);
      }
      done.push(op);
      return;
    }
    r = freeRow(name);
    cells.push({ tab: name, row: r + 1, col: guardCol + 1, value: fields[guardField], ifEmpty: true, keyCell: true, opRef: op.queuedAt });
    setCell(t, r, guardCol, fields[guardField]);
    const guard = { expectKey: { col: guardCol + 1, value: String(fields[guardField]) } };
    for (const [f, v] of Object.entries(fields)) {
      const c = l.cols[f];
      if (c === undefined || f === guardField || blank(v)) continue;
      cells.push({ tab: name, row: r + 1, col: c + 1, value: v, ifEmpty: true, ...guard });
      setCell(t, r, c, v);
    }
    done.push(op);
  };

  // ----- Settings lists (sizes, layers, menu, toppings, boxes, boards, add-ons) -----
  const menuOp = (op) => {
    const t = tab(grid, "Settings");
    const sec = SECTIONS[op.section];
    if (!sec) { conflicts.push({ op, reason: "unknown list" }); return; }
    const refs = settingsRefs(grid);
    const base = sectionLayout(t, op.section, refs);
    if (!base) { conflicts.push({ op, reason: `${sec.label} list not found in the Settings tab` }); return; }
    const kc = sec.keyCol;
    const keyStr = String(op.key);
    const find = () => base.rows.find((r) => str(cell(t, r, kc)) === keyStr) ?? -1;
    const r0 = find();
    if (op.action === "add") {
      if (r0 >= 0) { done.push(op); return; } // already there
      let lay = base;
      if (op.section === "recipeMenu" && op.fields.type) {
        const same = base.rows.filter((r) => str(cell(t, r, 0)) === op.fields.type);
        if (same.length) lay = sectionLayout(t, op.section, refs, { matchRows: same });
      }
      const r = lay.free[0];
      if (r === undefined) {
        conflicts.push({ op, full: true, reason: `${sec.label} list in the sheet is full (${lay.rangeText}). Clear an old row inside that block in the sheet, or ask to extend the Calculator's range.` });
        return;
      }
      const near = [...base.rows].filter((x) => x < r).pop() ?? base.rows.find((x) => x > r);
      const vals = { ...op.fields };
      vals[Object.keys(sec.cols).find((f) => sec.cols[f] === kc)] = op.key;
      cells.push({ tab: "Settings", row: r + 1, col: kc + 1, value: op.key, ifEmpty: true, keyCell: true, opRef: op.queuedAt });
      setCell(t, r, kc, op.key);
      const guard = { expectKey: { col: kc + 1, value: keyStr } };
      for (const [f, c] of Object.entries(sec.cols)) {
        if (c === kc) continue;
        const v = vals[f];
        if (blank(v) || formulaAt(t, r, c) || (near !== undefined && formulaAt(t, near, c))) continue;
        cells.push({ tab: "Settings", row: r + 1, col: c + 1, value: v, ifEmpty: true, ...guard });
        setCell(t, r, c, v);
      }
      if (near !== undefined) {
        for (let c = 0; c < Math.max(base.width, (t.formulas?.[near] || []).length); c++) {
          const f = formulaAt(t, near, c);
          if (!f || formulaAt(t, r, c)) continue;
          const nf = shiftFormula(f, r - near);
          cells.push({ tab: "Settings", row: r + 1, col: c + 1, formula: nf, ifEmpty: true, ...guard });
          setFormulaCell(t, r, c, nf);
        }
      }
      done.push(op);
      return;
    }
    if (r0 < 0) { conflicts.push({ op, reason: `${sec.label}: “${keyStr}” not found in the sheet (renamed or removed there?)` }); return; }
    const guard = { expectKey: { col: kc + 1, value: keyStr } };
    if (op.action === "remove") {
      for (let c = 0; c < base.width; c++) {
        if (c === kc || formulaAt(t, r0, c) || blank(cell(t, r0, c))) continue;
        pushValue("Settings", r0, c, "", guard);
      }
      pushValue("Settings", r0, kc, "", guard);
      done.push(op);
      return;
    }
    // update (rename = new key value, written last)
    let rename = null;
    for (const [f, c] of Object.entries(sec.cols)) {
      if (!(f in op.fields)) continue;
      const v = op.fields[f];
      if (same(cell(t, r0, c), v) || formulaAt(t, r0, c)) continue;
      if (c === kc) { rename = v; continue; }
      pushValue("Settings", r0, c, v, guard);
    }
    if (rename !== null && !blank(rename)) pushValue("Settings", r0, kc, rename, guard);
    done.push(op);
  };

  const clearRecipeLine = (op) => {
    const l = loc("Recipes");
    if (!l) { conflicts.push({ op, reason: "Recipes layout not found" }); return; }
    const r = findRecipeRow(op.key);
    if (r < 0) { done.push(op); return; }
    const t = tab(grid, "Recipes");
    const idCol = l.cols.lineId;
    const guard = { expectKey: { col: idCol + 1, value: str(cell(t, r, idCol)) } };
    for (const f of ["name", "type", "ingredient", "qtyPerKg"]) {
      const c = l.cols[f];
      if (c === undefined || formulaAt(t, r, c) || blank(cell(t, r, c))) continue;
      pushValue("Recipes", r, c, "", guard);
    }
    if (!blank(cell(t, r, idCol))) pushValue("Recipes", r, idCol, "", guard);
    done.push(op);
  };

  // rows the sheet's own formulas look at, e.g. 'Raw Materials'!$B$5:$B$80
  const formulaReach = (tabName) => {
    let max = 0;
    const re = new RegExp("(?:'" + tabName + "'|\\b" + tabName.replace(/ /g, "_") + ")!\\$?[A-Z]{1,3}\\$?\\d+:\\$?[A-Z]{1,3}\\$?(\\d+)", "g");
    for (const tn of ["Settings", "Calculator", "Inventory"]) {
      for (const row of grid.tabs?.[tn]?.formulas || []) for (const f of row || []) {
        if (!f) continue;
        let m;
        re.lastIndex = 0;
        while ((m = re.exec(f))) max = Math.max(max, Number(m[1]));
      }
    }
    return max;
  };
  const warnReach = (op, tabName, r) => {
    const reach = formulaReach(tabName);
    if (reach && r + 1 > reach) warnings.push({ op, reason: `saved at ${tabName} row ${r + 1}, below row ${reach} that the sheet's own formulas read; the app prices it correctly, but the sheet's Calculator won't see it until you widen those ranges` });
  };

  for (const op of ops) {
    try {
      if (op.kind === "row") { rowOp(op); continue; }
      if (op.kind === "menu") { menuOp(op); continue; }
      if (op.kind === "clearRecipeLine") { clearRecipeLine(op); continue; }
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
        warnReach(op, "Raw Materials", r);
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
        warnReach(op, "Recipes", r);
        done.push(op);
      } else {
        conflicts.push({ op, reason: "unknown operation" });
      }
    } catch (e) {
      conflicts.push({ op, reason: String(e && e.message) });
    }
  }
  return { cells, conflicts, warnings, done, newTabs };
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
