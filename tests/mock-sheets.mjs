// Local mock of BOTH back ends over an in-memory copy of the real Bakery Workbook
// (tests/fixtures/sheet-snapshot.json, pulled read-only from the live sheet):
//   /exec                      -> runs apps-script/Code.gs itself (doGet/doPost) in a VM
//                                 with fake SpreadsheetApp / ContentService / ... services
//   /v4/spreadsheets/...       -> emulates the Google Sheets REST API (batchGet, values:batchUpdate,
//                                 :batchUpdate appendDimension, metadata)
// Nothing here talks to Google.
import http from "http";
import vm from "vm";
import { readFileSync } from "fs";

const here = (p) => new URL(p, import.meta.url);

export function loadWorkbook() {
  const snap = JSON.parse(readFileSync(here("./fixtures/sheet-snapshot.json")));
  const wb = {};
  for (const [name, t] of Object.entries(snap.tabs)) {
    const m = snap.meta[name];
    wb[name] = {
      name,
      sheetId: m.sheetId,
      maxRows: m.rowCount,
      maxCols: m.columnCount,
      values: t.values.map((r) => r.slice()),
      formulas: t.formulas.map((r) => r.slice()),
      formats: {},
    };
  }
  return { id: snap.spreadsheetId, tabs: wb };
}

const blank = (v) => v === "" || v === null || v === undefined;
function get(t, r, c, which = "values") {
  return (t[which][r] && t[which][r][c] !== undefined ? t[which][r][c] : "");
}
function put(t, r, c, which, v) {
  while (t[which].length <= r) t[which].push([]);
  const row = t[which][r];
  while (row.length <= c) row.push("");
  row[c] = v;
}
function lastRow(t) {
  let last = 0;
  const n = Math.max(t.values.length, t.formulas.length);
  for (let r = 0; r < n; r++) {
    if ((t.values[r] || []).some((v) => !blank(v)) || (t.formulas[r] || []).some((v) => !blank(v))) last = r + 1;
  }
  return last;
}
function lastCol(t) {
  let last = 0;
  for (const which of ["values", "formulas"]) for (const row of t[which]) row.forEach((v, i) => { if (!blank(v)) last = Math.max(last, i + 1); });
  return last;
}
// Tiny evaluator: enough for the formula patterns copied into new rows.
function evalFormula(wb, f) {
  let m = f.match(/^=IF\('([^']+)'!([A-Z]+)(\d+)="","",'\1'!\2\3\)$/);
  if (m) return get(wb.tabs[m[1]], Number(m[3]) - 1, colIndex(m[2]));
  m = f.match(/^=IF\(A(\d+)="","",'([^']+)'!([A-Z]+)\1&""\)$/);
  if (m) return String(get(wb.tabs[m[2]], Number(m[1]) - 1, colIndex(m[3])));
  return "";
}
export function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function setFormula(wb, t, r, c, f) {
  put(t, r, c, "formulas", f);
  put(t, r, c, "values", evalFormula(wb, f));
}
function setValue(t, r, c, v) {
  put(t, r, c, "formulas", "");
  put(t, r, c, "values", v);
}

// Insert empty rows into a tab and re-point every reference to it, like Sheets does
// (a range that spans the insertion point grows; refs below it move down).
function insertRows(wb, t, at, n) {
  for (const which of ["values", "formulas"]) {
    while (t[which].length < at) t[which].push([]);
    t[which].splice(at, 0, ...Array.from({ length: n }, () => []));
  }
  t.maxRows += n;
  const fix = (text, sameTab) => {
    const name = t.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`(?:((?:'${name}'|\\b${name})!)|(?<![!\\w$']))(\\$?)([A-Z]{1,3})(\\$?)(\\d+)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?`, "g");
    return text.split('"').map((part, i) => i % 2 ? part : part.replace(re, (m, tabRef, d1, c1, d2, r1, d3, c2, d4, r2) => {
      if (!tabRef && !sameTab) return m;
      const mv = (r) => (Number(r) - 1 >= at ? Number(r) + n : Number(r));
      let a = mv(r1);
      let b = r2 ? (Number(r2) - 1 >= at ? Number(r2) + n : Number(r2)) : null;
      if (r2 && Number(r1) - 1 < at && Number(r2) - 1 >= at) { a = Number(r1); } // range spans the insert: grows
      return `${tabRef || ""}${d1}${c1}${d2}${a}${r2 ? `:${d3}${c2}${d4}${b}` : ""}`;
    })).join('"');
  };
  for (const tab of Object.values(wb.tabs)) {
    tab.formulas = tab.formulas.map((row) => (row || []).map((f) => (f ? fix(f, tab === t) : f)));
  }
  if (wb.validationRefs) wb.validationRefs = wb.validationRefs.map((v) => fix(v, false));
}

// ---------- fake Apps Script services ----------
function appsScriptContext(wb, props) {
  const sheetApi = (t) => {
    const range = (r, c, nr = 1, nc = 1) => ({
      getRow: () => r,
      getColumn: () => c,
      getLastRow: () => r + nr - 1,
      getLastColumn: () => c + nc - 1,
      getSheet: () => sheetApi(t),
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => get(t, r - 1 + i, c - 1 + j))),
      getFormulas: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => get(t, r - 1 + i, c - 1 + j, "formulas"))),
      getFormula: () => get(t, r - 1, c - 1, "formulas"),
      getValue: () => get(t, r - 1, c - 1),
      getDisplayValue: () => String(get(t, r - 1, c - 1)),
      setValue: (v) => {
        if (c > t.maxCols || r > t.maxRows) throw new Error("out of grid");
        if (typeof v === "string" && v.startsWith("=") && t.formats[`${r},${c}`] !== "@") setFormula(wb, t, r - 1, c - 1, v);
        else setValue(t, r - 1, c - 1, v);
      },
      setFormula: (f) => setFormula(wb, t, r - 1, c - 1, f),
      setNumberFormat: (fmt) => { t.formats[`${r},${c}`] = fmt; },
    });
    return {
      getName: () => t.name,
      getSheetId: () => t.sheetId,
      getLastRow: () => lastRow(t),
      getLastColumn: () => lastCol(t),
      getMaxRows: () => t.maxRows,
      getMaxColumns: () => t.maxCols,
      insertColumnsAfter: (_after, n) => { t.maxCols += n; },
      insertRowsAfter: (_after, n) => { t.maxRows += n; },
      getRange: range,
    };
  };
  const ss = { getSheetByName: (n) => (wb.tabs[n] ? sheetApi(wb.tabs[n]) : null) };
  return {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    ContentService: {
      MimeType: { JSON: "application/json" },
      createTextOutput: (s) => ({ content: s, setMimeType() { return this; } }),
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => "Asia/Kolkata" },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 10) },
    sheetApi,
  };
}

export async function startMock({ port = 0, props = {} } = {}) {
  const wb = loadWorkbook();
  const ctx = appsScriptContext(wb, props);
  const code = readFileSync(here("../apps-script/Code.gs"), "utf8");
  vm.createContext(ctx);
  vm.runInContext(code, ctx, { filename: "Code.gs" });
  const stats = { gets: 0, posts: 0, options: 0, restReads: 0, restWrites: 0, writeBodies: [] };
  let down = false;
  const TOKEN = "test-google-token";

  const server = http.createServer(async (req, res) => {
    if (down) { req.socket.destroy(); return; }
    const url = new URL(req.url, "http://x");
    const send = (status, obj) => {
      res.writeHead(status, { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, content-type" });
      res.end(JSON.stringify(obj));
    };
    let body = "";
    for await (const chunk of req) body += chunk;

    // ---- test control (inspect the in-memory sheet) ----
    if (url.pathname === "/__cell") {
      const t = wb.tabs[url.searchParams.get("tab")];
      const m = url.searchParams.get("a1").match(/^([A-Z]+)(\d+)$/);
      send(200, { value: get(t, Number(m[2]) - 1, colIndex(m[1])), formula: get(t, Number(m[2]) - 1, colIndex(m[1]), "formulas"), stats: { ...stats, writeBodies: stats.writeBodies.length } });
      return;
    }
    // ---- Apps Script web app ----
    if (url.pathname === "/exec") {
      if (req.method === "OPTIONS") { stats.options++; res.writeHead(405); res.end(); return; } // Apps Script can't do preflight
      if (req.method === "GET") {
        stats.gets++;
        const out = ctx.doGet({ parameter: Object.fromEntries(url.searchParams) });
        res.writeHead(200, { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" });
        res.end(out.content);
        return;
      }
      if (req.method === "POST") {
        stats.posts++;
        stats.writeBodies.push(body);
        if (!/^text\/plain/.test(req.headers["content-type"] || "")) { send(415, { ok: false, error: "expected text/plain" }); return; }
        const out = ctx.doPost({ postData: { contents: body } });
        res.writeHead(200, { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" });
        res.end(out.content);
        return;
      }
    }

    // ---- Google Sheets REST API ----
    if (url.pathname.startsWith("/v4/spreadsheets/")) {
      if (req.method === "OPTIONS") { res.writeHead(204, { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, content-type", "Access-Control-Allow-Methods": "GET,POST" }); res.end(); return; }
      if (req.headers.authorization !== "Bearer " + TOKEN) { send(401, { error: { message: "Unauthenticated" } }); return; }
      const rest = url.pathname.slice("/v4/spreadsheets/".length);
      const id = rest.split(/[/:]/)[0];
      if (id !== wb.id) { send(404, { error: { message: "not found" } }); return; }
      const parseRange = (a1) => {
        const m = a1.match(/^'?(.*?)'?(?:!([A-Z]+)(\d+))?$/);
        return { tab: m[1].replace(/''/g, "'"), col: m[2] ? colIndex(m[2]) : null, row: m[3] ? Number(m[3]) - 1 : null };
      };
      if (req.method === "GET" && rest === id && /dataValidation/.test(url.searchParams.get("fields") || "")) {
        stats.restReads++;
        const refs = wb.validationRefs || [];
        send(200, { sheets: [{ data: [{ rowData: refs.map((v) => ({ values: [{}, { dataValidation: { condition: { type: "ONE_OF_RANGE", values: [{ userEnteredValue: v }] } } }] })) }] }] });
        return;
      }
      if (req.method === "GET" && rest === id) {
        stats.restReads++;
        send(200, { sheets: Object.values(wb.tabs).map((t) => ({ properties: { sheetId: t.sheetId, title: t.name, gridProperties: { rowCount: t.maxRows, columnCount: t.maxCols } } })) });
        return;
      }
      if (req.method === "GET" && rest === id + "/values:batchGet") {
        stats.restReads++;
        const opt = url.searchParams.get("valueRenderOption");
        const valueRanges = url.searchParams.getAll("ranges").map((r) => {
          const t = wb.tabs[parseRange(r).tab];
          const rows = lastRow(t), cols = lastCol(t);
          const values = [];
          for (let i = 0; i < rows; i++) {
            const row = [];
            for (let j = 0; j < cols; j++) {
              const f = get(t, i, j, "formulas");
              row.push(opt === "FORMULA" && f ? f : get(t, i, j));
            }
            while (row.length && blank(row[row.length - 1])) row.pop();
            values.push(row);
          }
          return { range: r, values };
        });
        send(200, { valueRanges });
        return;
      }
      if (req.method === "POST" && rest === id + "/values:batchUpdate") {
        stats.restWrites++;
        const b = JSON.parse(body);
        stats.writeBodies.push(body);
        for (const d of b.data) {
          const { tab, row, col } = parseRange(d.range);
          const t = wb.tabs[tab];
          if (col + 1 > t.maxCols || row + 1 > t.maxRows) { send(400, { error: { message: `Range (${d.range}) exceeds grid limits` } }); return; }
        }
        for (const d of b.data) {
          const { tab, row, col } = parseRange(d.range);
          const t = wb.tabs[tab];
          const v = d.values[0][0];
          if (b.valueInputOption === "USER_ENTERED" && typeof v === "string" && v.startsWith("=")) setFormula(wb, t, row, col, v);
          else setValue(t, row, col, v);
        }
        send(200, { totalUpdatedCells: b.data.length });
        return;
      }
      const ap = rest.match(/^[^/]+\/values\/(.+):append$/);
      if (req.method === "POST" && ap) {
        stats.restWrites++;
        const b = JSON.parse(body);
        const { tab } = parseRange(decodeURIComponent(ap[1]));
        const t = wb.tabs[tab];
        let r = lastRow(t); // 0-based index of the first free row
        for (const rowVals of b.values) {
          rowVals.forEach((v, c) => { if (v !== "" && v !== null) setValue(t, r, c, v); });
          r++;
        }
        send(200, { updates: { updatedRange: `'${tab}'!A${r}` } });
        return;
      }
      if (req.method === "POST" && rest === id + ":batchUpdate") {
        stats.restWrites++;
        const b = JSON.parse(body);
        const replies = [];
        for (const r of b.requests) {
          if (r.appendDimension) {
            const a = r.appendDimension;
            const t = Object.values(wb.tabs).find((x) => x.sheetId === a.sheetId);
            if (a.dimension === "COLUMNS") t.maxCols += a.length;
            else t.maxRows += a.length;
            replies.push({});
          } else if (r.addSheet) {
            const title = r.addSheet.properties.title;
            if (wb.tabs[title]) { send(400, { error: { message: `A sheet with the name "${title}" already exists.` } }); return; }
            const sheetId = 900000 + Object.keys(wb.tabs).length;
            wb.tabs[title] = { name: title, sheetId, maxRows: 1000, maxCols: 26, values: [], formulas: [], formats: {} };
            stats.addedSheets = [...(stats.addedSheets || []), title];
            replies.push({ addSheet: { properties: { sheetId, title, gridProperties: { rowCount: 1000, columnCount: 26 } } } });
          } else if (r.insertDimension) {
            const d = r.insertDimension.range;
            const t = Object.values(wb.tabs).find((x) => x.sheetId === d.sheetId);
            insertRows(wb, t, d.startIndex, d.endIndex - d.startIndex);
            stats.insertedRows = [...(stats.insertedRows || []), { tab: t.name, at: d.startIndex, n: d.endIndex - d.startIndex }];
            replies.push({});
          } else replies.push({});
        }
        send(200, { replies });
        return;
      }
    }
    send(404, { error: "no route " + req.method + " " + url.pathname });
  });
  await new Promise((r) => server.listen(port, "127.0.0.1", r));
  const base = `http://127.0.0.1:${server.address().port}`;
  return {
    wb,
    stats,
    base,
    execUrl: base + "/exec",
    googleToken: TOKEN,
    setDown(v) { down = v; },
    close: () => new Promise((r) => { server.closeAllConnections?.(); server.close(r); }),
    // helpers for "she edits the sheet by hand"
    cell(tab, a1) {
      const m = a1.match(/^([A-Z]+)(\d+)$/);
      return get(wb.tabs[tab], Number(m[2]) - 1, colIndex(m[1]));
    },
    formula(tab, a1) {
      const m = a1.match(/^([A-Z]+)(\d+)$/);
      return get(wb.tabs[tab], Number(m[2]) - 1, colIndex(m[1]), "formulas");
    },
    handEdit(tab, a1, v) {
      const m = a1.match(/^([A-Z]+)(\d+)$/);
      const t = wb.tabs[tab];
      const r = Number(m[2]) - 1, c = colIndex(m[1]);
      setValue(t, r, c, v);
      // run the real onEdit simple trigger, like Sheets does for a manual edit
      const sh = ctx.sheetApi(t);
      ctx.onEdit({ range: sh.getRange(r + 1, c + 1) });
    },
  };
}

if (process.argv[1] && process.argv[1].endsWith("mock-sheets.mjs")) {
  const m = await startMock({ port: Number(process.env.PORT || 8787) });
  console.log("mock Apps Script:", m.execUrl, " mock Sheets API:", m.base, " google token:", m.googleToken);
}
