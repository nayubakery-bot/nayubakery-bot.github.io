// Two ways for the app to talk to the Google Sheet. Both expose:
//   read()                -> grid { tabs:{name:{values,formulas}}, meta }
//   write(cells, pristine)-> { results:[{ok}|{skipped:reason}] }
// Neither ever clears a tab or rewrites a whole range: writes are single cells.
import { TABS, colLetter } from "./sheet-model.js";

export class TransportError extends Error {
  constructor(message, kind = "error") {
    super(message);
    this.kind = kind; // "offline" | "auth" | "error"
  }
}

function netFail(e) {
  const offline = typeof navigator !== "undefined" && navigator.onLine === false;
  return new TransportError(offline ? "No internet connection" : "Could not reach the sheet: " + (e && e.message), offline ? "offline" : "offline");
}

/* ---------- 1. Google Apps Script web app bound to the sheet ---------- */
export class AppsScriptTransport {
  constructor({ url, token, fetchImpl }) {
    this.url = String(url || "").trim();
    this.token = String(token || "");
    this.fetch = fetchImpl || ((...a) => fetch(...a));
  }
  async _json(res) {
    let body;
    try { body = await res.json(); } catch { throw new TransportError("Sheet script did not return JSON. Check the web app URL and that access is set to Anyone.", "error"); }
    if (!body.ok) throw new TransportError(body.error || "Script error", body.error === "Bad token" ? "auth" : "error");
    return body;
  }
  async read() {
    const u = this.url + (this.url.includes("?") ? "&" : "?") + "action=read&token=" + encodeURIComponent(this.token) + "&_=" + Date.now();
    let res;
    try { res = await this.fetch(u, { method: "GET", cache: "no-store", redirect: "follow" }); } catch (e) { throw netFail(e); }
    const body = await this._json(res);
    return { tabs: body.tabs, meta: body.meta || {} };
  }
  async write(cells) {
    if (!cells.length) return { results: [] };
    let res;
    try {
      res = await this.fetch(this.url, {
        method: "POST",
        cache: "no-store",
        redirect: "follow",
        // text/plain keeps this a "simple" request: no CORS preflight (Apps Script can't answer OPTIONS).
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "write", token: this.token, cells }),
      });
    } catch (e) { throw netFail(e); }
    const body = await this._json(res);
    return { results: body.results || [] };
  }
}

/* ---------- 2. Direct Google Sheets API with "Sign in with Google" ---------- */
const SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const TOKEN_KEY = "bakery.v2.googleToken";

let gisLoading = null;
function loadGis() {
  if (typeof window === "undefined") return Promise.reject(new Error("no window"));
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  if (!gisLoading) {
    gisLoading = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://accounts.google.com/gsi/client";
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => { gisLoading = null; reject(new Error("Could not load Google sign-in")); };
      document.head.appendChild(s);
    });
  }
  return gisLoading;
}

export function cachedGoogleToken() {
  try {
    const t = JSON.parse(localStorage.getItem(TOKEN_KEY) || "null");
    if (t && t.token && t.exp > Date.now() + 60000) return t.token;
  } catch {}
  return null;
}
export function clearGoogleToken() { localStorage.removeItem(TOKEN_KEY); }

/** Must be called from a tap (user gesture) so the Google popup isn't blocked. */
export async function signInWithGoogle(clientId) {
  await loadGis();
  return new Promise((resolve, reject) => {
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPE,
      callback: (resp) => {
        if (resp.error) return reject(new Error(resp.error_description || resp.error));
        localStorage.setItem(TOKEN_KEY, JSON.stringify({ token: resp.access_token, exp: Date.now() + Number(resp.expires_in || 3600) * 1000 }));
        resolve(resp.access_token);
      },
      error_callback: (err) => reject(new Error(err?.message || err?.type || "Sign-in cancelled")),
    });
    client.requestAccessToken({ prompt: cachedGoogleToken() ? "" : "consent" });
  });
}

export class GoogleSheetsTransport {
  constructor({ spreadsheetId, getToken, fetchImpl, base }) {
    this.id = spreadsheetId;
    this.getToken = getToken || (() => cachedGoogleToken());
    this.fetch = fetchImpl || ((...a) => fetch(...a));
    this.base = (base || "https://sheets.googleapis.com").replace(/\/$/, "");
    this.meta = null;
  }
  async _call(path, opts = {}) {
    const token = await this.getToken();
    if (!token) throw new TransportError("Sign in with Google to sync", "auth");
    let res;
    try {
      res = await this.fetch(this.base + path, {
        ...opts,
        cache: "no-store",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", ...(opts.headers || {}) },
      });
    } catch (e) { throw netFail(e); }
    if (res.status === 401 || res.status === 403) {
      clearGoogleToken();
      throw new TransportError("Google sign-in expired or has no access to this sheet", "auth");
    }
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new TransportError(body?.error?.message || "Sheets API error " + res.status, "error");
    return body;
  }
  async read() {
    const meta = await this._call(`/v4/spreadsheets/${this.id}?fields=sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))`);
    this.meta = {};
    for (const s of meta.sheets || []) {
      this.meta[s.properties.title] = {
        sheetId: s.properties.sheetId,
        rowCount: s.properties.gridProperties?.rowCount || 1000,
        columnCount: s.properties.gridProperties?.columnCount || 26,
      };
    }
    const tabs = TABS.filter((t) => this.meta[t]);
    const q = tabs.map((t) => "ranges=" + encodeURIComponent(`'${t}'`)).join("&");
    const [vals, forms] = await Promise.all([
      this._call(`/v4/spreadsheets/${this.id}/values:batchGet?${q}&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER`),
      this._call(`/v4/spreadsheets/${this.id}/values:batchGet?${q}&valueRenderOption=FORMULA`),
    ]);
    const out = {};
    tabs.forEach((t, i) => {
      const values = vals.valueRanges?.[i]?.values || [];
      const f = forms.valueRanges?.[i]?.values || [];
      out[t] = {
        values,
        formulas: f.map((row) => row.map((c) => (typeof c === "string" && c.startsWith("=") ? c : ""))),
      };
    });
    return { tabs: out, meta: this.meta };
  }
  async write(cells, pristine) {
    if (!cells.length) return { results: [] };
    const results = new Array(cells.length);
    // New Orders rows go through values:append, which Google applies atomically
    // (two phones adding an order at the same moment each get their own row).
    const appendRows = new Map();
    cells.forEach((c) => {
      if (c.tab === "Orders" && c.keyCell) appendRows.set(c.tab + "!" + c.row, { tab: c.tab, row: c.row, cells: [] });
    });
    const rest = [];
    cells.forEach((c, i) => {
      const g = appendRows.get(c.tab + "!" + c.row);
      if (g && c.ifEmpty && !c.formula) g.cells.push({ c, i });
      else rest.push({ c, i });
    });
    for (const g of appendRows.values()) {
      const width = Math.max(...g.cells.map((x) => x.c.col));
      const row = new Array(width).fill("");
      for (const { c } of g.cells) row[c.col - 1] = c.value;
      await this._call(`/v4/spreadsheets/${this.id}/values/${encodeURIComponent(`'${g.tab}'!A1`)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
        method: "POST",
        body: JSON.stringify({ values: [row] }),
      });
      for (const { i } of g.cells) results[i] = { ok: true, appended: true };
    }
    const valueData = [];
    const formulaData = [];
    const grow = {};
    const pending = new Map(); // cells written earlier in this same batch
    const at = (tabName, r, c) => {
      const k = tabName + "!" + r + "!" + c;
      if (pending.has(k)) return pending.get(k);
      return pristine?.tabs?.[tabName]?.values?.[r - 1]?.[c - 1];
    };
    // Guards (the API has no per-cell condition, so check against the fresh read).
    for (const { c, i } of rest) {
      const t = pristine?.tabs?.[c.tab] || { values: [], formulas: [] };
      const cur = t.values[c.row - 1]?.[c.col - 1];
      const curF = t.formulas?.[c.row - 1]?.[c.col - 1];
      if (curF) { results[i] = { skipped: "formula" }; continue; }
      if (c.ifEmpty && cur !== undefined && cur !== "" && cur !== null) { results[i] = { skipped: "notEmpty" }; continue; }
      if (c.expectKey) {
        const k = at(c.tab, c.row, c.expectKey.col);
        if (String(k ?? "") !== String(c.expectKey.value)) { results[i] = { skipped: "moved" }; continue; }
      }
      const m = this.meta?.[c.tab];
      if (m) {
        const g = (grow[c.tab] = grow[c.tab] || { rows: 0, cols: 0 });
        g.rows = Math.max(g.rows, c.row - m.rowCount);
        g.cols = Math.max(g.cols, c.col - m.columnCount);
      }
      const range = `'${c.tab}'!${colLetter(c.col)}${c.row}`;
      pending.set(c.tab + "!" + c.row + "!" + c.col, c.formula ? "" : c.value);
      if (c.formula) formulaData.push({ range, values: [[c.formula]] });
      else valueData.push({ range, values: [[c.value]] });
      results[i] = { ok: true };
    }
    const requests = [];
    for (const [tabName, g] of Object.entries(grow)) {
      const sheetId = this.meta[tabName].sheetId;
      if (g.cols > 0) requests.push({ appendDimension: { sheetId, dimension: "COLUMNS", length: g.cols } });
      if (g.rows > 0) requests.push({ appendDimension: { sheetId, dimension: "ROWS", length: g.rows + 50 } });
    }
    if (requests.length) await this._call(`/v4/spreadsheets/${this.id}:batchUpdate`, { method: "POST", body: JSON.stringify({ requests }) });
    // RAW keeps text like "4-6" or "2026-11-15" exactly as typed; formulas go USER_ENTERED.
    if (valueData.length) await this._call(`/v4/spreadsheets/${this.id}/values:batchUpdate`, { method: "POST", body: JSON.stringify({ valueInputOption: "RAW", data: valueData }) });
    if (formulaData.length) await this._call(`/v4/spreadsheets/${this.id}/values:batchUpdate`, { method: "POST", body: JSON.stringify({ valueInputOption: "USER_ENTERED", data: formulaData }) });
    return { results };
  }
}
