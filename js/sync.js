// Sync engine: offline write queue + reads on load / manual refresh only.
// No polling and no focus/visibility refresh (by design).
import { planWrites } from "./sheet-model.js";
import { AppsScriptTransport, GoogleSheetsTransport, TransportError } from "./transports.js";

const CFG_KEY = "bakery.v2.sync";
const QUEUE_KEY = "bakery.v2.queue";
const META_KEY = "bakery.v2.syncMeta";
export const DEFAULT_SHEET_ID = "1LgfA7HGWy-KA5PcA6qQ7d1omzuJwiWKXIgpjEIcvyQI";

const clone = (v) => JSON.parse(JSON.stringify(v));

export function loadSyncConfig() {
  try {
    return { mode: "local", url: "", token: "2468", clientId: "", spreadsheetId: DEFAULT_SHEET_ID, ...(JSON.parse(localStorage.getItem(CFG_KEY) || "{}")) };
  } catch {
    return { mode: "local", url: "", token: "2468", clientId: "", spreadsheetId: DEFAULT_SHEET_ID };
  }
}
export function saveSyncConfig(cfg) { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); }

export function makeTransport(cfg, extra = {}) {
  if (cfg.mode === "appsscript" && cfg.url) return new AppsScriptTransport({ url: cfg.url, token: cfg.token, ...extra });
  if (cfg.mode === "google" && cfg.spreadsheetId) return new GoogleSheetsTransport({ spreadsheetId: cfg.spreadsheetId, ...extra });
  return null;
}

function sameTarget(a, b) {
  if (a.kind !== b.kind) return false;
  if (a.kind === "append") return a.queuedAt === b.queuedAt;
  if (a.kind === "set") return a.tab === b.tab && a.key === b.key && a.field === b.field;
  if (a.kind === "order") return a.id === b.id;
  return false;
}

export class SyncEngine {
  constructor({ transport, onState, onStatus, onFlushed, debounceMs = 600 }) {
    this.transport = transport;
    this.onState = onState || (() => {});
    this.onStatus = onStatus || (() => {});
    this.onFlushed = onFlushed || (() => {});
    this.debounceMs = debounceMs;
    this.queue = this._loadQueue();
    const meta = JSON.parse(localStorage.getItem(META_KEY) || "{}");
    this.status = { state: transport ? "idle" : "local", lastSynced: meta.lastSynced || null, message: "", conflicts: meta.conflicts || [], pending: this.queue.length };
    this._timer = null;
    this._running = null;
  }
  _loadQueue() {
    try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]"); } catch { return []; }
  }
  _saveQueue() {
    localStorage.setItem(QUEUE_KEY, JSON.stringify(this.queue));
    this.status.pending = this.queue.length;
  }
  _set(patch) {
    Object.assign(this.status, patch, { pending: this.queue.length });
    localStorage.setItem(META_KEY, JSON.stringify({ lastSynced: this.status.lastSynced, conflicts: this.status.conflicts.slice(-20) }));
    this.onStatus({ ...this.status });
  }

  enqueue(op) {
    const i = op.kind === "append" ? -1 : this.queue.findIndex((q) => sameTarget(q, op));
    if (i >= 0) {
      const prev = this.queue[i];
      if (op.kind === "set") this.queue[i] = { ...op, base: prev.base };
      else if (op.kind === "order") this.queue[i] = { ...op, fields: { ...prev.fields, ...op.fields }, isNew: prev.isNew || op.isNew };
    } else this.queue.push({ ...op, queuedAt: new Date().toISOString() + "#" + Math.random().toString(36).slice(2, 7) });
    this._saveQueue();
    if (!this.transport) return;
    this._set({ state: this.status.state === "offline" ? "offline" : "pending" });
    clearTimeout(this._timer);
    this._timer = setTimeout(() => this.flush().catch(() => {}), this.debounceMs);
  }

  /** Send queued edits. Reads the sheet first (to locate rows/cells) but does not re-render. */
  async flush() {
    if (!this.transport || !this.queue.length) return { written: 0 };
    if (this._running) { await this._running.catch(() => {}); if (!this.queue.length) return { written: 0 }; }
    const run = (async () => {
      this._set({ state: "syncing", message: "" });
      const ops = clone(this.queue);
      try {
        const grid = await this.transport.read();
        const pristine = clone(grid);
        const plan = planWrites(grid, ops);
        const res = await this.transport.write(plan.cells, pristine);
        const skipped = (res.results || []).map((r, i) => (r.skipped ? { cell: plan.cells[i], reason: r.skipped } : null)).filter(Boolean);
        // every sent op was either written or resolved as a conflict: drop it,
        // unless it was edited again while the flush was running
        // A new row whose key cell was taken by someone else at the same moment
        // (another phone appending) is retried on the next flush, at the next free row.
        const retry = new Set(skipped.filter((x) => x.cell.keyCell && x.reason === "notEmpty").map((x) => x.cell.opRef));
        this.queue = this.queue.filter((q) => {
          if (retry.has(q.queuedAt)) return true;
          const sent = ops.find((o) => o.queuedAt === q.queuedAt && sameTarget(o, q));
          return !sent || JSON.stringify(sent) !== JSON.stringify(q);
        });
        this._saveQueue();
        const conflicts = [
          ...this.status.conflicts,
          ...plan.conflicts.map((c) => ({ at: new Date().toISOString(), what: describe(c.op), reason: c.reason })),
          ...skipped.filter((x) => !x.cell.boot && !x.cell.opRef && !(x.reason === "moved" && x.cell.ifEmpty)).map((x) => ({
            at: new Date().toISOString(),
            what: `${x.cell.tab} row ${x.cell.row}`,
            reason: x.reason === "formula" ? "cell has a formula; not overwritten" : x.reason === "moved" ? "row moved in the sheet; edit again" : "cell already filled; not overwritten",
          })),
        ].slice(-20);
        this.onFlushed(plan.done.filter((o) => !retry.has(o.queuedAt)));
        if (retry.size) setTimeout(() => this.flush().catch(() => {}), 1500);
        this._set({ state: "synced", lastSynced: new Date().toISOString(), conflicts, message: plan.conflicts.length ? `${plan.conflicts.length} change(s) not applied — see Settings` : "" });
        return { written: plan.cells.length, cells: plan.cells, conflicts: plan.conflicts, results: res.results };
      } catch (e) {
        this._fail(e);
        throw e;
      }
    })();
    this._running = run;
    try { return await run; } finally { this._running = null; }
  }

  stop() {
    clearTimeout(this._timer);
  }

  /** Full fresh read of every tab. Called on page load and by the refresh button only. */
  async pull() {
    if (!this.transport) return null;
    try {
      if (this.queue.length) await this.flush();
      this._set({ state: "syncing" });
      const grid = await this.transport.read();
      this.onState(grid);
      this._set({ state: this.queue.length ? "pending" : "synced", lastSynced: new Date().toISOString() });
      return grid;
    } catch (e) {
      this._fail(e);
      return null;
    }
  }

  _fail(e) {
    const kind = e instanceof TransportError ? e.kind : "error";
    this._set({ state: kind === "offline" ? "offline" : kind === "auth" ? "signin" : "error", message: e.message || String(e) });
  }
}

function describe(op) {
  if (op.kind === "set") return `${op.tab}: ${op.key} → ${op.field}`;
  if (op.kind === "order") return `Order ${op.id}`;
  if (op.kind === "append") return `${op.tab}: add ${op.record?.name || op.record?.ingredient || "row"}`;
  return op.kind;
}
