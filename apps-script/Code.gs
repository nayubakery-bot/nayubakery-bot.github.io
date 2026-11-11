/**
 * Rise & Whisk — Google Apps Script bridge for the Bakery Workbook.
 * Paste into Extensions > Apps Script of the Bakery Workbook (bound script).
 *
 * GET  ?action=read&token=XXXX         -> every tab's values + formulas
 * GET  ?action=ping&token=XXXX         -> { ok:true }
 * POST (text/plain JSON) {action:"write", token, cells:[...]}
 *      each cell: {tab,row,col,value} or {tab,row,col,formula}
 *      optional: ifEmpty (only write a blank cell), expectKey {col,value}
 *      (only write if that row's key cell still matches).
 *
 * Safety rules enforced here, whatever the app sends:
 *  - a cell that holds a formula is never overwritten;
 *  - nothing is cleared, no rows/columns are deleted, no whole ranges rewritten;
 *  - only the listed tabs can be touched.
 *
 * Token: Project Settings > Script properties > APP_TOKEN (default 2468).
 */

var TABS = ['Raw Materials', 'Recipes', 'Settings', 'Inventory', 'Orders', 'Recipe Library'];
var DEFAULT_TOKEN = '2468';

function appToken_() {
  var t = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
  return t ? String(t) : DEFAULT_TOKEN;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cellOut_(v) {
  if (v instanceof Date) {
    // Dates go back as yyyy-mm-dd (or full ISO when there is a time part)
    var tz = Session.getScriptTimeZone();
    var hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz, hasTime ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
  }
  return v;
}

function readAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = {};
  var meta = {};
  for (var i = 0; i < TABS.length; i++) {
    var sh = ss.getSheetByName(TABS[i]);
    if (!sh) continue;
    var rows = Math.max(sh.getLastRow(), 1);
    var cols = Math.max(sh.getLastColumn(), 1);
    var rng = sh.getRange(1, 1, rows, cols);
    var values = rng.getValues();
    var formulas = rng.getFormulas();
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) values[r][c] = cellOut_(values[r][c]);
    }
    tabs[TABS[i]] = { values: values, formulas: formulas };
    meta[TABS[i]] = { sheetId: sh.getSheetId(), rowCount: sh.getMaxRows(), columnCount: sh.getMaxColumns() };
  }
  return { tabs: tabs, meta: meta };
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (String(p.token || '') !== appToken_()) return json_({ ok: false, error: 'Bad token' });
  if (p.action === 'ping') return json_({ ok: true });
  if (p.action === 'read' || !p.action) {
    var data = readAll_();
    return json_({ ok: true, tabs: data.tabs, meta: data.meta, at: new Date().toISOString() });
  }
  return json_({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Bad JSON' });
  }
  if (String(body.token || '') !== appToken_()) return json_({ ok: false, error: 'Bad token' });
  if (body.action !== 'write') return json_({ ok: false, error: 'Unknown action' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return json_({ ok: true, results: writeCells_(body.cells || []) });
  } finally {
    lock.releaseLock();
  }
}

function writeCells_(cells) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  for (var i = 0; i < cells.length; i++) {
    var c = cells[i];
    try {
      if (TABS.indexOf(c.tab) < 0) { results.push({ skipped: 'tab' }); continue; }
      var sh = ss.getSheetByName(c.tab);
      var row = Number(c.row), col = Number(c.col);
      if (!sh || !(row >= 1) || !(col >= 1)) { results.push({ skipped: 'address' }); continue; }
      if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
      if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows() + 50);
      var rng = sh.getRange(row, col);
      if (rng.getFormula()) { results.push({ skipped: 'formula' }); continue; }
      if (c.ifEmpty && rng.getValue() !== '') { results.push({ skipped: 'notEmpty' }); continue; }
      if (c.expectKey) {
        var k = sh.getRange(row, Number(c.expectKey.col)).getDisplayValue();
        if (String(k) !== String(c.expectKey.value)) { results.push({ skipped: 'moved' }); continue; }
      }
      if (c.formula) {
        rng.setFormula(String(c.formula));
      } else {
        var v = c.value === null || c.value === undefined ? '' : c.value;
        // Keep text such as "4-6" or "2026-11-15" as typed (don't let Sheets turn it into a date).
        if (typeof v === 'string' && /^[0-9=+-]/.test(v) && isNaN(Number(v))) rng.setNumberFormat('@');
        rng.setValue(v);
      }
      results.push({ ok: true });
    } catch (err) {
      results.push({ skipped: 'error', error: String(err) });
    }
  }
  return results;
}

/**
 * Simple trigger: when she edits an order by hand in the sheet, stamp "Updated at"
 * so the app's last-write-wins check knows the sheet copy is newer.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== 'Orders' || e.range.getRow() < 2) return;
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = -1;
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).toLowerCase().replace(/[^a-z]/g, '') === 'updatedat') { col = i + 1; break; }
  }
  if (col < 0) return;
  if (e.range.getColumn() <= col && e.range.getLastColumn() >= col) return; // edited the stamp itself
  var stamp = new Date().toISOString();
  for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
    var cell = sh.getRange(r, col);
    cell.setNumberFormat('@');
    cell.setValue(stamp);
  }
}
