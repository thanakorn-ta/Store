/**
 * SheetUtil.js
 * Small helpers shared by every module that reads/writes the database Sheet.
 */

function db_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/** Returns the sheet, creating it with a header row if missing. */
function ensureSheet_(ss, name, header) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Like ensureSheet_, but also appends any header columns that an older
 * version of the sheet is missing, so new fields never shift existing data.
 * Returns { sheet, header } where header is the sheet's actual column order.
 */
function ensureHeader_(ss, name, header) {
  var sheet = ensureSheet_(ss, name, header);
  var lastCol = Math.max(1, sheet.getLastColumn ? sheet.getLastColumn() : header.length);
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String).filter(String);
  var missing = header.filter(function (h) { return current.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    current = current.concat(missing);
  }
  return { sheet: sheet, header: current };
}

/** Appends one object as a row, placing each field in its named column. */
function appendObject_(ss, name, header, obj) {
  var h = ensureHeader_(ss, name, header);
  h.sheet.appendRow(h.header.map(function (k) { return obj[k] == null ? '' : obj[k]; }));
}

/** Appends many objects in one write. */
function appendObjects_(ss, name, header, objs) {
  if (!objs.length) return;
  var h = ensureHeader_(ss, name, header);
  var rows = objs.map(function (o) { return h.header.map(function (k) { return o[k] == null ? '' : o[k]; }); });
  h.sheet.getRange(h.sheet.getLastRow() + 1, 1, rows.length, h.header.length).setValues(rows);
}

function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/**
 * Guard for functions meant for the Apps Script editor or time triggers only.
 * The Web App is deployed for "Anyone", so every top-level function without a
 * trailing underscore can also be invoked from a browser via google.script.run.
 * Editor runs and installable triggers execute as the owner, so the active and
 * effective users match; anonymous or other visitors don't.
 */
function ownerOnly_() {
  var active = String(Session.getActiveUser().getEmail() || '');
  if (!active || active !== String(Session.getEffectiveUser().getEmail() || '')) {
    throw new Error('ฟังก์ชันนี้รันได้จาก Apps Script editor / trigger เท่านั้น');
  }
}
