/**
 * Import.js
 * Step 1-2 of the workflow: pull the 3 Store export files from a Drive
 * folder, convert them (.xls -> Google Sheet via Advanced Drive Service),
 * clean headers, and write normalized rows into raw_* sheets.
 *
 * Store team uploads (same filenames Store already uses):
 *   - ยอดคงเหลือสินค้า (2).xls
 *   - รวมการใช้ของ.xlsx  (or one file per month, see importUsageFolder_)
 *   - รายงานสินค้าถึงจุดสั่งซื้อ (2).xls
 *
 * TODO before first run:
 *   1. Set CONFIG.SPREADSHEET_ID in Config.js
 *   2. Set IMPORT_FOLDER_ID below to the Drive folder Store uploads into
 *   3. Enable "Drive API" advanced service (already declared in appsscript.json)
 */

var IMPORT_FOLDER_ID = 'PUT_DRIVE_FOLDER_ID_HERE';

function importAllFiles() {
  ownerOnly_();
  var folder = DriveApp.getFolderById(IMPORT_FOLDER_ID);
  importBalanceFile_(folder);
  importUsageFolder_(folder);
  importReorderFile_(folder);
  logActivity_('import', 'ok', 'imported balance/usage/reorder files');
}

/**
 * Converts a Drive File (.xls/.xlsx) to a temporary Google Sheet and
 * returns the opened Spreadsheet. Caller is responsible for trashing
 * the temp file when done (see convertOnce/cleanup pattern below).
 */
function convertToGoogleSheet_(file) {
  var resource = {
    title: file.getName() + ' (converted)',
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var converted = Drive.Files.copy(resource, file.getId());
  return { file: converted, ss: SpreadsheetApp.openById(converted.id) };
}

function importBalanceFile_(folder) {
  var files = folder.getFilesByName('ยอดคงเหลือสินค้า (2).xls');
  if (!files.hasNext()) { throw new Error('ไม่พบไฟล์ยอดคงเหลือสินค้าในโฟลเดอร์'); }
  var file = files.next();
  var conv = convertToGoogleSheet_(file);
  try {
    var rows = parseBalanceSheet_(conv.ss.getSheets()[0]);
    writeRows_(CONFIG.SHEETS.RAW_BALANCE,
      ['sku', 'name', 'location', 'balance', 'unit', 'unit_cost', 'value', 'pc_code', 'pc_name'],
      rows);
  } finally {
    DriveApp.getFileById(conv.file.id).setTrashed(true);
  }
}

/**
 * The balance export has a "PGxx .../PCxxxx <ชื่อ>" header row above each
 * block of items, followed by a column-header row, then data rows, with
 * blank separator rows between blocks. This walks the sheet and expands
 * each item row with its current PC code/name from the last seen header.
 */
function parseBalanceSheet_(sheet) {
  var data = sheet.getDataRange().getValues();
  var out = [];
  var pcCode = '', pcName = '';
  var pcHeaderRe = /PC(\d+)\s+(.*)$/;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var joined = row.join(' ').trim();
    var m = joined.match(pcHeaderRe);
    if (m) { pcCode = 'PC' + m[1]; pcName = m[2].trim(); continue; }

    // Item rows have a SKU like P###### in one of the first few columns
    var sku = findSku_(row);
    if (!sku) continue;

    var idx = row.indexOf(sku);
    var name = String(row[idx + 2] || '').trim();
    var location = String(row[idx + 5] || '').trim();
    var balance = Number(row[idx + 6]) || 0;
    var unit = String(row[idx + 8] || '').trim();
    var unitCost = Number(row[idx + 9]) || 0;
    var value = Number(row[idx + 11]) || 0;

    out.push([sku, name, location, balance, unit, unitCost, value, pcCode, pcName]);
  }
  return out;
}

function findSku_(row) {
  for (var c = 0; c < row.length; c++) {
    if (typeof row[c] === 'string' && /^P\d{5,}$/.test(row[c].trim())) return row[c].trim();
  }
  return null;
}

/**
 * Usage files: one per month (การใช้งานเดือน<ไทย>.xls) or the combined
 * รวมการใช้ของ.xlsx query table. Reads whichever is present and appends
 * (sku, month_number, qty, unit, price, total, source_file) rows.
 * Existing rows for the same (sku, month_number) are replaced so re-running
 * an import is idempotent.
 */
function importUsageFolder_(folder) {
  var THAI_MONTHS = {
    'ม.ค.': 1, 'ก.พ.': 2, 'มี.ค.': 3, 'เม.ย.': 4, 'พ.ค.': 5, 'มิ.ย.': 6,
    'ก.ค.': 7, 'ส.ค.': 8, 'ก.ย.': 9, 'ต.ค.': 10, 'พ.ย.': 11, 'ธ.ค.': 12
  };

  var rows = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName();
    var monthMatch = null;
    for (var key in THAI_MONTHS) {
      if (name.indexOf(key) !== -1) { monthMatch = THAI_MONTHS[key]; break; }
    }
    if (!monthMatch || name.indexOf('การใช้งานเดือน') === -1) continue;

    var conv = convertToGoogleSheet_(f);
    try {
      var data = conv.ss.getSheets()[0].getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var sku = findSku_(row);
        if (!sku) continue;
        var idx = row.indexOf(sku);
        var itemName = String(row[idx + 1] || '').trim();
        var qty = Number(String(row[idx + 4]).replace(/[ , ]/g, '')) || 0;
        var unit = String(row[idx + 5] || '').trim();
        var price = Number(String(row[idx + 6]).replace(/[ , ]/g, '')) || 0;
        var total = Number(String(row[idx + 7]).replace(/[ , ]/g, '')) || 0;
        rows.push([sku, itemName, monthMatch, qty, unit, price, total, name]);
      }
    } finally {
      DriveApp.getFileById(conv.file.id).setTrashed(true);
    }
  }

  writeRows_(CONFIG.SHEETS.RAW_USAGE,
    ['sku', 'name', 'month', 'qty', 'unit', 'price', 'total', 'source_file'],
    rows);
}

function importReorderFile_(folder) {
  var files = folder.getFilesByName('รายงานสินค้าถึงจุดสั่งซื้อ (2).xls');
  if (!files.hasNext()) { return; } // optional: used for cross-checking only
  var file = files.next();
  var conv = convertToGoogleSheet_(file);
  try {
    var data = conv.ss.getSheets()[0].getDataRange().getValues();
    var rows = [];
    var pcCode = '', pcName = '';
    var pcHeaderRe = /PC(\d+)\s+(.*)$/;
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var joined = row.join(' ').trim();
      var m = joined.match(pcHeaderRe);
      if (m) { pcCode = 'PC' + m[1]; pcName = m[2].trim(); continue; }
      var sku = findSku_(row);
      if (!sku) continue;
      var idx = row.indexOf(sku);
      rows.push([sku, String(row[idx + 3] || '').trim(), Number(row[idx + 4]) || 0,
        Number(row[idx + 5]) || 0, Number(row[idx + 7]) || 0, pcCode, pcName]);
    }
    writeRows_(CONFIG.SHEETS.RAW_REORDER,
      ['sku', 'name', 'balance', 'erp_min', 'suggest_buy', 'pc_code', 'pc_name'],
      rows);
  } finally {
    DriveApp.getFileById(conv.file.id).setTrashed(true);
  }
}

function writeRows_(sheetName, header, rows) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}
