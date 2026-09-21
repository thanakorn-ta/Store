/**
 * WebApp.js
 * Confirmation page (step 6 in the workflow): a PC owner opens their emailed
 * link, sees the suggested quantities, can edit/confirm/reject per line, and
 * submits. Deploy via `clasp deploy` (or the Apps Script editor) as a Web App
 * with access "Anyone within domain" (see appsscript.json).
 */

function doGet(e) {
  // Emailed confirm links carry ?pc=...&skus=...; anything else is the main UI (Ui.js)
  if (!e.parameter.pc) return renderReorderUi_();

  var pcCode = e.parameter.pc || '';
  var skus = (e.parameter.skus || '').split(',').filter(String);

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var queue = readSheetAsObjects_(ss, CONFIG.SHEETS.REORDER_QUEUE);
  var items = queue.filter(function (r) { return r.pc_code === pcCode && skus.indexOf(r.sku) !== -1; });

  var template = HtmlService.createTemplateFromFile('confirm');
  template.pcCode = pcCode;
  template.items = items;
  return template.evaluate().setTitle('ยืนยันรายการสั่งซื้อ');
}

/**
 * Called from confirm.html via google.script.run.
 * payload: [{ sku, confirmedQty, note }, ...]
 */
function submitConfirmation(pcCode, payload) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.REORDER_QUEUE);
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var col = {};
  header.forEach(function (h, i) { col[h] = i; });

  var bySku = {};
  payload.forEach(function (p) { bySku[p.sku] = p; });

  for (var i = 1; i < data.length; i++) {
    var sku = data[i][col.sku];
    var rowPc = data[i][col.pc_code];
    if (rowPc !== pcCode || !bySku[sku]) continue;
    var p = bySku[sku];
    var status = (Number(p.confirmedQty) > 0) ? 'ยืนยันแล้ว' : 'ปฏิเสธ';
    sheet.getRange(i + 1, col.status + 1).setValue(status);
    sheet.getRange(i + 1, col.confirmed_qty + 1).setValue(p.confirmedQty);
    sheet.getRange(i + 1, col.confirmed_at + 1).setValue(new Date());
  }

  logActivity_('submitConfirmation', 'ok', pcCode + ': ' + payload.length + ' lines');
  return { ok: true };
}
