/**
 * WebApp.js
 * doGet serves the web UI (Ui.js) for every URL, including links from emails
 * (?req=REQ-...). Sign-in happens inside the page (Members.js / Api.js).
 *
 * The older per-PC confirm page (confirm.html + submitConfirmation, fed by
 * Notify.js/reorder_queue) is superseded by the request workflow in
 * Requests.js; submitConfirmation is kept for the legacy pipeline and can
 * only run from the editor/trigger.
 */

// Bump when deploying so <exec url>?health=1 shows which code the deployment runs
var APP_VERSION = '2026-09-22';

function doGet(e) {
  // <exec url>?health=1 → JSON: quick check that the deployment runs this code (no sign-in needed)
  if (e && e.parameter && e.parameter.health) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'store-reorder-ai', version: APP_VERSION }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return renderReorderUi_();
}

/**
 * Called from confirm.html via google.script.run.
 * payload: [{ sku, confirmedQty, note }, ...]
 */
function submitConfirmation(pcCode, payload) {
  ownerOnly_();
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
