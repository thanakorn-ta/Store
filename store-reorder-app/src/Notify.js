/**
 * Notify.js
 * Groups reorder_queue rows by pc_code/owner_email and sends one email per
 * PC with a link into the confirm web app (WebApp.js). Also handles the
 * "no reply in 3 days -> resend" reminder.
 *
 * TODO decisions before turning this on for real users:
 *   - Final email copy / Thai wording (draft below is a starting point)
 *   - Whether to CC the manager or only notify them on the budget-exceeded path
 *   - Confirm master_pc.owner_email / manager_email are filled in and correct
 */

function sendReorderEmails() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.REORDER_QUEUE);
  var rows = readSheetAsObjects_(ss, CONFIG.SHEETS.REORDER_QUEUE);
  var pending = rows.filter(function (r) { return r.status === 'รอส่งอีเมล'; });

  var byPc = {};
  pending.forEach(function (r) {
    var key = r.pc_code + '|' + r.owner_email;
    if (!byPc[key]) byPc[key] = [];
    byPc[key].push(r);
  });

  Object.keys(byPc).forEach(function (key) {
    var items = byPc[key];
    var first = items[0];
    if (!first.owner_email) {
      logActivity_('sendReorderEmails', 'skip', 'no owner_email for ' + first.pc_code);
      return;
    }
    var link = buildConfirmLink_(first.pc_code, items.map(function (i) { return i.sku; }));
    var body = renderReorderEmail_(first.pc_name, items, link);
    MailApp.sendEmail({
      to: first.owner_email,
      subject: 'แจ้งเตือนสั่งซื้อของ ' + first.pc_name + ' (' + first.pc_code + ')',
      htmlBody: body
    });
  });

  markQueueStatus_(sheet, pending, 'ส่งอีเมลแล้ว รอตอบกลับ', new Date());
  logActivity_('sendReorderEmails', 'ok', Object.keys(byPc).length + ' emails sent');
}

function renderReorderEmail_(pcName, items, confirmLink) {
  var rowsHtml = items.map(function (i) {
    return '<tr><td>' + i.name + '</td><td style="text-align:right">' + i.balance +
      '</td><td style="text-align:right">' + i.suggested_qty + '</td></tr>';
  }).join('');

  // TODO: swap this template body for an AI-drafted version once MIN/MAX is
  // validated for a few cycles (see Ai.js draftReorderEmail_).
  return '' +
    '<p>เรียนผู้ดูแล ' + pcName + '</p>' +
    '<p>ระบบตรวจพบว่าของต่อไปนี้เหลือถึงจุดสั่งซื้อ กรุณายืนยัน/แก้จำนวนภายใน 3 วัน:</p>' +
    '<table border="1" cellpadding="4" style="border-collapse:collapse">' +
    '<tr><th>รายการ</th><th>คงเหลือ</th><th>แนะนำสั่งเพิ่ม</th></tr>' + rowsHtml + '</table>' +
    '<p><a href="' + confirmLink + '">คลิกเพื่อยืนยัน/แก้ไขจำนวน</a></p>';
}

/** Any queue rows still "ส่งอีเมลแล้ว รอตอบกลับ" after N days -> resend once. */
function remindUnansweredReorders() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.REORDER_QUEUE);
  var rows = readSheetAsObjects_(ss, CONFIG.SHEETS.REORDER_QUEUE);
  var cutoff = new Date(Date.now() - CONFIG.REMINDER_AFTER_DAYS * 24 * 60 * 60 * 1000);

  var overdue = rows.filter(function (r) {
    return r.status === 'ส่งอีเมลแล้ว รอตอบกลับ' && new Date(r.created_at) < cutoff;
  });
  if (!overdue.length) return;

  var byPc = {};
  overdue.forEach(function (r) {
    var key = r.pc_code + '|' + r.owner_email;
    (byPc[key] = byPc[key] || []).push(r);
  });
  Object.keys(byPc).forEach(function (key) {
    var items = byPc[key];
    var first = items[0];
    var link = buildConfirmLink_(first.pc_code, items.map(function (i) { return i.sku; }));
    MailApp.sendEmail({
      to: first.owner_email,
      subject: '[เตือนซ้ำ] แจ้งเตือนสั่งซื้อของ ' + first.pc_name + ' (' + first.pc_code + ')',
      htmlBody: renderReorderEmail_(first.pc_name, items, link)
    });
  });
  logActivity_('remindUnansweredReorders', 'ok', overdue.length + ' items reminded');
}

function markQueueStatus_(sheet, rowsMatched, newStatus, when) {
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var skuCol = header.indexOf('sku');
  var statusCol = header.indexOf('status');
  var matchedSkus = {};
  rowsMatched.forEach(function (r) { matchedSkus[r.sku] = true; });
  for (var i = 1; i < data.length; i++) {
    if (matchedSkus[data[i][skuCol]]) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
    }
  }
}

function buildConfirmLink_(pcCode, skus) {
  var base = ScriptApp.getService().getUrl(); // available once deployed as web app
  return base + '?pc=' + encodeURIComponent(pcCode) + '&skus=' + encodeURIComponent(skus.join(','));
}
