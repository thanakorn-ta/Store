/**
 * Ui.js
 * Serves the same reorder-comparison page as GitHub Pages (web/) from this
 * Apps Script Web App. The ui_*.html and Index.html files are GENERATED from
 * web/ by tools/build-appsscript.ps1 — edit web/, then rebuild; never edit
 * the generated files by hand.
 *
 * Routing lives in WebApp.js doGet(): links with ?pc=... (from the reorder
 * emails) open the confirm page, everything else opens this UI.
 */

function renderReorderUi_() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Store Reorder AI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Used by Index.html as <?!= include('ui_app'); ?> */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/**
 * Called from the page's "ผู้ช่วย AI" tab. Key stays in Script Properties
 * (ANTHROPIC_API_KEY), so viewers of the Web App never see it.
 */
function askClaudeFromUi(system, userContent) {
  if (String(userContent).length > 200000) throw new Error('ข้อมูลที่ส่งให้ AI ยาวเกินไป');
  return callClaude_(system, userContent, 4000);
}

/**
 * Called from the cart's "บันทึกลง Google Sheet". Appends the selected lines
 * to reorder_queue with status "รอส่งอีเมล", so the existing Notify.js /
 * confirm-page flow picks them up exactly like auto-queued items.
 * rows: objects produced by cartSheetRows() in web/app.js (Thai keys).
 */
function saveSelectionToQueue(rows) {
  if (!rows || !rows.length) throw new Error('ไม่มีรายการ');
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'PUT_SPREADSHEET_ID_HERE') {
    throw new Error('ยังไม่ได้ตั้ง CONFIG.SPREADSHEET_ID ใน Config.js');
  }
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var pcByCode = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.MASTER_PC).forEach(function (p) { pcByCode[p.pc_code] = p; });

  var header = ['created_at', 'sku', 'name', 'pc_code', 'pc_name', 'owner_email',
    'manager_email', 'balance', 'min', 'max', 'suggested_qty', 'unit_cost',
    'suggested_value', 'status', 'confirmed_qty', 'confirmed_at'];
  var sheet = ss.getSheetByName(CONFIG.SHEETS.REORDER_QUEUE) || ss.insertSheet(CONFIG.SHEETS.REORDER_QUEUE);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, header.length).setValues([header]);

  var now = new Date();
  var out = rows
    .filter(function (r) { return Number(r['จำนวนสั่ง']) > 0; })
    .map(function (r) {
      var pc = pcByCode[r['PC']] || {};
      var qty = Number(r['จำนวนสั่ง']);
      var cost = Number(r['ราคา/หน่วย']) || 0;
      return [now, r['รหัสสินค้า'], r['สินค้า'], r['PC'] || '', pc.pc_name || r['ชื่อ PC'] || '',
        pc.owner_email || '', pc.manager_email || '',
        Number(r['คงเหลือ']) || 0, '', '', qty, cost, round1_(qty * cost),
        'รอส่งอีเมล', '', ''];
    });
  if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, header.length).setValues(out);

  logActivity_('saveSelectionToQueue', 'ok', out.length + ' items from web UI by ' + Session.getActiveUser().getEmail());
  return { count: out.length };
}
