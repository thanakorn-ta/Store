/**
 * Reorder.js
 * Turns the minmax_calculated sheet into a per-PC reorder queue.
 * Only "เบิกประจำ" items with status "ต้องสั่งซื้อ" are auto-queued.
 * "เบิกตามงาน" items are never auto-queued; Notify.js should still let a
 * PC owner request one manually if they know a job is coming up.
 */

function buildReorderQueue() {
  ownerOnly_();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var minmax = readSheetAsObjects_(ss, CONFIG.SHEETS.MINMAX);
  var masterPc = readSheetAsObjects_(ss, CONFIG.SHEETS.MASTER_PC);
  var pcBySku = buildSkuToPcIndex_(ss); // see note below

  var pcByCode = {};
  masterPc.forEach(function (p) { pcByCode[p.pc_code] = p; });

  var rows = [];
  minmax.forEach(function (item) {
    if (item.category !== 'เบิกประจำ' || item.status !== 'ต้องสั่งซื้อ') return;
    var pcCode = pcBySku[item.sku] || '';
    var pc = pcByCode[pcCode] || {};
    var qty = Math.max(0, Number(item.max) - Number(item.balance));
    if (qty <= 0) return;

    rows.push([
      new Date(), item.sku, item.name, pcCode, pc.pc_name || '',
      pc.owner_email || '', pc.manager_email || '',
      item.balance, item.min, item.max, qty, item.unit_cost,
      round1_(qty * Number(item.unit_cost || 0)),
      'รอส่งอีเมล', '', '' // status, confirmed_qty, confirmed_at
    ]);
  });

  var sheet = ss.getSheetByName(CONFIG.SHEETS.REORDER_QUEUE) || ss.insertSheet(CONFIG.SHEETS.REORDER_QUEUE);
  var header = ['created_at', 'sku', 'name', 'pc_code', 'pc_name', 'owner_email',
    'manager_email', 'balance', 'min', 'max', 'suggested_qty', 'unit_cost',
    'suggested_value', 'status', 'confirmed_qty', 'confirmed_at'];
  // Append instead of clearing, so in-flight confirmations aren't lost.
  var existing = sheet.getLastRow();
  if (existing === 0) sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);

  logActivity_('buildReorderQueue', 'ok', rows.length + ' items queued');
  return rows;
}

/**
 * The balance export groups items under a PC header; master_item doesn't
 * necessarily carry pc_code. Rebuild sku -> pc_code from raw_balance each
 * run so Reorder.js stays correct even if items move between PCs.
 */
function buildSkuToPcIndex_(ss) {
  var raw = readSheetAsObjects_(ss, CONFIG.SHEETS.RAW_BALANCE);
  var idx = {};
  raw.forEach(function (r) { if (r.sku) idx[r.sku] = r.pc_code; });
  return idx;
}
