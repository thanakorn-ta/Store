/**
 * Requests.js
 * User -> Admin purchase requests. A user picks items + quantities in the
 * web UI, chooses which budget line (Company x Media Location x GL Code) and
 * which month the goods will be used, and submits. Admins are emailed and
 * approve/reject in the "คำขอ" tab. Approved requests are written to
 * pr_po_log so Budget.js sees them as committed spend.
 *
 * Budget comes from the "Budget STT 2026 - Revise-Budget" export, uploaded by
 * an admin in the Admin tab (parsed in the browser, aggregated per line/month,
 * stored in budget_master).
 *   plan      = Revise Budget when filled, else Budget (per source row, summed)
 *   reserved  = totals of requests still pending or approved for that line/month
 *   available = plan - actual - reserved
 */

var BUDGET_HEADER = ['key', 'company', 'division', 'media_type', 'media_group', 'media_location',
  'expense_group', 'gl_code', 'gl_name', 'year', 'month_number', 'budget', 'revise_budget', 'actual', 'plan'];
var REQUEST_HEADER = ['request_id', 'created_at', 'requester_email', 'requester_name', 'budget_key',
  'budget_label', 'budget_year', 'budget_month', 'total', 'item_count', 'available_at_submit', 'over_budget',
  'status', 'note', 'decided_by', 'decided_at', 'admin_note'];
var REQUEST_ITEM_HEADER = ['request_id', 'sku', 'name', 'pc_code', 'pc_name', 'unit', 'qty', 'unit_cost',
  'amount', 'note'];
var PR_LOG_HEADER = ['created_at', 'request_id', 'pc_code', 'budget_key', 'budget_year', 'month_number', 'amount', 'approved_by'];

var STATUS = { PENDING: 'รออนุมัติ', APPROVED: 'อนุมัติ', REJECTED: 'ไม่อนุมัติ', CANCELLED: 'ยกเลิก' };
var THAI_MONTH_NAMES = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

// ------------------------------------------------------------ budget

/** Admin uploads budget rows already aggregated by the page (see web/app.js parseBudgetWorkbook). */
function importBudget(rows) {
  var me = requireAdmin_();
  if (!rows || !rows.length) throw new Error('ไม่พบข้อมูล Budget ในไฟล์');
  if (rows.length > 20000) throw new Error('ข้อมูล Budget มากเกินไป');
  var values = rows.map(function (r) {
    return BUDGET_HEADER.map(function (h) { return r[h] == null ? '' : r[h]; });
  });
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = db_();
    var sheet = ensureSheet_(ss, CONFIG.SHEETS.BUDGET, BUDGET_HEADER);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, BUDGET_HEADER.length).setValues([BUDGET_HEADER]).setFontWeight('bold');
    sheet.getRange(2, 1, values.length, BUDGET_HEADER.length).setValues(values);
  } finally {
    lock.releaseLock();
  }
  logActivity_('importBudget', 'ok', me.email + ' imported ' + rows.length + ' budget rows');
  return getBudgetOptions();
}

function budgetLabel_(b) {
  return [b.media_location, b.gl_code + ' ' + b.gl_name, b.company].filter(String).join(' · ');
}

/** Budget lines with per-month plan/actual, plus amounts reserved by open requests. */
function getBudgetOptions() {
  requireActive_();
  var ss = db_();
  var rows = readSheetAsObjects_(ss, CONFIG.SHEETS.BUDGET);
  var lines = {};
  var year = 0;
  rows.forEach(function (b) {
    if (!b.key) return;
    var l = lines[b.key];
    if (!l) {
      l = lines[b.key] = {
        key: String(b.key), label: budgetLabel_(b), company: String(b.company),
        location: String(b.media_location), glCode: String(b.gl_code), glName: String(b.gl_name),
        expenseGroup: String(b.expense_group), months: {}
      };
    }
    var m = Number(b.month_number);
    if (!l.months[m]) l.months[m] = { plan: 0, actual: 0 };
    l.months[m].plan += Number(b.plan) || 0;
    l.months[m].actual += Number(b.actual) || 0;
    year = year || Number(b.year) || 0;
  });
  return {
    year: year,
    lines: Object.keys(lines).map(function (k) { return lines[k]; })
      .sort(function (a, b) { return a.label.localeCompare(b.label); }),
    reserved: reservedByBudget_(ss)
  };
}

/** { "key|month": amount } for requests that are pending or approved. */
function reservedByBudget_(ss) {
  var out = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS).forEach(function (r) {
    if (r.status !== STATUS.PENDING && r.status !== STATUS.APPROVED) return;
    var k = r.budget_key + '|' + Number(r.budget_month);
    out[k] = (out[k] || 0) + (Number(r.total) || 0);
  });
  return out;
}

// ------------------------------------------------------------ user: submit / list / cancel

/**
 * payload: { budgetKey, month, note, items: [{ sku, name, pc, pcName, unit, qty, unitCost, note }] }
 * Over-budget requests are accepted but flagged so the admin decides.
 */
function submitOrderRequest(payload) {
  var me = requireActive_();
  var items = (payload && payload.items || []).filter(function (i) { return Number(i.qty) > 0; });
  if (!items.length) throw new Error('ยังไม่มีรายการที่จำนวนมากกว่า 0');
  if (items.length > 500) throw new Error('รายการมากเกินไป (สูงสุด 500)');
  var month = Number(payload.month);
  if (!(month >= 1 && month <= 12)) throw new Error('กรุณาเลือกเดือนที่จะใช้ของ');

  var opts = getBudgetOptions();
  var line = opts.lines.filter(function (l) { return l.key === payload.budgetKey; })[0];
  if (!line) throw new Error('กรุณาเลือก Budget');

  var total = 0;
  var itemRows = items.map(function (i) {
    var qty = Math.round(Number(i.qty));
    var cost = Math.max(0, Number(i.unitCost) || 0);
    var amount = Math.round(qty * cost * 100) / 100;
    total += amount;
    return [null, String(i.sku), String(i.name || '').slice(0, 200), String(i.pc || ''), String(i.pcName || ''),
      String(i.unit || ''), qty, cost, amount, String(i.note || '').slice(0, 300)];
  });
  total = Math.round(total * 100) / 100;

  var mb = line.months[month] || { plan: 0, actual: 0 };
  var available = mb.plan - mb.actual - (opts.reserved[line.key + '|' + month] || 0);
  var over = total > available;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var id;
  try {
    var ss = db_();
    var reqSheet = ensureSheet_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER);
    var itemSheet = ensureSheet_(ss, CONFIG.SHEETS.REQUEST_ITEMS, REQUEST_ITEM_HEADER);
    id = 'REQ-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd') + '-' +
      ('000' + reqSheet.getLastRow()).slice(-4);
    reqSheet.appendRow([id, new Date(), me.email, me.name, line.key, line.label, opts.year, month, total,
      itemRows.length, Math.round(available * 100) / 100, over ? 'เกินงบ' : '', STATUS.PENDING,
      String(payload.note || '').slice(0, 500), '', '', '']);
    itemRows.forEach(function (r) { r[0] = id; });
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, REQUEST_ITEM_HEADER.length).setValues(itemRows);
  } finally {
    lock.releaseLock();
  }

  notifyAdminsNewRequest_(id, me, line, month, total, available, items);
  logActivity_('submitOrderRequest', 'ok', id + ' by ' + me.email + ' total ' + total);
  return { id: id, total: total, available: available, overBudget: over };
}

function listMyRequests() {
  var me = requireActive_();
  return loadRequests_(function (r) { return r.requester_email === me.email; });
}

function cancelMyRequest(id) {
  var me = requireActive_();
  return updateRequestStatus_(id, function (r) {
    if (r.requester_email !== me.email) throw new Error('ยกเลิกได้เฉพาะคำขอของตัวเอง');
    if (r.status !== STATUS.PENDING) throw new Error('ยกเลิกได้เฉพาะคำขอที่รออนุมัติ');
    return { status: STATUS.CANCELLED, decided_by: me.email, admin_note: 'ผู้ขอยกเลิกเอง' };
  });
}

// ------------------------------------------------------------ admin

function listAllRequests(status) {
  requireAdmin_();
  return loadRequests_(function (r) { return !status || r.status === status; });
}

/** decision: 'approve' | 'reject' */
function decideRequest(id, decision, note) {
  var me = requireAdmin_();
  var approve = decision === 'approve';
  var result = updateRequestStatus_(id, function (r) {
    if (r.status !== STATUS.PENDING) throw new Error('คำขอนี้ถูกดำเนินการไปแล้ว (' + r.status + ')');
    return { status: approve ? STATUS.APPROVED : STATUS.REJECTED, decided_by: me.email,
      admin_note: String(note || '').slice(0, 500) };
  });
  var req = result.request;

  if (approve) {
    // Committed spend per PC for Budget.js (checkBudget_) — one row per PC in the request
    var byPc = {};
    req.items.forEach(function (i) { byPc[i.pc_code] = (byPc[i.pc_code] || 0) + Number(i.amount); });
    var log = ensureSheet_(db_(), CONFIG.SHEETS.PR_LOG, PR_LOG_HEADER);
    Object.keys(byPc).forEach(function (pc) {
      log.appendRow([new Date(), req.request_id, pc, req.budget_key, req.budget_year, req.budget_month,
        Math.round(byPc[pc] * 100) / 100, me.email]);
    });
  }

  MailApp.sendEmail({
    to: req.requester_email,
    subject: '[Store Reorder] ' + req.request_id + ' ' + req.status,
    htmlBody: 'คำขอ ' + esc_(req.request_id) + ' (' + fmtMoney_(req.total) + ' บาท) : <b>' + esc_(req.status) + '</b>' +
      (req.admin_note ? '<br>หมายเหตุจาก Admin: ' + esc_(req.admin_note) : '') +
      '<br><a href="' + webAppUrl_() + '">เปิดระบบ</a>'
  });
  logActivity_('decideRequest', 'ok', id + ' ' + req.status + ' by ' + me.email);
  return listAllRequests();
}

// ------------------------------------------------------------ helpers

function loadRequests_(filterFn) {
  var ss = db_();
  var items = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS).forEach(function (i) {
    (items[i.request_id] = items[i.request_id] || []).push({
      sku: i.sku, name: i.name, pc_code: i.pc_code, pc_name: i.pc_name, unit: i.unit,
      qty: Number(i.qty), unit_cost: Number(i.unit_cost), amount: Number(i.amount), note: i.note
    });
  });
  return readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS)
    .filter(function (r) { return r.request_id && filterFn(r); })
    .map(function (r) { return toClientRequest_(r, items[r.request_id] || []); })
    .reverse()
    .slice(0, 300);
}

/** google.script.run can't return Date objects — convert to strings. */
function toClientRequest_(r, items) {
  var out = {};
  REQUEST_HEADER.forEach(function (h) {
    var v = r[h];
    out[h] = v instanceof Date ? Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : v;
  });
  out.month_label = THAI_MONTH_NAMES[Number(r.budget_month)] + ' ' + (Number(r.budget_year) ? Number(r.budget_year) : '');
  out.items = items;
  return out;
}

function updateRequestStatus_(id, mutate) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = db_();
    var sheet = ensureSheet_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER);
    var values = sheet.getDataRange().getValues();
    var col = {};
    values[0].forEach(function (h, i) { col[h] = i; });
    for (var r = 1; r < values.length; r++) {
      if (values[r][col.request_id] !== id) continue;
      var obj = {};
      REQUEST_HEADER.forEach(function (h) { obj[h] = values[r][col[h]]; });
      var change = mutate(obj);
      change.decided_at = new Date();
      Object.keys(change).forEach(function (k) {
        sheet.getRange(r + 1, col[k] + 1).setValue(change[k]);
        obj[k] = change[k];
      });
      var items = readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS).filter(function (i) { return i.request_id === id; });
      return { request: toClientRequest_(obj, items) };
    }
    throw new Error('ไม่พบคำขอ ' + id);
  } finally {
    lock.releaseLock();
  }
}

function notifyAdminsNewRequest_(id, me, line, month, total, available, items) {
  var admins = adminEmails_();
  if (!admins.length) return;
  var rows = items.map(function (i) {
    return '<tr><td>' + esc_(i.pc) + '</td><td>' + esc_(i.sku) + '</td><td>' + esc_(i.name) + '</td><td align="right">' +
      esc_(i.qty) + ' ' + esc_(i.unit) + '</td><td align="right">' + fmtMoney_(Number(i.qty) * (Number(i.unitCost) || 0)) + '</td></tr>';
  }).join('');
  MailApp.sendEmail({
    to: admins.join(','),
    subject: '[Store Reorder] คำขอสั่งซื้อใหม่ ' + id + ' จาก ' + (me.name || me.email) +
      (total > available ? ' (เกินงบ)' : ''),
    htmlBody:
      '<p><b>' + esc_(id) + '</b> โดย ' + esc_(me.name) + ' (' + esc_(me.email) + ')</p>' +
      '<p>Budget: ' + esc_(line.label) + '<br>ใช้เดือน: ' + THAI_MONTH_NAMES[month] +
      '<br>ยอดรวม: <b>' + fmtMoney_(total) + '</b> บาท · งบคงเหลือก่อนคำขอนี้: ' + fmtMoney_(available) + ' บาท' +
      (total > available ? ' <b style="color:#b23b2e">เกินงบ ' + fmtMoney_(total - available) + ' บาท</b>' : '') + '</p>' +
      '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse">' +
      '<tr><th>PC</th><th>รหัส</th><th>สินค้า</th><th>จำนวน</th><th>บาท</th></tr>' + rows + '</table>' +
      '<p><a href="' + webAppUrl_() + '">เปิดระบบเพื่ออนุมัติ</a></p>'
  });
}

function fmtMoney_(n) {
  return (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
