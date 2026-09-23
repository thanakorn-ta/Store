/**
 * Requests.js
 * Purchase requests — the approval flow the team uses by email today
 * ("ขออนุมัติสั่งซื้ออุปกรณ์ (ทีม Static)"):
 *
 *  1. User picks items + qty and, for EACH item, the budget line (Company x
 *     Media Location x GL Code) and the month it is charged to.
 *     If a line/month has no budget left the user must tick "ขอ Over Budget"
 *     and give a reason, otherwise the request is refused.       [รอ Admin ตรวจ]
 *  2. Admin reviews, then sends the approval email to the head (to/cc from
 *     Settings) in the team's format: per media + expense group, BG per month,
 *     items, Store balance, remarks, and the 14-column table.  [รอผู้บริหารอนุมัติ]
 *  3. Head replies "Approved" → admin records it (or the head, if a member,
 *     approves in the app).                                            [อนุมัติ]
 *     Admin can then send "ได้รับการอนุมัติสั่งซื้อ…" to Purchasing.
 *  4. issuePo(): PR/PO number, commits spend to pr_po_log         [ออก PR/PO แล้ว]
 *  5. markReceived()                                                 [รับของแล้ว]
 *
 * Admins can also send items to each PC's owners first (createRound →
 * [รอ PC ยืนยัน] → confirmRequest), which then enters step 2.
 *
 * Budget (budget_master, uploaded by an admin):
 *   plan      = Revise Budget when filled, else Budget
 *   reserved  = item amounts of requests waiting or approved, per item budget line/month
 *   available = plan - actual - reserved
 */

var BUDGET_HEADER = ['key', 'company', 'division', 'media_type', 'media_group', 'media_location',
  'expense_group', 'gl_code', 'gl_name', 'year', 'month_number', 'budget', 'revise_budget', 'actual', 'plan', 'remark'];
var REQUEST_HEADER = ['request_id', 'created_at', 'source', 'round_id', 'pc_code', 'pc_name', 'assigned_to',
  'requester_email', 'requester_name', 'budget_key', 'budget_label', 'budget_year', 'budget_month', 'budget_months', 'total',
  'item_count', 'available_at_submit', 'over_budget', 'over_reason', 'status', 'note', 'sent_at', 'reminded_at',
  'confirmed_by', 'confirmed_at', 'manager_email', 'manager_decision', 'manager_note', 'manager_at',
  'decided_by', 'decided_at', 'admin_note', 'approval_to', 'approval_cc', 'approval_subject', 'approval_sent_at',
  'approved_by', 'approved_at', 'approval_note', 'purchasing_sent_at', 'po_no', 'po_at', 'received_at'];
var REQUEST_ITEM_HEADER = ['request_id', 'sku', 'name', 'pc_code', 'pc_name', 'unit', 'suggested_qty', 'qty',
  'unit_cost', 'amount', 'note', 'budget_key', 'budget_label', 'budget_year', 'budget_month', 'store_balance',
  'over_budget', 'over_reason', 'pr_no', 'po_no', 'po_at', 'eta_date', 'received_at'];
var PR_LOG_HEADER = ['created_at', 'request_id', 'po_no', 'pc_code', 'budget_key', 'budget_year', 'month_number', 'amount', 'issued_by'];

var STATUS = {
  WAIT_PC: 'รอ PC ยืนยัน', PC_REJECTED: 'PC ปฏิเสธ',
  PENDING: 'รอ Admin ตรวจ', APPROVAL_WAIT: 'รอผู้บริหารอนุมัติ', OVER_WAIT: 'รอหัวหน้าอนุมัติเกินงบ',
  APPROVED: 'อนุมัติ', REJECTED: 'ไม่อนุมัติ', CANCELLED: 'ยกเลิก',
  PO: 'ออก PR/PO แล้ว', RECEIVED: 'รับของแล้ว',
  EDIT: 'รอผู้ขอแก้ไข'        // called back by the user or sent back by an admin; holds no budget
};
var LEGACY_PENDING = 'รออนุมัติ'; // status text before version 2026-09-22c; treated as PENDING
var RESERVING = [STATUS.PENDING, STATUS.APPROVAL_WAIT, STATUS.OVER_WAIT, STATUS.APPROVED, STATUS.PO, STATUS.RECEIVED];
var WAITING = [STATUS.PENDING, STATUS.APPROVAL_WAIT, STATUS.OVER_WAIT];

function normStatus_(s) { return s === LEGACY_PENDING ? STATUS.PENDING : String(s || ''); }
var THAI_MONTH_NAMES = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

// ============================================================ budget

/** Admin uploads budget rows already aggregated by the page (web/parsers.js parseBudget). */
function importBudget(rows) {
  var me = requireAdmin_();
  if (!rows || !rows.length) throw new Error('ไม่พบข้อมูล Budget ในไฟล์');
  if (rows.length > 20000) throw new Error('ข้อมูล Budget มากเกินไป');
  writeBudgetMaster_(rows);
  logActivity_('importBudget', 'ok', me.email + ' imported ' + rows.length + ' budget rows');
  return getBudgetOptions();
}

/** Replaces budget_master with the given aggregated rows (see BUDGET_HEADER). */
function writeBudgetMaster_(rows) {
  var values = rows.map(function (r) {
    return BUDGET_HEADER.map(function (h) { return r[h] == null ? '' : r[h]; });
  });
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = ensureSheet_(db_(), CONFIG.SHEETS.BUDGET, BUDGET_HEADER);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, BUDGET_HEADER.length).setValues([BUDGET_HEADER]).setFontWeight('bold');
    if (values.length) sheet.getRange(2, 1, values.length, BUDGET_HEADER.length).setValues(values);
  } finally {
    lock.releaseLock();
  }
}

function budgetLabel_(b) {
  return [b.media_location, b.gl_code + ' ' + b.gl_name, b.company].filter(String).join(' · ');
}

/** Budget lines with per-month plan/actual, plus amounts reserved by open requests. */
function getBudgetOptions() {
  requireActive_();
  return budgetOptions_(db_());
}

function budgetOptions_(ss) {
  var lines = {};
  var year = 0;
  readSheetAsObjects_(ss, CONFIG.SHEETS.BUDGET).forEach(function (b) {
    if (!b.key) return;
    var l = lines[b.key];
    if (!l) {
      l = lines[b.key] = {
        key: String(b.key), label: budgetLabel_(b), company: String(b.company),
        location: String(b.media_location), mediaGroup: String(b.media_group || ''), mediaType: String(b.media_type || ''),
        glCode: String(b.gl_code), glName: String(b.gl_name), expenseGroup: String(b.expense_group), months: {}
      };
    }
    var m = Number(b.month_number);
    if (!l.months[m]) l.months[m] = { plan: 0, actual: 0, remark: '' };
    l.months[m].plan += Number(b.plan) || 0;
    l.months[m].actual += Number(b.actual) || 0;
    if (b.remark) l.months[m].remark = String(b.remark);
    year = year || Number(b.year) || 0;
  });
  return {
    year: year,
    lines: Object.keys(lines).map(function (k) { return lines[k]; })
      .sort(function (a, b) { return a.label.localeCompare(b.label); }),
    reserved: reservedByBudget_(ss),
    reservedDetail: reservedDetail_(ss)
  };
}

/**
 * Budget held by open requests, per item: [{ key, month, amount, status, requestId }].
 * Each item carries its own budget line/month; items saved before that fall back
 * to the request header's single budget.
 */
function reservations_(ss) {
  var reqs = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS).forEach(function (r) {
    if (r.request_id) reqs[r.request_id] = { status: normStatus_(r.status), key: r.budget_key, month: Number(r.budget_month) };
  });
  var out = [];
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS).forEach(function (i) {
    var r = reqs[i.request_id];
    if (!r || RESERVING.indexOf(r.status) === -1 || !(Number(i.qty) > 0)) return;
    var key = i.budget_key || r.key, month = Number(i.budget_month) || r.month;
    if (!key || !month) return;
    out.push({ key: String(key), month: month, amount: Number(i.amount) || 0, status: r.status, requestId: i.request_id });
  });
  return out;
}

/**
 * { "key|month": { pending, committed } } for the budget page:
 *   pending   = waiting (รอ Admin ตรวจ, รอผู้บริหารอนุมัติ, รอหัวหน้าอนุมัติเกินงบ)
 *   committed = approved onward (อนุมัติ, ออก PR/PO แล้ว, รับของแล้ว)
 */
function reservedDetail_(ss) {
  var out = {};
  reservations_(ss).forEach(function (x) {
    var o = out[x.key + '|' + x.month] = out[x.key + '|' + x.month] || { pending: 0, committed: 0 };
    o[WAITING.indexOf(x.status) !== -1 ? 'pending' : 'committed'] += x.amount;
  });
  return out;
}

/** { "key|month": amount } held by open requests. excludeId skips one request. */
function reservedByBudget_(ss, excludeId) {
  var out = {};
  reservations_(ss).forEach(function (x) {
    if (x.requestId === excludeId) return;
    out[x.key + '|' + x.month] = (out[x.key + '|' + x.month] || 0) + x.amount;
  });
  return out;
}

/**
 * Checks each budget line/month used by the given item rows against what is left.
 * Lines that go over need a reason in overReasons["key|month"], else this throws.
 * Returns [{ key, month, label, plan, available, amount, over, reason }].
 */
function checkBudgetLines_(opts, rows, overReasons) {
  var groups = {}, order = [];
  rows.forEach(function (r) {
    var k = r.budget_key + '|' + r.budget_month;
    if (!groups[k]) { groups[k] = { key: r.budget_key, month: Number(r.budget_month), amount: 0 }; order.push(k); }
    groups[k].amount += Number(r.amount) || 0;
  });
  var missing = [];
  var out = order.map(function (k) {
    var g = groups[k];
    var av = availableFor_(opts, g.key, g.month);
    if (!av) throw new Error('ไม่พบ Budget ที่เลือก (' + g.key + ')');
    var mb = av.line.months[g.month] || { plan: 0 };
    var over = round2_(g.amount) > round2_(av.available);
    var reason = String((overReasons || {})[k] || '').trim().slice(0, 500);
    if (over && !reason) missing.push(av.line.label + ' เดือน ' + THAI_MONTH_NAMES[g.month] + ' (คงเหลือ ' + fmtMoney_(av.available) + ' บาท)');
    return { key: g.key, month: g.month, label: av.line.label, plan: mb.plan, available: round2_(av.available),
      amount: round2_(g.amount), over: over, reason: over ? reason : '' };
  });
  if (missing.length) throw new Error('งบไม่พอ ต้องขอ Over Budget พร้อมเหตุผล: ' + missing.join(' · '));
  return out;
}

/** Header fields summarising the budget lines of a request. */
function budgetSummaryFields_(lines, year) {
  var keys = uniq_(lines.map(function (l) { return l.key; }));
  var months = uniq_(lines.map(function (l) { return String(l.month); })).map(Number).sort(function (a, b) { return a - b; });
  var over = lines.filter(function (l) { return l.over; });
  return {
    budget_key: keys.length === 1 ? keys[0] : '',
    budget_label: keys.length === 1 ? lines[0].label : (keys.length + ' Budget: ' + uniq_(lines.map(function (l) { return l.label; })).join(' / ')).slice(0, 500),
    budget_year: year, budget_month: months.length === 1 ? months[0] : '',
    budget_months: months.map(function (m) { return THAI_MONTH_NAMES[m]; }).join(', '),
    available_at_submit: round2_(lines.reduce(function (a, l) { return a + l.available; }, 0)),
    over_budget: over.length ? 'Over Budget' : '',
    over_reason: over.map(function (l) { return l.label + ' ' + THAI_MONTH_NAMES[l.month] + ': ' + l.reason; }).join(' | ').slice(0, 1000)
  };
}

function availableFor_(opts, key, month) {
  var line = opts.lines.filter(function (l) { return l.key === key; })[0];
  if (!line) return null;
  var mb = line.months[month] || { plan: 0, actual: 0 };
  return { line: line, available: mb.plan - mb.actual - (opts.reserved[key + '|' + month] || 0) };
}

// ============================================================ (5) admin creates a round per PC

/**
 * items: [{ sku, name, pc, pcName, unit, qty, unitCost }] — usually the admin's cart.
 * Creates one request per PC in [รอ PC ยืนยัน] and emails that PC's owners.
 */
function createRound(items, note) {
  var me = requireAdmin_();
  items = (items || []).filter(function (i) { return Number(i.qty) > 0 && i.pc; });
  if (!items.length) throw new Error('ไม่มีรายการ (ต้องมีจำนวน > 0 และรู้ PC)');
  var ss = db_();
  var master = readMasterPc_(ss);
  var byPc = {};
  items.forEach(function (i) { (byPc[i.pc] = byPc[i.pc] || []).push(i); });

  var roundId = 'RND-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd-HHmm');
  var created = [], noOwner = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Object.keys(byPc).sort().forEach(function (pc) {
      var list = byPc[pc];
      var m = master[pc] || { owner_emails: [], manager_email: '', pc_name: list[0].pcName || '' };
      var id = nextRequestId_(ss);
      var rows = itemRows_(id, list, true);
      appendObject_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER, {
        request_id: id, created_at: new Date(), source: 'round', round_id: roundId,
        pc_code: pc, pc_name: m.pc_name || list[0].pcName || '', assigned_to: m.owner_emails.join(', '),
        requester_email: me.email, requester_name: me.name, total: sumAmount_(rows), item_count: rows.length,
        status: STATUS.WAIT_PC, note: String(note || '').slice(0, 500),
        sent_at: m.owner_emails.length ? new Date() : '', manager_email: m.manager_email
      });
      appendObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS, REQUEST_ITEM_HEADER, rows);
      created.push({ id: id, pc: pc, owners: m.owner_emails, items: rows });
      if (!m.owner_emails.length) noOwner.push(pc);
    });
  } finally {
    lock.releaseLock();
  }

  created.forEach(function (c) {
    if (c.owners.length) sendPcConfirmEmail_(c.id, c.pc, (master[c.pc] || {}).pc_name || '', c.owners, c.items, note, false);
  });
  logActivity_('createRound', 'ok', roundId + ' by ' + me.email + ': ' + created.length + ' PC, no owner: ' + noOwner.join(','));
  return { roundId: roundId, created: created.map(function (c) { return { id: c.id, pc: c.pc, sentTo: c.owners.join(', ') }; }), noOwner: noOwner };
}

// ============================================================ (1) user: submit from cart → รอ Admin ตรวจ

/**
 * payload: { note, overReasons: { "budgetKey|month": reason },
 *            items: [{ sku, name, pc, pcName, unit, qty, unitCost, note, budgetKey, month, storeBalance }],
 *            files: [{ name, type, data }]  — quotations (optional, Quotes.js)
 *            editId, keepFileIds            — resubmitting a request the user called back (รอผู้ขอแก้ไข) }
 * Every item carries its own budget line + month. A line/month without enough budget
 * left needs an Over Budget reason, else the request is refused.
 */
function submitOrderRequest(payload) {
  var me = requireActive_();
  payload = payload || {};
  var items = orderItems_(payload);
  checkUploads_(payload.files);
  var ss = db_();
  var editId = String(payload.editId || '');
  var prev = null;
  if (editId) {
    prev = findRequest_(ss, editId);
    if (prev.requester_email !== me.email && me.role !== 'admin') throw new Error('แก้ไขได้เฉพาะคำขอของตัวเอง');
    if (prev.status !== STATUS.EDIT) throw new Error('คำขอ ' + editId + ' ไม่ได้อยู่ในสถานะแก้ไข (' + prev.status + ') — กด "เรียกกลับมาแก้ไข" ก่อน');
  }

  var id, total, rows, lines, summary;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var opts = budgetOptions_(ss); // a request in รอผู้ขอแก้ไข holds no budget, so its old amounts don't count
    id = editId || nextRequestId_(ss);
    rows = itemRows_(id, items, false, opts);
    lines = checkBudgetLines_(opts, rows, payload.overReasons);
    markOverRows_(rows, lines);
    summary = budgetSummaryFields_(lines, opts.year);
    total = sumAmount_(rows);
    var pcs = uniq_(items.map(function (i) { return i.pc; }));
    var master = readMasterPc_(ss);
    var head = {
      pc_code: pcs.join(', '), total: total, item_count: rows.length,
      status: STATUS.PENDING, note: String(payload.note || '').slice(0, 500),
      confirmed_by: me.email, confirmed_at: new Date(),
      manager_email: pcs.length === 1 && master[pcs[0]] ? master[pcs[0]].manager_email : ''
    };
    Object.keys(summary).forEach(function (k) { head[k] = summary[k]; });
    if (prev) {
      head.admin_note = prev.admin_note ? 'แก้ไขแล้วส่งใหม่ (เดิม: ' + String(prev.admin_note).slice(0, 300) + ')' : 'แก้ไขแล้วส่งใหม่';
      head.decided_by = ''; head.decided_at = '';
      setRequestFields_(ss, id, head);
      replaceItems_(ss, id, rows);
    } else {
      head.request_id = id; head.created_at = new Date(); head.source = 'user';
      head.requester_email = me.email; head.requester_name = me.name;
      appendObject_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER, head);
      appendObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS, REQUEST_ITEM_HEADER, rows);
    }
    if (prev) pruneQuoteFiles_(ss, id, payload.keepFileIds);
    saveQuoteFiles_(ss, id, payload.files, me);
  } finally {
    lock.releaseLock();
  }
  notifyAdminsForReview_(id, !!prev);
  logActivity_(prev ? 'resubmitOrderRequest' : 'submitOrderRequest', 'ok', id + ' by ' + me.email + ' total ' + total + (summary.over_budget ? ' OVER' : ''));
  return { id: id, total: total, overBudget: !!summary.over_budget, lines: lines, resubmitted: !!prev };
}

/** Validated items with qty > 0; each needs a budget line and a month. */
function orderItems_(payload) {
  var items = (payload.items || []).filter(function (i) { return Number(i.qty) > 0; });
  if (!items.length) throw new Error('ยังไม่มีรายการที่จำนวนมากกว่า 0');
  if (items.length > 500) throw new Error('รายการมากเกินไป (สูงสุด 500)');
  items.forEach(function (i) {
    // older pages send one budget for the whole cart
    if (!i.budgetKey) i.budgetKey = payload.budgetKey;
    if (!i.month) i.month = payload.month;
    var m = Number(i.month);
    if (!String(i.name || i.sku || '').trim()) throw new Error('รายการที่ไม่มีชื่อสินค้า');
    if (!i.budgetKey) throw new Error('กรุณาเลือก Budget ของรายการ ' + (i.name || i.sku));
    if (!(m >= 1 && m <= 12)) throw new Error('กรุณาเลือกเดือนที่ใช้ของ ของรายการ ' + (i.name || i.sku));
    if (!i.sku) i.sku = 'NEW-' + String(i.name).slice(0, 30);
  });
  return items;
}

/** Replaces a request's item rows (callers hold the script lock). */
function replaceItems_(ss, id, rows) {
  rewriteSheetRows_(ss, CONFIG.SHEETS.REQUEST_ITEMS, REQUEST_ITEM_HEADER,
    function (row, col) { return row[col.request_id] !== id; }, rows);
}

// ============================================================ user calls a request back to edit it

/**
 * รอ Admin ตรวจ / ไม่อนุมัติ / รอผู้บริหารอนุมัติ(ไม่ได้) → รอผู้ขอแก้ไข.
 * Returns the request so the page can load it into the order form; resubmit with payload.editId.
 */
function recallRequest(id) {
  var me = requireActive_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.requester_email !== me.email && me.role !== 'admin') throw new Error('เรียกกลับได้เฉพาะคำขอของตัวเอง');
  if ([STATUS.PENDING, STATUS.EDIT, STATUS.REJECTED].indexOf(req.status) === -1) {
    throw new Error('เรียกกลับมาแก้ไขได้เฉพาะคำขอที่รอ Admin ตรวจ / ถูกส่งกลับ / ไม่อนุมัติ (ตอนนี้: ' + req.status + ')');
  }
  if (req.status !== STATUS.EDIT) {
    setRequestFields_(ss, id, { status: STATUS.EDIT,
      admin_note: req.status === STATUS.REJECTED ? 'ไม่อนุมัติ: ' + String(req.admin_note || '') : 'ผู้ขอเรียกกลับมาแก้ไข' });
    if (req.status === STATUS.PENDING) {
      mailAdmins_('[Store Reorder] ' + id + ' ผู้ขอเรียกกลับไปแก้ไข',
        esc_(me.name || me.email) + ' เรียกคำขอ ' + esc_(id) + ' กลับไปแก้ไข — จะส่งกลับมาให้ตรวจอีกครั้ง' + appLink_(id));
    }
    logActivity_('recallRequest', 'ok', id + ' by ' + me.email);
  }
  var files = quoteFilesByRequest_(ss);
  return toClientRequest_(findRequest_(ss, id), readItems_(ss, id), files[id] || []);
}

// ============================================================ admin: fix / send back / delete

/** Admin sends a request back to the requester to fix (it stops holding budget). */
function returnForEdit(id, note) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if ([STATUS.PENDING, STATUS.APPROVAL_WAIT].indexOf(req.status) === -1) throw new Error('ส่งกลับได้เฉพาะคำขอที่รอตรวจ/รออนุมัติ (' + req.status + ')');
  note = String(note || '').trim().slice(0, 500);
  if (!note) throw new Error('กรุณาบอกสิ่งที่ต้องแก้');
  setRequestFields_(ss, id, { status: STATUS.EDIT, admin_note: note, decided_by: me.email, decided_at: new Date() });
  mailPeople_([req.requester_email], '[Store Reorder] ' + id + ' ส่งกลับให้แก้ไข',
    'Admin ส่งคำขอ ' + esc_(id) + ' กลับให้แก้ไข<br>สิ่งที่ต้องแก้: <b>' + esc_(note) + '</b>' +
    '<br>เปิดระบบ → คำขอของฉัน → "แก้ไขแล้วส่งใหม่"' + appLink_(id));
  logActivity_('returnForEdit', 'ok', id + ' by ' + me.email);
  return loadRequestsFor_(me);
}

/**
 * Admin corrects a request (qty, price, budget line/month, note, remove lines).
 * payload: { items: [same shape as submitOrderRequest], overReasons, note }
 */
function adminUpdateRequest(id, payload) {
  var me = requireAdmin_();
  payload = payload || {};
  var ss = db_();
  var req = findRequest_(ss, id);
  if ([STATUS.PENDING, STATUS.APPROVAL_WAIT, STATUS.EDIT].indexOf(req.status) === -1) {
    throw new Error('แก้ไขได้เฉพาะคำขอที่ยังไม่อนุมัติ (' + req.status + ')');
  }
  var items = orderItems_(payload);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var total, summary;
  try {
    var base = budgetOptions_(ss);
    var opts = { year: base.year, lines: base.lines, reserved: reservedByBudget_(ss, id) };
    var rows = itemRows_(id, items, false, opts);
    var lines = checkBudgetLines_(opts, rows, payload.overReasons);
    markOverRows_(rows, lines);
    summary = budgetSummaryFields_(lines, opts.year);
    total = sumAmount_(rows);
    var fields = { total: total, item_count: rows.length, pc_code: uniq_(items.map(function (i) { return i.pc; })).join(', '),
      admin_note: ('แก้ไขโดย Admin ' + (me.name || me.email) + (payload.adminNote ? ': ' + payload.adminNote : '')).slice(0, 500) };
    if (payload.note != null) fields.note = String(payload.note).slice(0, 500);
    Object.keys(summary).forEach(function (k) { fields[k] = summary[k]; });
    setRequestFields_(ss, id, fields);
    replaceItems_(ss, id, rows);
  } finally {
    lock.releaseLock();
  }
  mailPeople_([req.requester_email], '[Store Reorder] ' + id + ' Admin แก้ไขข้อมูลคำขอ',
    'Admin ' + esc_(me.name || me.email) + ' แก้ไขคำขอ ' + esc_(id) + ' — ยอดรวมใหม่ ' + fmtMoney_(total) + ' บาท' +
    (payload.adminNote ? '<br>หมายเหตุ: ' + esc_(payload.adminNote) : '') + itemsTable_(readItems_(ss, id)) + appLink_(id));
  logActivity_('adminUpdateRequest', 'ok', id + ' by ' + me.email + ' total ' + total);
  return loadRequestsFor_(me);
}

/** Admin deletes a request, its items and its attachment list (files go to Drive trash). */
function deleteRequest(id) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if ([STATUS.PO, STATUS.RECEIVED].indexOf(req.status) !== -1) throw new Error('ลบไม่ได้: ออก PR/PO แล้ว (ตัดงบไปแล้ว)');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    pruneQuoteFiles_(ss, id, []);
    replaceItems_(ss, id, []);
    rewriteSheetRows_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER, function (row, col) { return row[col.request_id] !== id; });
  } finally {
    lock.releaseLock();
  }
  logActivity_('deleteRequest', 'ok', id + ' (' + req.status + ', ' + req.total + ') by ' + me.email);
  return loadRequestsFor_(me);
}

/** Copies the Over Budget flag/reason of each budget line onto its item rows. */
function markOverRows_(rows, lines) {
  var by = {};
  lines.forEach(function (l) { by[l.key + '|' + l.month] = l; });
  rows.forEach(function (r) {
    var l = by[r.budget_key + '|' + r.budget_month];
    r.over_budget = l && l.over ? 'Over Budget' : '';
    r.over_reason = l && l.over ? l.reason : '';
  });
}

// ============================================================ (6) PC owner confirms / rejects

/**
 * payload: { items: [{ sku, qty, note }], budgetKey, month, note, overReason }
 * qty 0 drops the line. The chosen budget/month applies to every item of this PC.
 */
function confirmRequest(id, payload) {
  var me = requireActive_();
  payload = payload || {};
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.WAIT_PC) throw new Error('คำขอนี้ไม่ได้รอการยืนยันแล้ว (' + req.status + ')');
  if (!canActAsPc_(me, req)) throw new Error('คุณไม่ได้เป็นผู้รับผิดชอบ ' + req.pc_code);
  var month = Number(payload.month);
  if (!(month >= 1 && month <= 12)) throw new Error('กรุณาเลือกเดือนที่จะใช้ของ');
  var opts = budgetOptions_(ss);
  var av = availableFor_(opts, payload.budgetKey, month);
  if (!av) throw new Error('กรุณาเลือก Budget');

  var edits = {};
  (payload.items || []).forEach(function (i) { edits[i.sku] = i; });
  var planned = readItems_(ss, id).map(function (i) {
    var qty = edits[i.sku] ? Math.max(0, Math.round(Number(edits[i.sku].qty) || 0)) : i.qty;
    return { budget_key: av.line.key, budget_month: month, amount: round2_(qty * (Number(i.unit_cost) || 0)), qty: qty };
  }).filter(function (i) { return i.qty > 0; });
  if (!planned.length) throw new Error('ทุกรายการเป็น 0 — ถ้าไม่ต้องการสั่ง ให้กด "ไม่สั่งรอบนี้"');
  var reasons = {};
  reasons[av.line.key + '|' + month] = payload.overReason;
  var lines = checkBudgetLines_(opts, planned, reasons);
  var line = lines[0];

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var total;
  try {
    total = updateItems_(ss, id, function (item) {
      var e = edits[item.sku];
      if (e) {
        item.qty = Math.max(0, Math.round(Number(e.qty) || 0));
        item.amount = round2_(item.qty * (Number(item.unit_cost) || 0));
        if (e.note != null) item.note = String(e.note).slice(0, 300);
      }
      item.budget_key = av.line.key; item.budget_label = av.line.label;
      item.budget_year = opts.year; item.budget_month = month;
      item.over_budget = line.over ? 'Over Budget' : ''; item.over_reason = line.over ? line.reason : '';
      return item;
    });
    var fields = budgetSummaryFields_(lines, opts.year);
    fields.total = total.amount; fields.item_count = total.count; fields.status = STATUS.PENDING;
    fields.note = String(payload.note || req.note || '').slice(0, 500);
    fields.confirmed_by = me.email; fields.confirmed_at = new Date();
    setRequestFields_(ss, id, fields);
  } finally {
    lock.releaseLock();
  }
  notifyAdminsForReview_(id);
  logActivity_('confirmRequest', 'ok', id + ' by ' + me.email + ' total ' + total.amount);
  return loadRequestsFor_(me);
}

function rejectByPc(id, note) {
  var me = requireActive_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.WAIT_PC) throw new Error('คำขอนี้ไม่ได้รอการยืนยันแล้ว');
  if (!canActAsPc_(me, req)) throw new Error('คุณไม่ได้เป็นผู้รับผิดชอบ ' + req.pc_code);
  setRequestFields_(ss, id, { status: STATUS.PC_REJECTED, confirmed_by: me.email, confirmed_at: new Date(),
    note: String(note || '').slice(0, 500) });
  mailAdmins_('[Store Reorder] ' + id + ' ' + req.pc_code + ' ไม่สั่งรอบนี้',
    esc_(me.name || me.email) + ' แจ้งว่า ' + esc_(req.pc_code) + ' ไม่สั่งรอบนี้' + (note ? '<br>เหตุผล: ' + esc_(note) : '') + appLink_(id));
  logActivity_('rejectByPc', 'ok', id + ' by ' + me.email);
  return loadRequestsFor_(me);
}

function cancelMyRequest(id) {
  var me = requireActive_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.requester_email !== me.email && me.role !== 'admin') throw new Error('ยกเลิกได้เฉพาะคำขอของตัวเอง');
  var cancellable = [STATUS.PENDING, STATUS.WAIT_PC, STATUS.OVER_WAIT, STATUS.EDIT].concat(me.role === 'admin' ? [STATUS.APPROVAL_WAIT] : []);
  if (cancellable.indexOf(req.status) === -1) throw new Error('ยกเลิกไม่ได้ในสถานะ ' + req.status);
  setRequestFields_(ss, id, { status: STATUS.CANCELLED, decided_by: me.email, decided_at: new Date(), admin_note: 'ยกเลิกโดย ' + me.email });
  logActivity_('cancelRequest', 'ok', id + ' by ' + me.email);
  return loadRequestsFor_(me);
}

// ============================================================ (7)(8) admin review + budget check

/**
 * Admin rejects a request (while reviewing, or when the head said no).
 * Approval goes through sendApprovalEmail → recordApproval (Approval.js).
 */
function decideRequest(id, decision, note) {
  var me = requireAdmin_();
  if (decision === 'approve') throw new Error('อนุมัติผ่าน "ส่งอีเมลขออนุมัติ" แล้วบันทึกผลเมื่อผู้บริหารตอบกลับ');
  var ss = db_();
  var req = findRequest_(ss, id);
  if ([STATUS.PENDING, STATUS.APPROVAL_WAIT].indexOf(req.status) === -1) throw new Error('ไม่อนุมัติได้เฉพาะคำขอที่รอตรวจ/รออนุมัติ (' + req.status + ')');
  setRequestFields_(ss, id, { status: STATUS.REJECTED, decided_by: me.email,
    decided_at: new Date(), admin_note: String(note || '').slice(0, 500) });
  mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ' + STATUS.REJECTED,
    'คำขอ ' + esc_(id) + ' (' + fmtMoney_(req.total) + ' บาท): <b>' + STATUS.REJECTED + '</b>' +
    (note ? '<br>หมายเหตุจาก Admin: ' + esc_(note) : '') + appLink_(id));
  logActivity_('decideRequest', 'ok', id + ' reject by ' + me.email);
  return loadRequestsFor_(me);
}

/** (8b) AI-drafted over-budget approval request to the PC manager. */
function requestOverBudget(id, managerEmail, reason) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.PENDING) throw new Error('ขออนุมัติเกินงบได้เฉพาะคำขอที่รออนุมัติ');
  var manager = String(managerEmail || req.manager_email || '').trim().toLowerCase();
  if (!/@/.test(manager)) throw new Error('ยังไม่มีอีเมลหัวหน้า PC — ใส่ใน Master PC หรือกรอกตอนกดขออนุมัติ');
  var av = availableFor_({ lines: budgetOptions_(ss).lines, reserved: reservedByBudget_(ss, id) }, req.budget_key, Number(req.budget_month));
  var shortfall = round2_(Number(req.total) - (av ? av.available : 0));
  setRequestFields_(ss, id, { status: STATUS.OVER_WAIT, manager_email: manager, manager_decision: '', admin_note: String(reason || '').slice(0, 500) });

  var items = readItems_(ss, id);
  var fallback = 'เรียน หัวหน้า ' + esc_(req.pc_code) + '<br><br>คำขอสั่งซื้อ ' + esc_(id) + ' ยอด ' + fmtMoney_(req.total) +
    ' บาท เกินงบ ' + esc_(req.budget_label) + ' เดือน ' + THAI_MONTH_NAMES[Number(req.budget_month)] + ' อยู่ ' + fmtMoney_(shortfall) +
    ' บาท' + (reason ? '<br>เหตุผล: ' + esc_(reason) : '') + '<br>กรุณาพิจารณาอนุมัติ';
  var body = aiDraft_(function () {
    return draftBudgetExceptionRequest_(req.pc_code + ' ' + req.pc_name, items.map(function (i) { return { name: i.name, confirmed_qty: i.qty }; }), shortfall)
      .replace(/\n/g, '<br>');
  }, fallback);
  mailPeople_([manager], '[Store Reorder] ขออนุมัติเกินงบ ' + id + ' (' + fmtMoney_(shortfall) + ' บาท)',
    body + itemsTable_(items) + appLink_(id, 'เปิดระบบเพื่ออนุมัติ/ไม่อนุมัติ'));
  logActivity_('requestOverBudget', 'ok', id + ' -> ' + manager + ' shortfall ' + shortfall);
  return loadRequestsFor_(me);
}

/** Manager (or admin on their behalf) answers the over-budget request. */
function managerDecision(id, approve, note) {
  var me = requireActive_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.OVER_WAIT) throw new Error('คำขอนี้ไม่ได้รออนุมัติเกินงบ');
  if (me.email !== String(req.manager_email).toLowerCase() && me.role !== 'admin') throw new Error('เฉพาะหัวหน้าที่ได้รับคำขอเท่านั้น');
  setRequestFields_(ss, id, {
    status: approve ? STATUS.PENDING : STATUS.REJECTED, manager_decision: approve ? 'อนุมัติ' : 'ไม่อนุมัติ',
    manager_note: String(note || '').slice(0, 500), manager_at: new Date()
  });
  mailAdmins_('[Store Reorder] หัวหน้า' + (approve ? 'อนุมัติ' : 'ไม่อนุมัติ') + 'เกินงบ ' + id,
    esc_(me.name || me.email) + ' ' + (approve ? 'อนุมัติ' : 'ไม่อนุมัติ') + 'เกินงบคำขอ ' + esc_(id) +
    (note ? '<br>หมายเหตุ: ' + esc_(note) : '') + (approve ? '<br>ออก PR/PO ต่อได้' : '') + appLink_(id));
  logActivity_('managerDecision', 'ok', id + ' ' + (approve ? 'approve' : 'reject') + ' by ' + me.email);
  return loadRequestsFor_(me);
}

// ============================================================ (9)(10) PR/PO + receive

function issuePo(id, poNo) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.APPROVED) throw new Error('ออก PR/PO ได้เฉพาะคำขอที่อนุมัติแล้ว');
  poNo = String(poNo || '').trim().slice(0, 60);
  if (!poNo) throw new Error('กรุณาใส่เลข PR/PO');
  setRequestFields_(ss, id, { status: STATUS.PO, po_no: poNo, po_at: new Date() });
  // Commit spend per PC x budget line x month (Budget.js / checkBudget_ reads pr_po_log)
  // Items that already carry their own PO number were committed by savePrPo (Delivery.js).
  var groups = {};
  readItems_(ss, id).forEach(function (i) {
    if (!(i.qty > 0) || String(i.po_no || '').trim()) return;
    var key = i.budget_key || req.budget_key, m = i.budget_month || Number(req.budget_month);
    var g = groups[i.pc_code + '|' + key + '|' + m] = groups[i.pc_code + '|' + key + '|' + m] ||
      { pc: i.pc_code, key: key, month: m, year: i.budget_year || req.budget_year, amount: 0 };
    g.amount += Number(i.amount) || 0;
  });
  appendObjects_(ss, CONFIG.SHEETS.PR_LOG, PR_LOG_HEADER, Object.keys(groups).map(function (k) {
    var g = groups[k];
    return { created_at: new Date(), request_id: id, po_no: poNo, pc_code: g.pc, budget_key: g.key,
      budget_year: g.year, month_number: g.month, amount: round2_(g.amount), issued_by: me.email };
  }));
  var stamp = new Date();
  updateItems_(ss, id, function (item) { // เลขเดียวกันทั้งคำขอ — แยกรายชิ้นได้ที่ Delivery.js
    if (Number(item.qty) > 0 && !String(item.po_no || '').trim()) { item.po_no = poNo; item.po_at = stamp; }
    return item;
  });
  mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ออก PR/PO แล้ว: ' + poNo,
    'คำขอ ' + esc_(id) + ' ออก PR/PO เลขที่ <b>' + esc_(poNo) + '</b> แล้ว' + appLink_(id));
  logActivity_('issuePo', 'ok', id + ' ' + poNo + ' by ' + me.email);
  return loadRequestsFor_(me);
}

function markReceived(id, note) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.PO) throw new Error('บันทึกรับของได้เฉพาะคำขอที่ออก PR/PO แล้ว');
  var now = new Date();
  setRequestFields_(ss, id, { status: STATUS.RECEIVED, received_at: now,
    admin_note: [req.admin_note, note].filter(String).join(' | ').slice(0, 500) });
  updateItems_(ss, id, function (item) { // ของที่ยังไม่ได้ติ๊กรับรายชิ้น ถือว่ามาพร้อมกันวันนี้
    if (Number(item.qty) > 0 && !item.received_at) item.received_at = now;
    return item;
  });
  mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ของเข้าแล้ว',
    'สินค้าตามคำขอ ' + esc_(id) + ' (PR/PO ' + esc_(req.po_no) + ') รับเข้า Store แล้ว' + appLink_(id));
  logActivity_('markReceived', 'ok', id + ' by ' + me.email);
  return loadRequestsFor_(me);
}

// ============================================================ reminders (daily trigger)

/** Re-sends the confirm email for requests still waiting on the PC after REMINDER_AFTER_DAYS. */
function remindPendingConfirmations() {
  ownerOnly_();
  var ss = db_();
  var days = CONFIG.REMINDER_AFTER_DAYS || 3;
  var cutoff = Date.now() - days * 86400000;
  var master = readMasterPc_(ss);
  var n = 0;
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS).forEach(function (r) {
    if (r.status !== STATUS.WAIT_PC || !r.sent_at) return;
    var last = r.reminded_at instanceof Date ? r.reminded_at : r.sent_at;
    if (!(last instanceof Date) || last.getTime() > cutoff) return;
    var owners = splitEmails_(r.assigned_to);
    if (!owners.length) return;
    sendPcConfirmEmail_(r.request_id, r.pc_code, (master[r.pc_code] || {}).pc_name || r.pc_name, owners, readItems_(ss, r.request_id), r.note, true);
    setRequestFields_(ss, r.request_id, { reminded_at: new Date() });
    n++;
  });
  logActivity_('remindPendingConfirmations', 'ok', n + ' reminders');
  return n;
}

/** Admin tab button: daily reminder trigger at 09:00 (replaces older copies). */
function installReminderTrigger() {
  requireAdmin_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'remindPendingConfirmations') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('remindPendingConfirmations').timeBased().everyDays(1).atHour(9).create();
  return 'ตั้งเตือนอัตโนมัติทุกวัน 09:00 แล้ว (เตือนเมื่อ PC ไม่ตอบเกิน ' + (CONFIG.REMINDER_AFTER_DAYS || 3) + ' วัน)';
}

// ============================================================ listing

function listMyRequests() {
  return loadRequestsFor_(requireActive_());
}

function listAllRequests() {
  requireAdmin_();
  return loadRequests_(function () { return true; });
}

/** Admins see everything; others see what they requested, are assigned, or must approve. */
function loadRequestsFor_(me) {
  if (me.role === 'admin') return loadRequests_(function () { return true; });
  return loadRequests_(function (r) {
    return r.requester_email === me.email || splitEmails_(r.assigned_to).indexOf(me.email) !== -1 ||
      String(r.manager_email).toLowerCase() === me.email || splitEmails_(r.approval_to).indexOf(me.email) !== -1;
  });
}

function loadRequests_(filterFn) {
  var ss = db_();
  var items = {};
  var files = quoteFilesByRequest_(ss);
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS).forEach(function (i) {
    (items[i.request_id] = items[i.request_id] || []).push(toClientItem_(i));
  });
  return readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS)
    .filter(function (r) { return r.request_id && filterFn(r); })
    .map(function (r) { return toClientRequest_(r, items[r.request_id] || [], files[r.request_id] || []); })
    .reverse()
    .slice(0, 400);
}

function toClientItem_(i) {
  return { sku: String(i.sku), name: String(i.name), pc_code: String(i.pc_code), pc_name: String(i.pc_name || ''),
    unit: String(i.unit), suggested_qty: Number(i.suggested_qty) || 0, qty: Number(i.qty), unit_cost: Number(i.unit_cost),
    amount: Number(i.amount), note: String(i.note || ''), budget_key: String(i.budget_key || ''),
    budget_label: String(i.budget_label || ''), budget_year: Number(i.budget_year) || 0, budget_month: Number(i.budget_month) || 0,
    store_balance: i.store_balance === '' || i.store_balance == null ? '' : Number(i.store_balance),
    over_budget: String(i.over_budget || ''), over_reason: String(i.over_reason || ''),
    // PR/PO + การส่งของ รายชิ้น (Delivery.js)
    pr_no: String(i.pr_no || ''), po_no: String(i.po_no || ''), po_at: dayString_(i.po_at),
    eta_date: dayString_(i.eta_date), received_at: dayString_(i.received_at) };
}

/** google.script.run can't return Date objects — convert to strings. */
function toClientRequest_(r, items, files) {
  var out = {};
  REQUEST_HEADER.forEach(function (h) {
    var v = r[h];
    out[h] = v instanceof Date ? Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : (v == null ? '' : v);
  });
  out.status = normStatus_(r.status);
  out.approvers = splitEmails_(r.approval_to);
  out.assigned = splitEmails_(r.assigned_to);
  out.month_label = r.budget_month ? THAI_MONTH_NAMES[Number(r.budget_month)] + ' ' + (Number(r.budget_year) || '') : '';
  out.items = items;
  out.files = files || [];
  return out;
}

// ============================================================ sheet helpers

function nextRequestId_(ss) {
  var sheet = ensureHeader_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER).sheet;
  // max sequence of today's ids + 1 (row count would repeat ids after a request is deleted)
  var prefix = 'REQ-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd') + '-';
  var max = 0;
  sheet.getDataRange().getValues().forEach(function (r) {
    var v = String(r[0]);
    if (v.indexOf(prefix) === 0) max = Math.max(max, Number(v.slice(prefix.length)) || 0);
  });
  return prefix + ('000' + (max + 1)).slice(-4);
}

/** opts (budgetOptions_) is given when items carry budgetKey/month (user submit). */
function itemRows_(id, items, fromRound, opts) {
  return items.map(function (i) {
    var qty = Math.max(0, Math.round(Number(i.qty) || 0));
    var cost = Math.max(0, Number(i.unitCost) || 0);
    var row = { request_id: id, sku: String(i.sku), name: String(i.name || '').slice(0, 200), pc_code: String(i.pc || ''),
      pc_name: String(i.pcName || ''), unit: String(i.unit || ''), suggested_qty: fromRound ? qty : '',
      qty: qty, unit_cost: cost, amount: round2_(qty * cost), note: String(i.note || '').slice(0, 300),
      store_balance: i.storeBalance === '' || i.storeBalance == null || isNaN(Number(i.storeBalance)) ? '' : Number(i.storeBalance) };
    if (opts && i.budgetKey) {
      var line = opts.lines.filter(function (l) { return l.key === i.budgetKey; })[0];
      row.budget_key = String(i.budgetKey);
      row.budget_label = line ? line.label : '';
      row.budget_year = opts.year;
      row.budget_month = Number(i.month);
    }
    return row;
  });
}

function sumAmount_(rows) {
  return round2_(rows.reduce(function (a, r) { return a + (Number(r.amount) || 0); }, 0));
}

function findRequest_(ss, id) {
  var r = readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS).filter(function (x) { return x.request_id === id; })[0];
  if (!r) throw new Error('ไม่พบคำขอ ' + id);
  r.status = normStatus_(r.status);
  return r;
}

function readItems_(ss, id) {
  return readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS)
    .filter(function (i) { return i.request_id === id; }).map(toClientItem_);
}

function setRequestFields_(ss, id, fields) {
  var h = ensureHeader_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER);
  var values = h.sheet.getDataRange().getValues();
  var col = {};
  h.header.forEach(function (k, i) { col[k] = i; });
  for (var r = 1; r < values.length; r++) {
    if (values[r][col.request_id] !== id) continue;
    Object.keys(fields).forEach(function (k) {
      if (col[k] == null) return;
      h.sheet.getRange(r + 1, col[k] + 1).setValue(fields[k]);
    });
    return;
  }
  throw new Error('ไม่พบคำขอ ' + id);
}

/** Rewrites this request's item rows through mutate(item); returns { amount, count } of qty > 0 lines. */
function updateItems_(ss, id, mutate) {
  var h = ensureHeader_(ss, CONFIG.SHEETS.REQUEST_ITEMS, REQUEST_ITEM_HEADER);
  var values = h.sheet.getDataRange().getValues();
  var col = {};
  h.header.forEach(function (k, i) { col[k] = i; });
  var amount = 0, count = 0;
  for (var r = 1; r < values.length; r++) {
    if (values[r][col.request_id] !== id) continue;
    var item = {};
    h.header.forEach(function (k, i) { item[k] = values[r][i]; });
    item = mutate(item);
    h.sheet.getRange(r + 1, 1, 1, h.header.length).setValues([h.header.map(function (k) { return item[k] == null ? '' : item[k]; })]);
    if (Number(item.qty) > 0) { amount += Number(item.amount) || 0; count++; }
  }
  return { amount: round2_(amount), count: count };
}

function canActAsPc_(me, req) {
  return me.role === 'admin' || splitEmails_(req.assigned_to).indexOf(me.email) !== -1;
}

function requestPeople_(req) {
  return uniq_([req.requester_email].concat(splitEmails_(req.assigned_to)).concat([req.confirmed_by]).filter(String));
}

function uniq_(a) {
  return a.filter(function (x, i) { return x && a.indexOf(x) === i; });
}

function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }

// ============================================================ email

function appLink_(id, label) {
  var url = webAppUrl_();
  if (!url) return '';
  return '<p><a href="' + url + (id ? '?req=' + encodeURIComponent(id) : '') + '">' + (label || 'เปิดระบบ Store Reorder') + '</a></p>';
}

function mailPeople_(to, subject, html) {
  to = uniq_((to || []).map(function (e) { return String(e).toLowerCase(); }).filter(function (e) { return /@/.test(e); }));
  if (!to.length) return;
  MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html });
}

function mailAdmins_(subject, html) {
  mailPeople_(adminEmails_(), subject, html);
}

function itemsTable_(items) {
  return '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px">' +
    '<tr><th>PC</th><th>รหัส</th><th>สินค้า</th><th>จำนวน</th><th>บาท</th></tr>' +
    items.filter(function (i) { return Number(i.qty) > 0; }).map(function (i) {
      return '<tr><td>' + esc_(i.pc_code) + '</td><td>' + esc_(i.sku) + '</td><td>' + esc_(i.name) + '</td><td align="right">' +
        esc_(i.qty) + ' ' + esc_(i.unit) + '</td><td align="right">' + fmtMoney_(i.amount) + '</td></tr>';
    }).join('') + '</table>';
}

/** Uses Claude when ANTHROPIC_API_KEY is set; any failure falls back to the plain template. */
function aiDraft_(fn, fallback) {
  try {
    if (!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY')) return fallback;
    var text = fn();
    return text ? text : fallback;
  } catch (e) {
    return fallback;
  }
}

/** (5) email to the PC owners with the suggested list and a link to confirm. */
function sendPcConfirmEmail_(id, pc, pcName, owners, items, note, isReminder) {
  var fallback = 'เรียน ผู้รับผิดชอบ ' + esc_(pc) + ' ' + esc_(pcName) + '<br><br>' +
    'ระบบแนะนำให้สั่งซื้อรายการด้านล่าง กรุณายืนยัน/แก้จำนวน และเลือก Budget กับเดือนที่จะใช้ของ ภายใน ' +
    (CONFIG.REMINDER_AFTER_DAYS || 3) + ' วัน' + (note ? '<br>หมายเหตุจาก Admin: ' + esc_(note) : '');
  var body = aiDraft_(function () {
    return draftReorderEmail_(pc + ' ' + pcName, items.map(function (i) {
      return { name: i.name, balance: '', suggested_qty: i.qty + ' ' + i.unit };
    })).replace(/\{\{CONFIRM_LINK\}\}/g, '').replace(/\n/g, '<br>');
  }, fallback);
  mailPeople_(owners, (isReminder ? '[เตือน] ' : '') + '[Store Reorder] ยืนยันรายการสั่งซื้อ ' + pc + ' (' + id + ')',
    body + itemsTable_(items) + appLink_(id, 'ยืนยัน / แก้จำนวน / ไม่สั่ง'));
}

/** (7) tell admins a request is ready for review, with an AI summary when available. */
function notifyAdminsForReview_(id, resubmitted) {
  var ss = db_();
  var req = findRequest_(ss, id);
  var items = readItems_(ss, id);
  var over = req.over_budget ? ' (เกินงบ)' : '';
  var summary = aiDraft_(function () {
    return summarizeConfirmationForAdmin_(req.pc_code + ' ' + req.pc_name, items.filter(function (i) { return i.qty > 0; })
      .map(function (i) { return { name: i.name, qty: i.qty, amount: i.amount }; }), req.note).replace(/\n/g, '<br>');
  }, '');
  var from = req.source === 'round' ? req.confirmed_by : (req.requester_name || req.requester_email);
  var byLine = {}, order = [];
  items.filter(function (i) { return i.qty > 0; }).forEach(function (i) {
    var k = (i.budget_key || req.budget_key) + '|' + (i.budget_month || req.budget_month);
    if (!byLine[k]) { byLine[k] = { label: i.budget_label || req.budget_label, month: i.budget_month || Number(req.budget_month), amount: 0, over: i.over_budget, reason: i.over_reason }; order.push(k); }
    byLine[k].amount += i.amount;
  });
  var lineHtml = order.map(function (k) {
    var l = byLine[k];
    return '<li>' + esc_(l.label) + ' · เดือน ' + THAI_MONTH_NAMES[Number(l.month)] + ' · ' + fmtMoney_(l.amount) + ' บาท' +
      (l.over ? ' <b style="color:#b23b2e">ขอ Over Budget' + (l.reason ? ': ' + esc_(l.reason) : '') + '</b>' : '') + '</li>';
  }).join('');
  mailAdmins_('[Store Reorder] ' + (resubmitted ? 'แก้ไขแล้วส่งใหม่ ' : 'รอ Admin ตรวจ ') + id + ' ' + (req.pc_code || '') + ' จาก ' + from + (over ? ' (Over Budget)' : ''),
    (summary ? '<p>' + summary + '</p>' : '') +
    '<p><b>' + esc_(id) + '</b> · ยอดรวม <b>' + fmtMoney_(req.total) + '</b> บาท (ไม่รวม VAT)</p><ul>' + lineHtml + '</ul>' +
    (req.note ? '<p>หมายเหตุ: ' + esc_(req.note) + '</p>' : '') +
    ((quoteFilesByRequest_(ss)[id] || []).length ? '<p>แนบใบเสนอราคา ' + quoteFilesByRequest_(ss)[id].length + ' ไฟล์ (เปิดดูในระบบ)</p>' : '') + itemsTable_(items) +
    appLink_(id, 'เปิดระบบเพื่อตรวจ แล้วส่งอีเมลขออนุมัติผู้บริหาร'));
}

function fmtMoney_(n) {
  return (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
