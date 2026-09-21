/**
 * Requests.js
 * Purchase requests following store-ai-workflow.html:
 *
 *  (4) Admin picks the items to order ─► (5) createRound(): one request per PC,
 *      AI-drafted email to the PC owners from master_pc              [รอ PC ยืนยัน]
 *  (6) PC owner confirms / edits qty / rejects, and picks the budget
 *      line + month the goods will be used          ─► [รออนุมัติ] / [PC ปฏิเสธ]
 *      no answer in CONFIG.REMINDER_AFTER_DAYS ─► reminder email (daily trigger)
 *  (7) Admin reviews (AI summary in the email); budget check:
 *        enough      ─► approve                                         [อนุมัติ]
 *        not enough  ─► (8b) requestOverBudget(): email the PC manager
 *                        [รอหัวหน้าอนุมัติเกินงบ] ─► manager approves ─► [รออนุมัติ]
 *  (9) issuePo(): PR/PO number, commits spend to pr_po_log         [ออก PR/PO แล้ว]
 * (10) markReceived()                                                 [รับของแล้ว]
 *
 * Users can also start a request themselves from their cart (submitOrderRequest),
 * which enters at step 6's output ([รออนุมัติ]).
 *
 * Budget: an admin uploads the "Budget STT 2026 - Revise-Budget" export (Admin tab),
 * aggregated per Company x Media Location x GL Code x month in budget_master.
 *   plan      = Revise Budget when filled, else Budget
 *   reserved  = totals of requests from [รออนุมัติ] onward that are not rejected/cancelled
 *   available = plan - actual - reserved
 */

var BUDGET_HEADER = ['key', 'company', 'division', 'media_type', 'media_group', 'media_location',
  'expense_group', 'gl_code', 'gl_name', 'year', 'month_number', 'budget', 'revise_budget', 'actual', 'plan'];
var REQUEST_HEADER = ['request_id', 'created_at', 'source', 'round_id', 'pc_code', 'pc_name', 'assigned_to',
  'requester_email', 'requester_name', 'budget_key', 'budget_label', 'budget_year', 'budget_month', 'total',
  'item_count', 'available_at_submit', 'over_budget', 'status', 'note', 'sent_at', 'reminded_at',
  'confirmed_by', 'confirmed_at', 'manager_email', 'manager_decision', 'manager_note', 'manager_at',
  'decided_by', 'decided_at', 'admin_note', 'po_no', 'po_at', 'received_at'];
var REQUEST_ITEM_HEADER = ['request_id', 'sku', 'name', 'pc_code', 'pc_name', 'unit', 'suggested_qty', 'qty',
  'unit_cost', 'amount', 'note'];
var PR_LOG_HEADER = ['created_at', 'request_id', 'po_no', 'pc_code', 'budget_key', 'budget_year', 'month_number', 'amount', 'issued_by'];

var STATUS = {
  WAIT_PC: 'รอ PC ยืนยัน', PC_REJECTED: 'PC ปฏิเสธ',
  PENDING: 'รออนุมัติ', OVER_WAIT: 'รอหัวหน้าอนุมัติเกินงบ',
  APPROVED: 'อนุมัติ', REJECTED: 'ไม่อนุมัติ', CANCELLED: 'ยกเลิก',
  PO: 'ออก PR/PO แล้ว', RECEIVED: 'รับของแล้ว'
};
var RESERVING = [STATUS.PENDING, STATUS.OVER_WAIT, STATUS.APPROVED, STATUS.PO, STATUS.RECEIVED];
var THAI_MONTH_NAMES = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

// ============================================================ budget

/** Admin uploads budget rows already aggregated by the page (web/parsers.js parseBudget). */
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
    var sheet = ensureSheet_(db_(), CONFIG.SHEETS.BUDGET, BUDGET_HEADER);
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

/** { "key|month": amount } for requests that hold budget. excludeId skips one request. */
function reservedByBudget_(ss, excludeId) {
  var out = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS).forEach(function (r) {
    if (RESERVING.indexOf(r.status) === -1 || !r.budget_key || r.request_id === excludeId) return;
    var k = r.budget_key + '|' + Number(r.budget_month);
    out[k] = (out[k] || 0) + (Number(r.total) || 0);
  });
  return out;
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

// ============================================================ user: submit from cart (enters as รออนุมัติ)

/** payload: { budgetKey, month, note, items: [{ sku, name, pc, pcName, unit, qty, unitCost, note }] } */
function submitOrderRequest(payload) {
  var me = requireActive_();
  var items = (payload && payload.items || []).filter(function (i) { return Number(i.qty) > 0; });
  if (!items.length) throw new Error('ยังไม่มีรายการที่จำนวนมากกว่า 0');
  if (items.length > 500) throw new Error('รายการมากเกินไป (สูงสุด 500)');
  var month = Number(payload.month);
  if (!(month >= 1 && month <= 12)) throw new Error('กรุณาเลือกเดือนที่จะใช้ของ');

  var ss = db_();
  var opts = budgetOptions_(ss);
  var av = availableFor_(opts, payload.budgetKey, month);
  if (!av) throw new Error('กรุณาเลือก Budget');

  var id, total, rows;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    id = nextRequestId_(ss);
    rows = itemRows_(id, items, false);
    total = sumAmount_(rows);
    var pcs = uniq_(items.map(function (i) { return i.pc; }));
    var master = readMasterPc_(ss);
    appendObject_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER, {
      request_id: id, created_at: new Date(), source: 'user', pc_code: pcs.join(', '),
      requester_email: me.email, requester_name: me.name, budget_key: av.line.key, budget_label: av.line.label,
      budget_year: opts.year, budget_month: month, total: total, item_count: rows.length,
      available_at_submit: round2_(av.available), over_budget: total > av.available ? 'เกินงบ' : '',
      status: STATUS.PENDING, note: String(payload.note || '').slice(0, 500),
      confirmed_by: me.email, confirmed_at: new Date(),
      manager_email: pcs.length === 1 && master[pcs[0]] ? master[pcs[0]].manager_email : ''
    });
    appendObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS, REQUEST_ITEM_HEADER, rows);
  } finally {
    lock.releaseLock();
  }
  notifyAdminsForReview_(id);
  logActivity_('submitOrderRequest', 'ok', id + ' by ' + me.email + ' total ' + total);
  return { id: id, total: total, available: av.available, overBudget: total > av.available };
}

// ============================================================ (6) PC owner confirms / rejects

/** payload: { items: [{ sku, qty, note }], budgetKey, month, note } — qty 0 drops the line. */
function confirmRequest(id, payload) {
  var me = requireActive_();
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
  var remaining = readItems_(ss, id).filter(function (i) {
    return (edits[i.sku] ? Math.round(Number(edits[i.sku].qty) || 0) : i.qty) > 0;
  });
  if (!remaining.length) throw new Error('ทุกรายการเป็น 0 — ถ้าไม่ต้องการสั่ง ให้กด "ไม่สั่งรอบนี้"');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var total;
  try {
    total = updateItems_(ss, id, function (item) {
      var e = edits[item.sku];
      if (!e) return item;
      item.qty = Math.max(0, Math.round(Number(e.qty) || 0));
      item.amount = round2_(item.qty * (Number(item.unit_cost) || 0));
      if (e.note != null) item.note = String(e.note).slice(0, 300);
      return item;
    });
    setRequestFields_(ss, id, {
      budget_key: av.line.key, budget_label: av.line.label, budget_year: opts.year, budget_month: month,
      total: total.amount, item_count: total.count, available_at_submit: round2_(av.available),
      over_budget: total.amount > av.available ? 'เกินงบ' : '', status: STATUS.PENDING,
      note: String(payload.note || req.note || '').slice(0, 500), confirmed_by: me.email, confirmed_at: new Date()
    });
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
  if ([STATUS.PENDING, STATUS.WAIT_PC, STATUS.OVER_WAIT].indexOf(req.status) === -1) throw new Error('ยกเลิกไม่ได้ในสถานะ ' + req.status);
  setRequestFields_(ss, id, { status: STATUS.CANCELLED, decided_by: me.email, decided_at: new Date(), admin_note: 'ยกเลิกโดย ' + me.email });
  logActivity_('cancelRequest', 'ok', id + ' by ' + me.email);
  return loadRequestsFor_(me);
}

// ============================================================ (7)(8) admin review + budget check

/** decision: 'approve' | 'reject'. Over-budget needs the manager's approval first (8b). */
function decideRequest(id, decision, note) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  if (req.status !== STATUS.PENDING) throw new Error('คำขอนี้ไม่ได้อยู่ในสถานะรออนุมัติ (' + req.status + ')');
  var approve = decision === 'approve';
  if (approve) {
    var av = availableFor_({ lines: budgetOptions_(ss).lines, reserved: reservedByBudget_(ss, id) }, req.budget_key, Number(req.budget_month));
    var over = !av || Number(req.total) > av.available;
    if (over && req.manager_decision !== 'อนุมัติ') {
      throw new Error('งบไม่พอ (คงเหลือ ' + fmtMoney_(av ? av.available : 0) + ' บาท) — กด "ขออนุมัติเกินงบ" ให้หัวหน้า PC อนุมัติก่อน');
    }
  }
  setRequestFields_(ss, id, { status: approve ? STATUS.APPROVED : STATUS.REJECTED, decided_by: me.email,
    decided_at: new Date(), admin_note: String(note || '').slice(0, 500) });
  mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ' + (approve ? STATUS.APPROVED : STATUS.REJECTED),
    'คำขอ ' + esc_(id) + ' (' + fmtMoney_(req.total) + ' บาท): <b>' + (approve ? STATUS.APPROVED : STATUS.REJECTED) + '</b>' +
    (note ? '<br>หมายเหตุจาก Admin: ' + esc_(note) : '') + appLink_(id));
  logActivity_('decideRequest', 'ok', id + ' ' + decision + ' by ' + me.email);
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
  // Commit spend per PC (Budget.js / checkBudget_ reads pr_po_log)
  var byPc = {};
  readItems_(ss, id).forEach(function (i) { byPc[i.pc_code] = (byPc[i.pc_code] || 0) + Number(i.amount); });
  appendObjects_(ss, CONFIG.SHEETS.PR_LOG, PR_LOG_HEADER, Object.keys(byPc).map(function (pc) {
    return { created_at: new Date(), request_id: id, po_no: poNo, pc_code: pc, budget_key: req.budget_key,
      budget_year: req.budget_year, month_number: req.budget_month, amount: round2_(byPc[pc]), issued_by: me.email };
  }));
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
  setRequestFields_(ss, id, { status: STATUS.RECEIVED, received_at: new Date(),
    admin_note: [req.admin_note, note].filter(String).join(' | ').slice(0, 500) });
  mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ของเข้าแล้ว',
    'สินค้าตามคำขอ ' + esc_(id) + ' (PR/PO ' + esc_(req.po_no) + ') รับเข้า Store แล้ว' + appLink_(id));
  logActivity_('markReceived', 'ok', id + ' by ' + me.email);
  return loadRequestsFor_(me);
}

// ============================================================ reminders (daily trigger)

/** Re-sends the confirm email for requests still waiting on the PC after REMINDER_AFTER_DAYS. */
function remindPendingConfirmations() {
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
      String(r.manager_email).toLowerCase() === me.email;
  });
}

function loadRequests_(filterFn) {
  var ss = db_();
  var items = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS).forEach(function (i) {
    (items[i.request_id] = items[i.request_id] || []).push(toClientItem_(i));
  });
  return readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS)
    .filter(function (r) { return r.request_id && filterFn(r); })
    .map(function (r) { return toClientRequest_(r, items[r.request_id] || []); })
    .reverse()
    .slice(0, 400);
}

function toClientItem_(i) {
  return { sku: String(i.sku), name: String(i.name), pc_code: String(i.pc_code), pc_name: String(i.pc_name || ''),
    unit: String(i.unit), suggested_qty: Number(i.suggested_qty) || 0, qty: Number(i.qty), unit_cost: Number(i.unit_cost),
    amount: Number(i.amount), note: String(i.note || '') };
}

/** google.script.run can't return Date objects — convert to strings. */
function toClientRequest_(r, items) {
  var out = {};
  REQUEST_HEADER.forEach(function (h) {
    var v = r[h];
    out[h] = v instanceof Date ? Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : (v == null ? '' : v);
  });
  out.assigned = splitEmails_(r.assigned_to);
  out.month_label = r.budget_month ? THAI_MONTH_NAMES[Number(r.budget_month)] + ' ' + (Number(r.budget_year) || '') : '';
  out.items = items;
  return out;
}

// ============================================================ sheet helpers

function nextRequestId_(ss) {
  var sheet = ensureHeader_(ss, CONFIG.SHEETS.REQUESTS, REQUEST_HEADER).sheet;
  return 'REQ-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd') + '-' + ('000' + sheet.getLastRow()).slice(-4);
}

function itemRows_(id, items, fromRound) {
  return items.map(function (i) {
    var qty = Math.max(0, Math.round(Number(i.qty) || 0));
    var cost = Math.max(0, Number(i.unitCost) || 0);
    return { request_id: id, sku: String(i.sku), name: String(i.name || '').slice(0, 200), pc_code: String(i.pc || ''),
      pc_name: String(i.pcName || ''), unit: String(i.unit || ''), suggested_qty: fromRound ? qty : '',
      qty: qty, unit_cost: cost, amount: round2_(qty * cost), note: String(i.note || '').slice(0, 300) };
  });
}

function sumAmount_(rows) {
  return round2_(rows.reduce(function (a, r) { return a + (Number(r.amount) || 0); }, 0));
}

function findRequest_(ss, id) {
  var r = readSheetAsObjects_(ss, CONFIG.SHEETS.REQUESTS).filter(function (x) { return x.request_id === id; })[0];
  if (!r) throw new Error('ไม่พบคำขอ ' + id);
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
  to = uniq_((to || []).map(function (e) { return String(e).toLowerCase(); }));
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
function notifyAdminsForReview_(id) {
  var ss = db_();
  var req = findRequest_(ss, id);
  var items = readItems_(ss, id);
  var over = req.over_budget ? ' (เกินงบ)' : '';
  var summary = aiDraft_(function () {
    return summarizeConfirmationForAdmin_(req.pc_code + ' ' + req.pc_name, items.filter(function (i) { return i.qty > 0; })
      .map(function (i) { return { name: i.name, qty: i.qty, amount: i.amount }; }), req.note).replace(/\n/g, '<br>');
  }, '');
  var from = req.source === 'round' ? req.confirmed_by : (req.requester_name || req.requester_email);
  mailAdmins_('[Store Reorder] รออนุมัติ ' + id + ' ' + (req.pc_code || '') + ' ยืนยันโดย ' + from + over,
    (summary ? '<p>' + summary + '</p>' : '') +
    '<p><b>' + esc_(id) + '</b> · Budget: ' + esc_(req.budget_label) + ' · ใช้เดือน ' + THAI_MONTH_NAMES[Number(req.budget_month)] +
    '<br>ยอดรวม <b>' + fmtMoney_(req.total) + '</b> บาท · งบคงเหลือก่อนคำขอนี้ ' + fmtMoney_(req.available_at_submit) + ' บาท' +
    (over ? ' <b style="color:#b23b2e">เกินงบ ' + fmtMoney_(Number(req.total) - Number(req.available_at_submit)) + ' บาท</b>' : '') +
    (req.note ? '<br>หมายเหตุ: ' + esc_(req.note) : '') + '</p>' + itemsTable_(items) + appLink_(id, 'เปิดระบบเพื่อตรวจ/อนุมัติ'));
}

function fmtMoney_(n) {
  return (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
