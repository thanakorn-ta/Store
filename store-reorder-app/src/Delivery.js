/**
 * Delivery.js
 * PR/PO numbers, expected delivery date and receiving - per ITEM, not per request.
 *
 * One request often turns into several PR/PO numbers (one per supplier), and the
 * boxes arrive on different days. So every item row carries its own:
 *   pr_no · po_no · po_at · eta_date (ประมาณว่าของจะมาวันไหน) · received_at (ของมาจริงวันไหน)
 *
 * savePrPo()     writes those to the ticked item rows (one item, or every item that
 *                shares the same PR/PO). The first time an item gets a PO number its
 *                amount is committed to pr_po_log, so the budget is cut when the PO is
 *                issued - not when the whole request happens to be finished.
 * receiveItems() stamps the received date; the request flips to รับของแล้ว once every
 *                line with qty > 0 is in.
 *
 * The delivery analysis (lead time, on time / late, suggested ETA) is done in the page
 * from these fields - see web/app.js, tab ติดตามการส่งของ.
 */

var DELIVERY_FIELDS = ['pr_no', 'po_no', 'po_at', 'eta_date', 'received_at'];
var DEFAULT_LEAD_DAYS = 14; // ใช้เมื่อยังไม่มีประวัติการส่งของให้คำนวณ

function parseDay_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(v).trim());
  if (!m) return '';
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function dayString_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

/** ItemRow indexes the page sent us, checked against the sku it saw there. */
function pickRows_(items, rows) {
  var want = {};
  (rows || []).forEach(function (r) {
    var n = typeof r === 'object' ? Number(r.row) : Number(r);
    if (!(n >= 0)) return;
    want[n] = typeof r === 'object' && r.sku ? String(r.sku) : null;
  });
  var out = [];
  Object.keys(want).forEach(function (n) {
    var i = Number(n);
    if (!items[i]) throw new Error('ไม่พบรายการสินค้าลำดับที่ ' + (i + 1) + ' — รีเฟรชหน้าเว็บแล้วลองใหม่');
    if (want[n] && String(items[i].sku) !== want[n]) throw new Error('รายการสินค้าเปลี่ยนไปแล้ว — รีเฟรชหน้าเว็บแล้วลองใหม่');
    out.push(i);
  });
  if (!out.length) throw new Error('ยังไม่ได้เลือกรายการสินค้า');
  return out;
}

/** pr_po_log rows for the items that just got a PO number (budget is cut here). */
function commitPo_(ss, req, items, picked, poNo, who) {
  var groups = {};
  picked.forEach(function (i) {
    var it = items[i];
    if (!(Number(it.qty) > 0)) return;
    var key = it.budget_key || req.budget_key, m = Number(it.budget_month) || Number(req.budget_month);
    var k = it.pc_code + '|' + key + '|' + m;
    groups[k] = groups[k] || { pc: it.pc_code, key: key, month: m, year: Number(it.budget_year) || Number(req.budget_year), amount: 0 };
    groups[k].amount += Number(it.amount) || 0;
  });
  var rows = Object.keys(groups).map(function (k) {
    var g = groups[k];
    return { created_at: new Date(), request_id: req.request_id, po_no: poNo, pc_code: g.pc, budget_key: g.key,
      budget_year: g.year, month_number: g.month, amount: round2_(g.amount), issued_by: who };
  });
  appendObjects_(ss, CONFIG.SHEETS.PR_LOG, PR_LOG_HEADER, rows);
  return rows.reduce(function (a, r) { return a + r.amount; }, 0);
}

/**
 * savePrPo(id, { rows: [{row, sku}…], pr, po, eta })
 * Writes the numbers to the ticked items. An item that had no PO number yet and gets one
 * commits its amount to pr_po_log; the request moves to ออก PR/PO แล้ว when every line has one.
 */
function savePrPo(id, payload) {
  var me = requireAdmin_();
  payload = payload || {};
  var ss = db_();
  var req = findRequest_(ss, id);
  var ok = [STATUS.APPROVED, STATUS.PO, STATUS.RECEIVED];
  if (ok.indexOf(normStatus_(req.status)) === -1) throw new Error('ใส่เลข PR/PO ได้เฉพาะคำขอที่อนุมัติแล้ว (สถานะตอนนี้: ' + req.status + ')');
  var pr = String(payload.pr == null ? '' : payload.pr).trim().slice(0, 60);
  var po = String(payload.po == null ? '' : payload.po).trim().slice(0, 60);
  var eta = payload.eta === '' ? '' : parseDay_(payload.eta);
  if (payload.eta && !eta) throw new Error('วันที่คาดว่าจะได้รับไม่ถูกต้อง');
  if (!pr && !po && payload.eta == null) throw new Error('กรุณาใส่เลข PR หรือ PO หรือวันที่คาดว่าจะได้รับ');

  var items = readItems_(ss, id);
  var picked = pickRows_(items, payload.rows);
  var fresh = picked.filter(function (i) { return po && !String(items[i].po_no || '').trim(); });
  picked.forEach(function (i) {
    if (!po && String(items[i].po_no || '').trim() && payload.po === '') {
      throw new Error('ลบเลข PO ที่ตัดงบไปแล้วไม่ได้ — แก้เป็นเลขใหม่แทน');
    }
  });

  var n = 0;
  updateItems_(ss, id, function (item) {
    if (picked.indexOf(n++) === -1) return item;
    if (pr) item.pr_no = pr;
    if (po) {
      if (!String(item.po_no || '').trim()) item.po_at = new Date();
      item.po_no = po;
    }
    if (payload.eta != null) item.eta_date = eta;
    return item;
  });

  var committed = fresh.length ? commitPo_(ss, req, items, fresh, po, me.email) : 0;
  var after = readItems_(ss, id);
  var live = after.filter(function (i) { return Number(i.qty) > 0; });
  var allPo = live.length && live.every(function (i) { return String(i.po_no || '').trim(); });
  var fields = {};
  if (allPo && normStatus_(req.status) === STATUS.APPROVED) { fields.status = STATUS.PO; fields.po_at = new Date(); }
  var numbers = uniq_(live.map(function (i) { return String(i.po_no || '').trim(); }).filter(String));
  if (numbers.length) fields.po_no = numbers.join(', ');
  if (Object.keys(fields).length) setRequestFields_(ss, id, fields);

  if (fresh.length) {
    mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ออก PR/PO แล้ว: ' + po,
      'สินค้า ' + fresh.length + ' รายการของคำขอ ' + esc_(id) + ' ออก PR/PO เลขที่ <b>' + esc_(po) + '</b> แล้ว' +
      (eta ? '<br>คาดว่าจะได้รับ ' + esc_(dayString_(eta)) : '') + appLink_(id));
  }
  logActivity_('savePrPo', 'ok', id + ' ' + picked.length + ' รายการ ' + (po || pr) + ' ตัดงบ ' + committed + ' by ' + me.email);
  return loadRequestsFor_(me);
}

/**
 * receiveItems(id, { rows: [{row, sku}…], date, note })
 * date ว่าง = วันนี้ · คำขอจะเป็น "รับของแล้ว" เมื่อทุกรายการที่สั่งมาครบ
 */
function receiveItems(id, payload) {
  var me = requireAdmin_();
  payload = payload || {};
  var ss = db_();
  var req = findRequest_(ss, id);
  var when = payload.date ? parseDay_(payload.date) : new Date();
  if (!when) throw new Error('วันที่รับของไม่ถูกต้อง');
  var items = readItems_(ss, id);
  var picked = pickRows_(items, payload.rows);
  var clear = payload.date === '' && payload.clear === true;

  var n = 0;
  updateItems_(ss, id, function (item) {
    if (picked.indexOf(n++) === -1) return item;
    item.received_at = clear ? '' : when;
    return item;
  });

  var after = readItems_(ss, id);
  var live = after.filter(function (i) { return Number(i.qty) > 0; });
  var got = live.filter(function (i) { return !!i.received_at; });
  var fields = { admin_note: [req.admin_note, payload.note].filter(String).join(' | ').slice(0, 500) };
  if (live.length && got.length === live.length) { fields.status = STATUS.RECEIVED; fields.received_at = when; }
  else if (normStatus_(req.status) === STATUS.RECEIVED) { fields.status = STATUS.PO; fields.received_at = ''; }
  setRequestFields_(ss, id, fields);

  if (fields.status === STATUS.RECEIVED) {
    mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ของเข้าครบแล้ว',
      'สินค้าตามคำขอ ' + esc_(id) + ' (PR/PO ' + esc_(req.po_no) + ') รับเข้า Store ครบแล้ว' + appLink_(id));
  }
  logActivity_('receiveItems', 'ok', id + ' ' + picked.length + ' รายการ ' + dayString_(when) + ' (' + got.length + '/' + live.length + ') by ' + me.email);
  return loadRequestsFor_(me);
}
