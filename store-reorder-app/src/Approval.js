/**
 * Approval.js
 * Admin → head approval email in the team's format
 * ("ขออนุมัติสั่งซื้ออุปกรณ์ (ทีม Static)"), recording the head's answer, and the
 * follow-up email to Purchasing ("ได้รับการอนุมัติสั่งซื้ออุปกรณ์ …").
 *
 * Default recipients / wording live in the `settings` sheet (Admin tab → ตั้งค่าอีเมล).
 */

var SETTINGS_HEADER = ['key', 'value'];
var SETTINGS_DEFAULTS = {
  team_name: 'ทีม Static',
  approval_to: 'thanakorn@planbmedia.co.th', // first stop: order list + budget go here, then on to the head
  approval_cc: '',
  approval_greeting: 'ผู้บริหาร',
  approval_intro: '',          // extra line under the budget summary, e.g. "****เบื้องต้นได้ปรึกษา … เรียบร้อยค่ะ"
  approval_link: '',           // optional link shown as "โปรดคลิกลิ้งค์"
  purchasing_to: '',
  purchasing_cc: '',
  signature: ''
};
var THAI_MONTH_FULL = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

// ============================================================ settings

function readSettings_(ss) {
  var out = {};
  Object.keys(SETTINGS_DEFAULTS).forEach(function (k) { out[k] = SETTINGS_DEFAULTS[k]; });
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) return out;
  readSheetAsObjects_(ss, CONFIG.SHEETS.SETTINGS).forEach(function (r) {
    if (!r.key || !Object.prototype.hasOwnProperty.call(SETTINGS_DEFAULTS, r.key)) return;
    var v = String(r.value == null ? '' : r.value);
    if (v !== '' || r.key !== 'approval_to') out[r.key] = v; // blank recipient falls back to the default
  });
  return out;
}

function getAppSettings() {
  requireAdmin_();
  return readSettings_(db_());
}

function saveAppSettings(values) {
  var me = requireAdmin_();
  var ss = db_();
  var cur = readSettings_(ss);
  Object.keys(SETTINGS_DEFAULTS).forEach(function (k) {
    if (values && values[k] != null) cur[k] = String(values[k]).slice(0, 2000);
  });
  var sheet = ensureSheet_(ss, CONFIG.SHEETS.SETTINGS, SETTINGS_HEADER);
  sheet.clearContents();
  var rows = [SETTINGS_HEADER].concat(Object.keys(cur).map(function (k) { return [k, cur[k]]; }));
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  logActivity_('saveAppSettings', 'ok', me.email);
  return cur;
}

// ============================================================ approval email

/**
 * kind 'approval'   → to the head:        "ขออนุมัติสั่งซื้ออุปกรณ์ (ทีม …)"
 * kind 'purchasing' → to Purchasing after: "ได้รับการอนุมัติสั่งซื้ออุปกรณ์ (ทีม …)"
 * Returns the default recipients + subject + preview so the admin can edit before sending.
 */
function getApprovalDraft(id, kind, overrides) {
  requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  var s = readSettings_(ss);
  kind = kind === 'purchasing' ? 'purchasing' : 'approval';
  var o = kind === 'purchasing'
    ? { to: s.purchasing_to, cc: s.purchasing_cc, greeting: 'ฝ่ายจัดซื้อ', intro: s.approval_intro, link: s.approval_link }
    : { to: req.approval_to || s.approval_to, cc: req.approval_cc || s.approval_cc, greeting: s.approval_greeting,
        intro: s.approval_intro, link: s.approval_link };
  o.kind = kind;
  o.subject = approvalSubject_(kind, s);
  // the composer's current values, so the preview shows what will be sent
  ['to', 'cc', 'subject', 'greeting', 'intro', 'link'].forEach(function (k) {
    if (overrides && overrides[k] != null) o[k] = String(overrides[k]);
  });
  o.html = buildApprovalHtml_(ss, req, o, s, kind);
  o.files = quoteFilesByRequest_(ss)[id] || [];   // quotations the requester attached
  o.excelName = orderFileName_(req) + '.xlsx';    // order list + budget, generated at send time
  return o;
}

/**
 * Attachments for the approval / purchasing email: the generated Excel file
 * (order list + budget per line/month) and the quotations.
 * opts.attachExcel (default true), opts.attachFileIds (default: all of this request's files).
 */
function approvalAttachments_(ss, req, opts) {
  var out = [];
  if (opts.attachExcel !== false) out.push(orderSheetBlob_(ss, req));
  var ids = opts.attachFileIds == null ? null : [].concat(opts.attachFileIds);
  return out.concat(quoteBlobs_(ss, req.request_id, ids));
}

function orderFileName_(req) {
  return 'รายการสั่งซื้อและ Budget ' + req.request_id;
}

/**
 * Excel workbook: sheet 1 = the 14 columns of the email table, sheet 2 = budget per line/month.
 * Built as a temporary Google Sheet exported to .xlsx; falls back to a UTF-8 CSV if that fails.
 */
function orderSheetBlob_(ss, req) {
  var d = approvalData_(ss, req);
  var name = orderFileName_(req);
  var head = ['Company/นามบริษัท', 'Media Type', 'Sub Media Type', 'Expense Group/Part Code Detail.', 'Remark ตาม Budget',
    'รหัสสินค้า', 'รายการสั่งซื้ออุปกรณ์', 'จำนวนที่สั่งซื้อ', 'หน่วย', 'จำนวนคงเหลือในStore', 'ขอบเขตการใช้งาน',
    'ค่าใช้จ่ายตาม Budget', 'ค่าใช้จ่ายโดยประมาณ', 'เดือนสั่งอุปกรณ์', 'สถานะการขออนุมัติ', 'Over Budget / เหตุผล'];
  var rows = d.items.map(function (i) {
    return [i.line.company, i.line.mediaType || i.line.mediaGroup || '', i.line.location || i.line.mediaGroup || '',
      i.line.expenseGroup || i.line.glName || '', i.remark, i.sku, i.name, i.qty, i.unit, i.balance, i.note,
      i.plan, i.amount, THAI_MONTH_FULL[i.month] + ' ' + d.be, normStatus_(req.status), i.over ? 'Over Budget: ' + i.overReason : ''];
  });
  rows.push(['รวม', '', '', '', '', '', '', '', '', '', '', d.bgTotal, d.total, '', '', '']);
  var seen = {}, bRows = [];
  d.items.forEach(function (i) {
    var k = i.key + '|' + i.month;
    if (seen[k]) { seen[k][4] += i.amount; return; }
    seen[k] = [i.line.label, i.line.glCode || '', THAI_MONTH_FULL[i.month] + ' ' + d.be, i.plan, i.amount, i.over ? 'Over Budget' : '', i.overReason];
    bRows.push(seen[k]);
  });
  var tmp = null;
  var bHead = ['Budget', 'GL Code', 'เดือน', 'งบตาม BG', 'คำขอนี้', 'สถานะงบ', 'เหตุผล Over Budget'];
  var info = [['เลขที่คำขอ', req.request_id], ['ผู้ขอ', req.requester_name || req.requester_email], ['วันที่', req.created_at],
    ['หมายเหตุ', req.note || ''], ['งบตาม BG รวม', d.bgTotal], ['ค่าใช้จ่ายโดยประมาณ (ไม่รวม VAT)', d.total]];
  try {
    tmp = SpreadsheetApp.create(name);
    var s1 = tmp.getSheets()[0].setName('รายการสั่งซื้อ');
    s1.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s1.getRange(2, 1, rows.length, head.length).setValues(rows);
    var s2 = tmp.insertSheet('Budget');
    s2.getRange(1, 1, info.length, 2).setValues(info);
    s2.getRange(info.length + 2, 1, 1, bHead.length).setValues([bHead]).setFontWeight('bold');
    if (bRows.length) s2.getRange(info.length + 3, 1, bRows.length, bHead.length).setValues(bRows);
    SpreadsheetApp.flush();
    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
    if (resp.getResponseCode() !== 200) throw new Error('export ' + resp.getResponseCode());
    return resp.getBlob().setName(name + '.xlsx');
  } catch (e) {
    var csvCell = function (v) { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var csv = [head].concat(rows).concat([[]]).concat(info).concat([[]]).concat([bHead]).concat(bRows)
      .map(function (r) { return r.map(csvCell).join(','); }).join('\r\n');
    if (tmp) { try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e2) { /* already trashed */ } }
    return Utilities.newBlob(String.fromCharCode(0xFEFF) + csv, 'text/csv', name + '.csv'); // BOM so Excel reads Thai as UTF-8
  }
}

/** Admin sends the approval email to the head. รอ Admin ตรวจ → รอผู้บริหารอนุมัติ (resend allowed while waiting). */
function sendApprovalEmail(id, opts) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  var st = normStatus_(req.status);
  if ([STATUS.PENDING, STATUS.APPROVAL_WAIT].indexOf(st) === -1) throw new Error('ส่งขออนุมัติได้เฉพาะคำขอที่รอ Admin ตรวจ (' + st + ')');
  opts = opts || {};
  var to = splitEmails_(opts.to), cc = splitEmails_(opts.cc);
  if (!to.length) throw new Error('กรุณาใส่อีเมลผู้รับ (ถึง)');
  var s = readSettings_(ss);
  var subject = String(opts.subject || approvalSubject_('approval', s)).slice(0, 250) + ' [' + id + ']';
  var html = buildApprovalHtml_(ss, req, opts, s, 'approval') +
    appLink_(id, 'เปิดระบบเพื่ออนุมัติ / ไม่อนุมัติ (สำหรับผู้อนุมัติที่เป็นสมาชิก)');
  var attachments = approvalAttachments_(ss, req, opts);
  MailApp.sendEmail({ to: to.join(','), cc: cc.join(','), replyTo: me.email, subject: subject, htmlBody: html, attachments: attachments,
    name: 'Store Reorder · ' + (me.name || me.email) });
  setRequestFields_(ss, id, { status: STATUS.APPROVAL_WAIT, approval_to: to.join(', '), approval_cc: cc.join(', '),
    approval_subject: subject, approval_sent_at: new Date(), decided_by: me.email,
    admin_note: String(opts.adminNote || req.admin_note || '').slice(0, 500) });
  mailPeople_(requestPeople_(req), '[Store Reorder] ' + id + ' ส่งขออนุมัติผู้บริหารแล้ว',
    'คำขอ ' + esc_(id) + ' (' + fmtMoney_(req.total) + ' บาท) Admin ตรวจแล้ว และส่งอีเมลขออนุมัติถึง ' + esc_(to.join(', ')) + appLink_(id));
  logActivity_('sendApprovalEmail', 'ok', id + ' -> ' + to.join(',') + ' by ' + me.email + ' attachments ' + attachments.length);
  return loadRequestsFor_(me);
}

/**
 * Head's answer: admin records it after the head replies "Approved" by email,
 * or the head approves in the app when they are a member listed in approval_to.
 */
function recordApproval(id, approve, note) {
  var me = requireActive_();
  var ss = db_();
  var req = findRequest_(ss, id);
  var st = normStatus_(req.status);
  if (st !== STATUS.APPROVAL_WAIT) throw new Error('คำขอนี้ไม่ได้รอผู้บริหารอนุมัติ (' + st + ')');
  if (me.role !== 'admin' && splitEmails_(req.approval_to).indexOf(me.email) === -1) throw new Error('เฉพาะผู้อนุมัติหรือ Admin เท่านั้น');
  var result = approve ? STATUS.APPROVED : STATUS.REJECTED;
  setRequestFields_(ss, id, { status: result, approved_by: me.role === 'admin' ? (req.approval_to + ' (บันทึกโดย ' + me.email + ')') : me.email,
    approved_at: new Date(), approval_note: String(note || '').slice(0, 500) });
  mailPeople_(requestPeople_(req).concat(me.role === 'admin' ? [] : adminEmails_()), '[Store Reorder] ' + id + ' ' + result,
    'คำขอ ' + esc_(id) + ' (' + fmtMoney_(req.total) + ' บาท): <b>' + result + '</b>' +
    (note ? '<br>หมายเหตุ: ' + esc_(note) : '') + (approve ? '<br>ขั้นต่อไป: Admin ส่งแจ้งฝ่ายจัดซื้อ / ออก PR/PO' : '') + appLink_(id));
  logActivity_('recordApproval', 'ok', id + ' ' + (approve ? 'approve' : 'reject') + ' by ' + me.email);
  return loadRequestsFor_(me);
}

/** After approval: "เรียน ฝ่ายจัดซื้อ … ได้รับการอนุมัติสั่งซื้ออุปกรณ์ …". */
function sendPurchasingEmail(id, opts) {
  var me = requireAdmin_();
  var ss = db_();
  var req = findRequest_(ss, id);
  var st = normStatus_(req.status);
  if ([STATUS.APPROVED, STATUS.PO].indexOf(st) === -1) throw new Error('ส่งฝ่ายจัดซื้อได้หลังได้รับอนุมัติแล้ว (' + st + ')');
  opts = opts || {};
  var to = splitEmails_(opts.to), cc = splitEmails_(opts.cc);
  if (!to.length) throw new Error('กรุณาใส่อีเมลฝ่ายจัดซื้อ (ถึง)');
  var s = readSettings_(ss);
  var subject = String(opts.subject || approvalSubject_('purchasing', s)).slice(0, 250) + ' [' + id + ']';
  MailApp.sendEmail({ to: to.join(','), cc: cc.join(','), replyTo: me.email, subject: subject,
    htmlBody: buildApprovalHtml_(ss, req, opts, s, 'purchasing'), attachments: approvalAttachments_(ss, req, opts),
    name: 'Store Reorder · ' + (me.name || me.email) });
  setRequestFields_(ss, id, { purchasing_sent_at: new Date() });
  logActivity_('sendPurchasingEmail', 'ok', id + ' -> ' + to.join(',') + ' by ' + me.email);
  return loadRequestsFor_(me);
}

function approvalSubject_(kind, s) {
  var team = s.team_name ? ' (' + s.team_name + ')' : '';
  return (kind === 'purchasing' ? 'ได้รับการอนุมัติสั่งซื้ออุปกรณ์' : 'ขออนุมัติสั่งซื้ออุปกรณ์') + team;
}

/** Items of a request with their budget line/month details, plus the BG and estimate totals. */
function approvalData_(ss, req) {
  var opts = budgetOptions_(ss);
  var lineBy = {};
  opts.lines.forEach(function (l) { lineBy[l.key] = l; });
  var year = Number(req.budget_year) || opts.year || new Date().getFullYear();
  var be = year + 543;
  var items = readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_ITEMS).filter(function (i) {
    return i.request_id === req.request_id && Number(i.qty) > 0;
  }).map(function (i) {
    var key = String(i.budget_key || req.budget_key || '');
    var m = Number(i.budget_month) || Number(req.budget_month) || 0;
    var line = lineBy[key] || { label: String(i.budget_label || req.budget_label || ''), months: {}, company: '', location: '', expenseGroup: '', glName: '' };
    var mb = line.months[m] || { plan: 0, remark: '' };
    return { name: String(i.name), sku: String(i.sku), qty: Number(i.qty), unit: String(i.unit || ''), amount: Number(i.amount) || 0,
      note: String(i.note || ''), balance: i.store_balance === '' || i.store_balance == null ? '' : Number(i.store_balance),
      key: key, month: m, line: line, plan: Number(mb.plan) || 0, remark: String(mb.remark || ''),
      over: !!i.over_budget, overReason: String(i.over_reason || '') };
  });

  // BG total = plan of each distinct line/month used
  var seen = {}, bgTotal = 0;
  items.forEach(function (i) { var k = i.key + '|' + i.month; if (!seen[k]) { seen[k] = 1; bgTotal += i.plan; } });
  var total = items.reduce(function (a, i) { return a + i.amount; }, 0);
  var months = uniq_(items.map(function (i) { return String(i.month); })).map(Number).sort(function (a, b) { return a - b; });
  return { items: items, bgTotal: bgTotal, total: total, months: months, be: be };
}

/**
 * Body in the team's format:
 *   เรียน / สำเนา / เรื่อง, Budget months, BG total vs estimate (ไม่รวม VAT)
 *   per "สื่อ {location} หมวด{expense group}" → per month "งบประมาณ ตาม BG {plan} เดือน {month} {พ.ศ.}"
 *     → "- item qty unit amount บาท", STORE balance, หมายเหตุ
 *   then the 14-column table and "จึงเรียนมาเพื่อโปรดพิจารณา".
 */
function buildApprovalHtml_(ss, req, o, s, kind) {
  o = o || {};
  var d = approvalData_(ss, req);
  var items = d.items, bgTotal = d.bgTotal, total = d.total, months = d.months, be = d.be;
  var monthText = months.map(function (m) { return THAI_MONTH_FULL[m]; }).join(',') + ' ' + be;
  var team = s.team_name ? ' (' + esc_(s.team_name) + ')' : '';
  var approved = kind === 'purchasing';
  var P = 'style="margin:0 0 4px"';

  var h = '<div style="font-family:Tahoma,Arial,sans-serif;font-size:14px;line-height:1.55;color:#222">';
  h += '<p ' + P + '>เรียน&nbsp;&nbsp;&nbsp;&nbsp;' + esc_(o.greeting || (approved ? 'ฝ่ายจัดซื้อ' : s.approval_greeting)) + '</p>';
  h += '<p ' + P + '>สำเนา&nbsp;&nbsp;ผู้เกี่ยวข้อง</p>';
  h += '<p ' + P + '>เรื่อง&nbsp;&nbsp;&nbsp;&nbsp;' + (approved ? 'ได้รับการอนุมัติสั่งซื้ออุปกรณ์' : 'ขออนุมัติสั่งซื้ออุปกรณ์') + team + '</p><br>';
  h += '<p ' + P + '>' + (approved ? 'ได้รับการอนุมัติสั่งซื้ออุปกรณ์' : 'ขออนุมัติสั่งซื้ออุปกรณ์') + team +
    ' ใช้ Budget เดือน ' + esc_(monthText) + ' รายละเอียดตามแนบ</p>';
  h += '<p ' + P + '>งบประมาณตาม BG <b>' + fmtBaht_(bgTotal) + '</b> บาท ค่าใช้จ่ายโดยประมาณ <b>' + fmtBaht_(total) + '</b> บาท (ไม่รวม VAT)</p>';
  var overItems = items.filter(function (i) { return i.over; });
  if (overItems.length) {
    h += '<p ' + P + '><b style="color:#b23b2e">** มีรายการขอ Over Budget: ' +
      esc_(uniq_(overItems.map(function (i) { return i.line.label + ' เดือน ' + THAI_MONTH_FULL[i.month] + (i.overReason ? ' — เหตุผล: ' + i.overReason : ''); })).join(' · ')) +
      '</b></p>';
  }
  if (o.intro) h += '<p ' + P + '>' + esc_(o.intro).replace(/\n/g, '<br>') + '</p>';
  if (req.note) h += '<p ' + P + '>หมายเหตุผู้ขอ: ' + esc_(req.note) + '</p>';
  if (o.link) h += '<p ' + P + '>โปรดคลิกลิ้งค์ : <a href="' + esc_(o.link) + '">' + esc_(o.link) + '</a></p>';
  h += '<br>';

  // sections: สื่อ × หมวด, then month
  var groups = [], gBy = {};
  items.forEach(function (i) {
    var media = i.line.location || i.line.mediaGroup || i.line.mediaType || i.line.label;
    var cat = i.line.expenseGroup || i.line.glName || '';
    var gk = media + '|' + cat;
    if (!gBy[gk]) { gBy[gk] = { media: media, cat: cat, months: [], mBy: {} }; groups.push(gBy[gk]); }
    var g = gBy[gk], mk = i.key + '|' + i.month;
    if (!g.mBy[mk]) { g.mBy[mk] = { month: i.month, plan: i.plan, over: i.over, reason: i.overReason, items: [] }; g.months.push(g.mBy[mk]); }
    g.mBy[mk].items.push(i);
  });
  groups.forEach(function (g) {
    h += '<p style="margin:10px 0 2px"><b><u>สื่อ ' + esc_(g.media) + ' หมวด' + esc_(g.cat) + '</u></b></p>';
    g.months.sort(function (a, b) { return a.month - b.month; }).forEach(function (m) {
      h += '<p ' + P + '>งบประมาณ&nbsp;&nbsp;ตาม BG ' + fmtBaht_(m.plan) + ' เดือน ' + THAI_MONTH_FULL[m.month] + ' ' + be +
        (m.over ? ' <b style="color:#b23b2e">(ขอ Over Budget' + (m.reason ? ': ' + esc_(m.reason) : '') + ')</b>' : '') + '</p>';
      m.items.forEach(function (i) {
        h += '<p style="margin:0 0 2px 16px">- ' + esc_(i.name) + ' ' + esc_(i.qty) + ' ' + esc_(i.unit) + ' ' + fmtBaht_(i.amount) + ' บาท</p>';
        h += '<p style="margin:0 0 2px 16px">ตรวจสอบจาก STORE ยอดคงเหลือ ' + (i.balance === '' ? '-' : esc_(i.balance)) + '</p>';
        if (i.note) h += '<p style="margin:0 0 6px 16px">หมายเหตุ :&nbsp;&nbsp;' + esc_(i.note) + '</p>';
      });
    });
  });

  // 14-column table
  var th = 'style="background:#1f4e79;color:#fff;border:1px solid #999;padding:4px 6px;font-size:12px"';
  var td = 'style="border:1px solid #999;padding:4px 6px;font-size:12px;vertical-align:top"';
  var tdr = 'style="border:1px solid #999;padding:4px 6px;font-size:12px;vertical-align:top;text-align:right"';
  var cols = ['Company/นามบริษัท', 'Media Type', 'Sub Media Type', 'Expense Group/Part Code Detail.', 'Remark ตาม Budget',
    'รายการสั่งซื้ออุปกรณ์', 'จำนวนที่สั่งซื้อ', 'จำนวนคงเหลือในStore', 'ชื่ออุปกรณ์/Store', 'ขอบเขตการใช้งาน',
    'ค่าใช้จ่ายตาม Budget รายปี', 'ค่าใช้จ่ายโดยประมาณ', 'เดือนสั่งอุปกรณ์', 'สถานะการขออนุมัติ'];
  h += '<br><table cellspacing="0" cellpadding="0" style="border-collapse:collapse">' +
    '<tr>' + cols.map(function (c) { return '<th ' + th + '>' + esc_(c) + '</th>'; }).join('') + '</tr>';
  items.forEach(function (i) {
    var status = approved ? 'อนุมัติ' : (i.over ? 'ส่งขออนุมัติ (Over Budget)' : 'ส่งขออนุมัติ');
    h += '<tr>' + [
      [td, i.line.company], [td, i.line.mediaType || i.line.mediaGroup], [td, i.line.location || i.line.mediaGroup],
      [td, i.line.expenseGroup || i.line.glName], [td, i.remark], [td, i.name], [tdr, i.qty + (i.unit ? ' ' + i.unit : '')],
      [tdr, i.balance === '' ? '-' : i.balance], [td, i.sku + ' ' + i.name], [td, i.note],
      [tdr, fmtBaht_(i.plan)], [tdr, fmtBaht_(i.amount)], [td, THAI_MONTH_FULL[i.month]], [td, status]
    ].map(function (c) { return '<td ' + c[0] + '>' + esc_(c[1] == null ? '' : c[1]) + '</td>'; }).join('') + '</tr>';
  });
  h += '<tr><td ' + td + ' colspan="10"><b>รวม</b></td><td ' + tdr + '><b>' + fmtBaht_(bgTotal) + '</b></td><td ' + tdr + '><b>' +
    fmtBaht_(total) + '</b></td><td ' + td + ' colspan="2"></td></tr></table>';

  h += '<br><p ' + P + '>' + (approved ? 'จึงเรียนมาเพื่อโปรดดำเนินการ' : 'จึงเรียนมาเพื่อโปรดพิจารณา') + '</p>';
  if (s.signature) h += '<br><p ' + P + '>' + esc_(s.signature).replace(/\n/g, '<br>') + '</p>';
  h += '<p style="margin:12px 0 0;color:#888;font-size:11px">เลขที่คำขอ ' + esc_(req.request_id) + ' · ผู้ขอ ' +
    esc_(req.requester_name || req.requester_email) + '</p></div>';
  return h;
}

/** 140500 → "140,500"; keeps satang only when present. */
function fmtBaht_(n) {
  n = Math.round((Number(n) || 0) * 100) / 100;
  return n.toLocaleString('en-US', { minimumFractionDigits: n % 1 ? 2 : 0, maximumFractionDigits: 2 });
}
