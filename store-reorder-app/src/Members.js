/**
 * Members.js
 * Membership for the Web App. Identity comes from the Google Workspace login
 * (Session.getActiveUser) — no separate passwords. This only works when the
 * Web App is deployed with "Who has access: Anyone within planbmedia.co.th";
 * with "Anyone" Google does not reveal the viewer's email.
 *
 * members sheet: email | name | role (admin/user) | status (active/pending/disabled)
 *                | created_at | updated_at | updated_by
 * CONFIG.ADMIN_EMAILS are auto-created as active admins on first visit.
 */

var MEMBER_HEADER = ['email', 'name', 'role', 'status', 'created_at', 'updated_at', 'updated_by'];

function db_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/** Returns the sheet, creating it with a header row if missing. */
function ensureSheet_(ss, name, header) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Like ensureSheet_, but also appends any header columns that an older
 * version of the sheet is missing, so new fields never shift existing data.
 * Returns { sheet, header } where header is the sheet's actual column order.
 */
function ensureHeader_(ss, name, header) {
  var sheet = ensureSheet_(ss, name, header);
  var lastCol = Math.max(1, sheet.getLastColumn ? sheet.getLastColumn() : header.length);
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String).filter(String);
  var missing = header.filter(function (h) { return current.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    current = current.concat(missing);
  }
  return { sheet: sheet, header: current };
}

/** Appends one object as a row, placing each field in its named column. */
function appendObject_(ss, name, header, obj) {
  var h = ensureHeader_(ss, name, header);
  h.sheet.appendRow(h.header.map(function (k) { return obj[k] == null ? '' : obj[k]; }));
}

/** Appends many objects in one write. */
function appendObjects_(ss, name, header, objs) {
  if (!objs.length) return;
  var h = ensureHeader_(ss, name, header);
  var rows = objs.map(function (o) { return h.header.map(function (k) { return o[k] == null ? '' : o[k]; }); });
  h.sheet.getRange(h.sheet.getLastRow() + 1, 1, rows.length, h.header.length).setValues(rows);
}

function activeEmail_() {
  var email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    throw new Error('ระบบอ่านอีเมลผู้ใช้ไม่ได้ — Deploy Web App ต้องตั้ง "ผู้มีสิทธิ์เข้าถึง: ทุกคนใน planbmedia.co.th"');
  }
  return email;
}

/** All member rows as objects with their 1-based sheet row number. */
function readMembers_(ss) {
  var sheet = ensureSheet_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({
      row: i + 1,
      email: String(values[i][0]).trim().toLowerCase(),
      name: String(values[i][1] || ''),
      role: String(values[i][2] || 'user'),
      status: String(values[i][3] || 'pending')
    });
  }
  return out;
}

/** Current viewer's member record, or {status:'none'} if not registered. */
function currentMember_() {
  var email = activeEmail_();
  var ss = db_();
  var m = readMembers_(ss).filter(function (x) { return x.email === email; })[0];
  if (m) return m;

  var bootstrap = CONFIG.ADMIN_EMAILS.map(function (e) { return e.toLowerCase(); }).indexOf(email) !== -1;
  if (bootstrap) {
    var now = new Date();
    var sheet = ensureSheet_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER);
    sheet.appendRow([email, email.split('@')[0], 'admin', 'active', now, now, 'bootstrap']);
    return { email: email, name: email.split('@')[0], role: 'admin', status: 'active', row: sheet.getLastRow() };
  }
  return { email: email, name: '', role: '', status: 'none' };
}

function requireActive_() {
  var m = currentMember_();
  if (m.status !== 'active') throw new Error('บัญชีนี้ยังไม่ได้รับอนุมัติให้ใช้งาน');
  return m;
}

function requireAdmin_() {
  var m = requireActive_();
  if (m.role !== 'admin') throw new Error('เฉพาะ Admin เท่านั้น');
  return m;
}

function adminEmails_(ss) {
  return readMembers_(ss || db_())
    .filter(function (m) { return m.role === 'admin' && m.status === 'active'; })
    .map(function (m) { return m.email; });
}

function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

// ------------------------------------------------------------ called from the page

function getSessionInfo() {
  var m = currentMember_();
  return { email: m.email, name: m.name, role: m.role, status: m.status, appUrl: webAppUrl_() };
}

/** A signed-in but unregistered viewer asks to join; admins get an email. */
function requestAccess(name) {
  var email = activeEmail_();
  name = String(name || '').trim().slice(0, 80);
  if (!name) throw new Error('กรุณากรอกชื่อ');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = db_();
    if (readMembers_(ss).some(function (m) { return m.email === email; })) return getSessionInfo();
    var now = new Date();
    ensureSheet_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER).appendRow([email, name, 'user', 'pending', now, now, email]);
    var admins = adminEmails_(ss);
    if (admins.length) {
      MailApp.sendEmail({
        to: admins.join(','),
        subject: '[Store Reorder] ขอเข้าใช้งาน: ' + name,
        htmlBody: esc_(name) + ' (' + esc_(email) + ') ขอเข้าใช้งานระบบ<br>' +
          'อนุมัติได้ที่แท็บ Admin: <a href="' + webAppUrl_() + '">' + webAppUrl_() + '</a>'
      });
    }
  } finally {
    lock.releaseLock();
  }
  return getSessionInfo();
}

function listMembers() {
  requireAdmin_();
  return readMembers_(db_()).map(function (m) {
    return { email: m.email, name: m.name, role: m.role, status: m.status };
  });
}

/** Add or update a member. Admins cannot demote/disable themselves (avoids lock-out). */
function saveMember(input) {
  var me = requireAdmin_();
  var email = String(input.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  var role = input.role === 'admin' ? 'admin' : 'user';
  var status = ['active', 'pending', 'disabled'].indexOf(input.status) !== -1 ? input.status : 'active';
  if (email === me.email && (role !== 'admin' || status !== 'active')) {
    throw new Error('เปลี่ยนสิทธิ์/ระงับบัญชีตัวเองไม่ได้');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = db_();
    var sheet = ensureSheet_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER);
    var now = new Date();
    var existing = readMembers_(ss).filter(function (m) { return m.email === email; })[0];
    var name = String(input.name || (existing && existing.name) || email.split('@')[0]).slice(0, 80);
    if (existing) {
      sheet.getRange(existing.row, 2, 1, 3).setValues([[name, role, status]]);
      sheet.getRange(existing.row, 6, 1, 2).setValues([[now, me.email]]);
      if (existing.status === 'pending' && status === 'active') {
        MailApp.sendEmail(email, '[Store Reorder] อนุมัติการเข้าใช้งานแล้ว',
          'บัญชีของคุณได้รับอนุมัติแล้ว เข้าใช้งานได้ที่ ' + webAppUrl_());
      }
    } else {
      sheet.appendRow([email, name, role, status, now, now, me.email]);
    }
  } finally {
    lock.releaseLock();
  }
  logActivity_('saveMember', 'ok', me.email + ' -> ' + email + ' ' + role + '/' + status);
  return listMembers();
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
