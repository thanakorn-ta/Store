/**
 * Members.js
 * Username/password membership (Admin / User) for the Web App and for the
 * GitHub/Vercel page that calls this script as its backend (see Api.js).
 *
 * members sheet: username | name | email | role (admin/user) | status (active/pending/disabled)
 *                | pw_hash | pw_salt | must_change | created_at | updated_at | updated_by | last_login
 *   - passwords are stored salted + hashed (SHA-256, iterated), never in plain text
 *   - email is used for notifications and to match Master PC owners/managers
 *
 * First run: when no admin exists, CONFIG.DEFAULT_ADMIN is created with
 * must_change = true, and the page asks for a new password after login.
 *
 * Sessions: login() returns a random token kept in CacheService for
 * SESSION_SECONDS (sliding). Every page call goes through api(token, fn, args),
 * which sets CURRENT_USER_ before running the requested function.
 */

var MEMBER_HEADER = ['username', 'name', 'email', 'role', 'status', 'pw_hash', 'pw_salt', 'must_change',
  'created_at', 'updated_at', 'updated_by', 'last_login'];
var SESSION_SECONDS = 6 * 3600;
var MAX_FAILED_LOGINS = 5;
var CURRENT_USER_ = null;

// ------------------------------------------------------------ password + session primitives

function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  for (var i = 0; i < 300; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes.concat(Utilities.newBlob(salt).getBytes()));
  }
  return Utilities.base64Encode(bytes);
}

function newSalt_() {
  return Utilities.getUuid().replace(/-/g, '');
}

function validUsername_(u) {
  return /^[a-z0-9._-]{3,40}$/.test(u);
}

function checkNewPassword_(p) {
  if (String(p || '').length < 8) throw new Error('รหัสผ่านต้องยาวอย่างน้อย 8 ตัวอักษร');
}

function readMembers_(ss) {
  var h = ensureHeader_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER);
  var values = h.sheet.getDataRange().getValues();
  var col = {};
  h.header.forEach(function (k, i) { col[k] = i; });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var username = String(row[col.username] || '').trim().toLowerCase();
    if (!username) continue;
    out.push({
      row: i + 1, username: username, name: String(row[col.name] || ''),
      email: String(row[col.email] || '').trim().toLowerCase(), role: String(row[col.role] || 'user'),
      status: String(row[col.status] || 'pending'), pw_hash: String(row[col.pw_hash] || ''),
      pw_salt: String(row[col.pw_salt] || ''), must_change: row[col.must_change] === true || row[col.must_change] === 'TRUE'
    });
  }
  return out;
}

function writeMemberFields_(ss, rowNum, fields) {
  var h = ensureHeader_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER);
  Object.keys(fields).forEach(function (k) {
    var c = h.header.indexOf(k);
    if (c !== -1) h.sheet.getRange(rowNum, c + 1).setValue(fields[k]);
  });
}

/** Creates the default admin when the sheet has no admin at all. */
function ensureDefaultAdmin_(ss) {
  var members = readMembers_(ss);
  if (members.some(function (m) { return m.role === 'admin' && m.status === 'active'; })) return members;
  var d = CONFIG.DEFAULT_ADMIN;
  var salt = newSalt_();
  var now = new Date();
  appendObject_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER, {
    username: d.username, name: 'Admin', email: (CONFIG.ADMIN_EMAILS || [])[0] || '', role: 'admin', status: 'active',
    pw_hash: hashPassword_(d.password, salt), pw_salt: salt, must_change: true,
    created_at: now, updated_at: now, updated_by: 'bootstrap'
  });
  return readMembers_(ss);
}

function toMe_(m) {
  // email doubles as the identity used across requests / Master PC; fall back to the username
  return { id: m.username, name: m.name || m.username, email: m.email || m.username, role: m.role, status: m.status,
    mustChange: m.must_change, hasEmail: !!m.email };
}

// ------------------------------------------------------------ who is calling

function currentMember_() {
  if (!CURRENT_USER_) throw new Error('กรุณาเข้าสู่ระบบ');
  return CURRENT_USER_;
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
    .filter(function (m) { return m.role === 'admin' && m.status === 'active' && /@/.test(m.email); })
    .map(function (m) { return m.email; });
}

/** Resolves a session token to the member (re-read from the sheet so role/status changes apply at once). */
function userFromToken_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var username = cache.get('sess:' + token);
  if (!username) return null;
  var m = readMembers_(db_()).filter(function (x) { return x.username === username; })[0];
  if (!m || m.status === 'disabled') { cache.remove('sess:' + token); return null; }
  cache.put('sess:' + token, username, SESSION_SECONDS); // sliding expiry
  return toMe_(m);
}

// ------------------------------------------------------------ public (no session needed; called via api())

/** Returns { token, user } or throws. Locks a username for 10 minutes after repeated failures. */
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  var cache = CacheService.getScriptCache();
  var failKey = 'fail:' + username;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILED_LOGINS) throw new Error('ใส่รหัสผิดหลายครั้ง — ลองใหม่ใน 10 นาที');

  var ss = db_();
  var m = ensureDefaultAdmin_(ss).filter(function (x) { return x.username === username; })[0];
  if (!m || !m.pw_hash || hashPassword_(String(password || ''), m.pw_salt) !== m.pw_hash) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('ID หรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(failKey);
  if (m.status === 'pending') throw new Error('บัญชีนี้รอ Admin อนุมัติ');
  if (m.status === 'disabled') throw new Error('บัญชีนี้ถูกระงับ ติดต่อ Admin');

  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  cache.put('sess:' + token, m.username, SESSION_SECONDS);
  writeMemberFields_(ss, m.row, { last_login: new Date() });
  logActivity_('login', 'ok', m.username);
  return { token: token, user: toMe_(m), appUrl: webAppUrl_() };
}

/** Self sign-up → pending until an admin approves. */
function register(input) {
  var username = String(input.username || '').trim().toLowerCase();
  var email = String(input.email || '').trim().toLowerCase();
  var name = String(input.name || '').trim().slice(0, 80);
  if (!validUsername_(username)) throw new Error('ID ใช้ได้เฉพาะ a-z 0-9 . _ - ยาว 3–40 ตัว');
  if (!name) throw new Error('กรุณากรอกชื่อ');
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  checkNewPassword_(input.password);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = db_();
    if (ensureDefaultAdmin_(ss).some(function (m) { return m.username === username; })) throw new Error('ID นี้มีคนใช้แล้ว');
    var salt = newSalt_(), now = new Date();
    appendObject_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER, {
      username: username, name: name, email: email, role: 'user', status: 'pending',
      pw_hash: hashPassword_(String(input.password), salt), pw_salt: salt, must_change: false,
      created_at: now, updated_at: now, updated_by: 'register'
    });
  } finally {
    lock.releaseLock();
  }
  var admins = adminEmails_();
  if (admins.length) {
    MailApp.sendEmail({ to: admins.join(','), subject: '[Store Reorder] สมัครสมาชิกใหม่: ' + name,
      htmlBody: esc_(name) + ' (ID: ' + esc_(username) + (email ? ', ' + esc_(email) : '') + ') สมัครสมาชิก — อนุมัติได้ที่แท็บ Admin' +
        (webAppUrl_() ? '<br><a href="' + webAppUrl_() + '">' + webAppUrl_() + '</a>' : '') });
  }
  logActivity_('register', 'ok', username);
  return 'สมัครแล้ว — รอ Admin อนุมัติ แล้วเข้าสู่ระบบด้วย ID นี้';
}

// ------------------------------------------------------------ signed-in user

function getSessionInfo() {
  var me = currentMember_();
  return { id: me.id, name: me.name, email: me.hasEmail ? me.email : '', role: me.role, status: me.status,
    mustChange: me.mustChange, appUrl: webAppUrl_() };
}

function logout() {
  if (CURRENT_USER_ && CURRENT_USER_.token) CacheService.getScriptCache().remove('sess:' + CURRENT_USER_.token);
  return true;
}

function changePassword(oldPassword, newPassword) {
  var me = requireActive_();
  checkNewPassword_(newPassword);
  if (String(newPassword) === CONFIG.DEFAULT_ADMIN.password) throw new Error('ห้ามใช้รหัสผ่านเริ่มต้น');
  var ss = db_();
  var m = readMembers_(ss).filter(function (x) { return x.username === me.id; })[0];
  if (hashPassword_(String(oldPassword || ''), m.pw_salt) !== m.pw_hash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  var salt = newSalt_();
  writeMemberFields_(ss, m.row, { pw_hash: hashPassword_(String(newPassword), salt), pw_salt: salt, must_change: false,
    updated_at: new Date(), updated_by: me.id });
  logActivity_('changePassword', 'ok', me.id);
  return getSessionInfoAfterChange_(me);
}

function getSessionInfoAfterChange_(me) {
  var info = getSessionInfo();
  info.mustChange = false;
  return info;
}

/** The signed-in user updates their own name/email. */
function updateProfile(input) {
  var me = requireActive_();
  var email = String(input.email || '').trim().toLowerCase();
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  var ss = db_();
  var m = readMembers_(ss).filter(function (x) { return x.username === me.id; })[0];
  writeMemberFields_(ss, m.row, { name: String(input.name || m.name).slice(0, 80), email: email, updated_at: new Date(), updated_by: me.id });
  return true;
}

// ------------------------------------------------------------ admin: members

function listMembers() {
  requireAdmin_();
  return ensureDefaultAdmin_(db_()).map(function (m) {
    return { username: m.username, name: m.name, email: m.email, role: m.role, status: m.status, mustChange: m.must_change };
  });
}

/**
 * Add or update a member. input: { username, name, email, role, status, password? }
 * password (when given) sets/resets it and forces a change at next login.
 * Admins cannot demote/disable themselves (avoids lock-out).
 */
function saveMember(input) {
  var me = requireAdmin_();
  var username = String(input.username || '').trim().toLowerCase();
  if (!validUsername_(username)) throw new Error('ID ใช้ได้เฉพาะ a-z 0-9 . _ - ยาว 3–40 ตัว');
  var email = String(input.email || '').trim().toLowerCase();
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  var role = input.role === 'admin' ? 'admin' : 'user';
  var status = ['active', 'pending', 'disabled'].indexOf(input.status) !== -1 ? input.status : 'active';
  if (username === me.id && (role !== 'admin' || status !== 'active')) throw new Error('เปลี่ยนสิทธิ์/ระงับบัญชีตัวเองไม่ได้');
  if (input.password) checkNewPassword_(input.password);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var approvedNow = false, existing;
  try {
    var ss = db_();
    existing = readMembers_(ss).filter(function (m) { return m.username === username; })[0];
    var now = new Date();
    var fields = { name: String(input.name || (existing && existing.name) || username).slice(0, 80), email: email,
      role: role, status: status, updated_at: now, updated_by: me.id };
    if (input.password) {
      var salt = newSalt_();
      fields.pw_hash = hashPassword_(String(input.password), salt);
      fields.pw_salt = salt;
      fields.must_change = true;
    }
    if (existing) {
      writeMemberFields_(ss, existing.row, fields);
      approvedNow = existing.status === 'pending' && status === 'active';
    } else {
      if (!input.password) throw new Error('สมาชิกใหม่ต้องตั้งรหัสผ่านเริ่มต้น');
      fields.username = username;
      fields.created_at = now;
      appendObject_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER, fields);
    }
  } finally {
    lock.releaseLock();
  }
  if (approvedNow && email) {
    MailApp.sendEmail(email, '[Store Reorder] อนุมัติการเข้าใช้งานแล้ว',
      'บัญชี ID ' + username + ' ได้รับอนุมัติแล้ว เข้าใช้งานได้ที่ ' + webAppUrl_());
  }
  logActivity_('saveMember', 'ok', me.id + ' -> ' + username + ' ' + role + '/' + status + (input.password ? ' +password' : ''));
  return listMembers();
}

// ------------------------------------------------------------ admin: passwords you can read

/**
 * Passwords are stored salted + hashed, so nobody - not even an admin - can read an existing one.
 * To hand a login out, the admin sets a new password here and the page shows it once:
 *   resetMemberPassword({ username, password?, forceChange? }) -> { username, name, email, role, password }
 * password  omitted -> the system makes an easy-to-read one
 * forceChange true  -> the user must change it at the next login (default: keep it, so the list stays valid)
 */
var PW_SAFE_ = 'abcdefghijkmnpqrstuvwxyz23456789'; // no l/o/0/1 - they get misread when passed around

function randomPassword_(len) {
  var n = len || 10, out = '';
  for (var i = 0; i < n; i++) out += PW_SAFE_.charAt(Math.floor(Math.random() * PW_SAFE_.length));
  return out.slice(0, 4) + '-' + out.slice(4); // store-1 like grouping reads better over the phone
}

function resetMemberPassword(input) {
  var me = requireAdmin_();
  input = input || {};
  var username = String(input.username || '').trim().toLowerCase();
  var password = input.password ? String(input.password) : randomPassword_(10);
  if (input.password) checkNewPassword_(password);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var m;
  try {
    var ss = db_();
    m = readMembers_(ss).filter(function (x) { return x.username === username; })[0];
    if (!m) throw new Error('ไม่พบสมาชิก ' + username);
    var salt = newSalt_();
    writeMemberFields_(ss, m.row, { pw_hash: hashPassword_(password, salt), pw_salt: salt,
      must_change: input.forceChange === true, updated_at: new Date(), updated_by: me.id });
  } finally {
    lock.releaseLock();
  }
  logActivity_('resetMemberPassword', 'ok', me.id + ' -> ' + username); // never log the password itself
  return { username: m.username, name: m.name, email: m.email, role: m.role, status: m.status,
    password: password, mustChange: input.forceChange === true };
}

/** Same, for several accounts at once: { usernames: [...], forceChange? } -> [ { username, password, ... } ] */
function resetMemberPasswords(input) {
  requireAdmin_();
  input = input || {};
  var names = input.usernames && input.usernames.length ? input.usernames
    : listMembers().map(function (m) { return m.username; });
  return names.map(function (u) {
    return resetMemberPassword({ username: u, forceChange: input.forceChange === true });
  });
}
