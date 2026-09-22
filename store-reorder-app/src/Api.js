/**
 * Api.js
 * Single entry point for the web page.
 *   - Apps Script page:       google.script.run.api(token, fn, args)
 *   - GitHub/Vercel page:     POST <exec url>  body: {"token","fn","args"}  (Content-Type text/plain,
 *                             so the browser sends no CORS preflight; Apps Script answers with
 *                             Access-Control-Allow-Origin: *)
 * Only functions listed below are reachable; everything except login/register
 * needs a valid session token.
 * Requires the Web App deployment "Who has access: Anyone" so the static page can reach it.
 */

// Names, not references: with separate .gs files the other files may not be loaded
// yet when this file's top level runs, so resolve at call time.
var PUBLIC_FNS = ['login', 'register'];
var MEMBER_FNS = [
  'getSessionInfo', 'logout', 'changePassword', 'updateProfile',
  // shared Store data
  'getStoreData', 'saveDataFile', 'removeDataFile',
  // budget + requests (workflow)
  'getBudgetOptions', 'importBudget', 'importBudgetFromDrive', 'submitOrderRequest', 'createRound', 'confirmRequest', 'rejectByPc',
  'cancelMyRequest', 'decideRequest', 'requestOverBudget', 'managerDecision', 'issuePo', 'markReceived',
  'getApprovalDraft', 'sendApprovalEmail', 'recordApproval', 'sendPurchasingEmail', 'getAppSettings', 'saveAppSettings',
  'listMyRequests', 'listAllRequests', 'installReminderTrigger',
  // admin
  'listMembers', 'saveMember', 'listMasterPc', 'saveMasterPc',
  // AI
  'askClaudeFromUi'
];

function api(token, fn, args) {
  fn = String(fn);
  args = args || [];
  if (PUBLIC_FNS.indexOf(fn) !== -1) return globalThis[fn].apply(null, args);
  if (MEMBER_FNS.indexOf(fn) === -1) throw new Error('ไม่รู้จักคำสั่ง ' + fn);
  var user = userFromToken_(token);
  if (!user) throw new Error('SESSION_EXPIRED: กรุณาเข้าสู่ระบบใหม่');
  user.token = token;
  CURRENT_USER_ = user;
  try {
    return toJsonSafe_(globalThis[fn].apply(null, args));
  } finally {
    CURRENT_USER_ = null;
  }
}

/** HTTP transport for the GitHub/Vercel page. Always answers 200 with {ok, result|error}. */
function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    out = { ok: true, result: toJsonSafe_(api(body.token, body.fn, body.args)) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Dates → 'yyyy-MM-dd HH:mm' strings (google.script.run can't return Date objects either). */
function toJsonSafe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
  if (Array.isArray(v)) return v.map(toJsonSafe_);
  if (v && typeof v === 'object') {
    var o = {};
    Object.keys(v).forEach(function (k) { o[k] = toJsonSafe_(v[k]); });
    return o;
  }
  return v === undefined ? null : v;
}
