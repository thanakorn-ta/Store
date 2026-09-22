/**
 * Setup.js
 * Run setup() once from the Apps Script editor (Run ▶ setup) after pasting the
 * code. It asks for all permissions the Web App needs and checks the project,
 * printing ✓ / ✗ lines to the execution log so problems are easy to spot.
 * Safe to run again at any time.
 */

function setup() {
  ownerOnly_();
  var lines = [];
  var ok = function (msg) { lines.push('✓ ' + msg); };
  var bad = function (msg) { lines.push('✗ ' + msg); };

  // 1. Code from every module is present (a missing or mis-pasted file shows up here)
  var required = ['doGet', 'doPost', 'api', 'login', 'getStoreData', 'createRound', 'listMasterPc',
    'getBudgetOptions', 'callClaude_', 'readSheetAsObjects_', 'logActivity_', 'ensureHeader_'];
  var missing = required.filter(function (fn) { return typeof globalThis[fn] !== 'function'; });
  if (missing.length) bad('ไม่พบฟังก์ชัน: ' + missing.join(', ') + ' — วาง Code.gs ใหม่ทั้งไฟล์ (ดู README)');
  else ok('โค้ดครบทุกส่วน');
  // Old per-module .gs files left next to the bundle redefine functions and win (files load in order)
  if (typeof INDEX_HTML_ === 'string' && String(renderReorderUi_).indexOf('INDEX_HTML_') === -1) {
    bad('มีไฟล์ .gs เก่าทับโค้ดใน Code.gs — ลบไฟล์อื่นทั้งหมด ให้เหลือแค่ Code.gs (และ appsscript.json)');
  }

  // 2. The page itself
  try {
    renderReorderUi_();
    ok('หน้าเว็บพร้อม' + (typeof INDEX_HTML_ === 'string' ? ' (ฝังใน Code.gs)' : ' (ไฟล์ Index.html)'));
  } catch (e) {
    bad('หน้าเว็บ: ' + e.message + ' — ใช้ Code.gs จาก build/apps-script/ ซึ่งฝังหน้าเว็บไว้แล้ว');
  }

  // 3. Database sheet
  try {
    var ss = db_();
    ok('เปิดฐานข้อมูลได้: ' + ss.getName());
    var admins = ensureDefaultAdmin_(ss).filter(function (m) { return m.role === 'admin'; });
    ok('สมาชิก Admin: ' + admins.map(function (m) { return m.username; }).join(', ') +
      (admins.some(function (m) { return m.must_change; }) ? ' (ยังใช้รหัสเริ่มต้น — เปลี่ยนหลังเข้าสู่ระบบ)' : ''));
  } catch (e) {
    bad('ฐานข้อมูล: ' + e.message + ' — ตรวจ CONFIG.SPREADSHEET_ID และสิทธิ์เข้าถึง Sheet');
  }

  // 4. Services the app uses (running them here triggers the permission prompt)
  try { ok('ส่งอีเมลได้ (โควตาวันนี้เหลือ ' + MailApp.getRemainingDailyQuota() + ' ฉบับ)'); }
  catch (e) { bad('อีเมล: ' + e.message); }
  try {
    var c = CacheService.getScriptCache(); c.put('setup-check', '1', 60);
    if (c.get('setup-check') === '1') ok('CacheService ใช้ได้ (เก็บการเข้าสู่ระบบ)'); else bad('CacheService อ่านค่าไม่ได้');
  } catch (e) { bad('CacheService: ' + e.message); }
  try { ScriptApp.getProjectTriggers(); ok('ตั้ง trigger ได้ (ใช้กับการเตือนอัตโนมัติ)'); }
  catch (e) { bad('Trigger: ' + e.message); }
  ok(PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY')
    ? 'ตั้ง ANTHROPIC_API_KEY แล้ว (AI ช่วยร่างอีเมล)' : 'ยังไม่ตั้ง ANTHROPIC_API_KEY (ไม่บังคับ — อีเมลใช้แบบฟอร์มแทน)');

  // 5. Deployment
  var url = webAppUrl_();
  if (url) ok('ลิงก์ Web App: ' + url + ' — ตรวจว่า deployment รันโค้ดนี้: ' + url + '?health=1 ต้องได้ version ' + APP_VERSION + ' (ถ้าไม่ใช่ ให้ Deploy เวอร์ชันใหม่)');
  else bad('ยังไม่ได้ Deploy — การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป (Execute as: Me, Who has access: Anyone)');

  var report = lines.join('\n');
  Logger.log(report);
  return report;
}
