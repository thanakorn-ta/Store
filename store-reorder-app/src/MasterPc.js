/**
 * MasterPc.js
 * "Master: PC → อีเมลผู้รับผิดชอบ" in the workflow diagram. One row per PC:
 * who confirms the suggested order (owner_emails, comma separated) and who
 * approves when it goes over budget (manager_email). default_budget_key
 * pre-selects the budget line when the PC owner confirms.
 * Edited in the Admin tab (or directly in the master_pc sheet).
 */

var MASTER_PC_HEADER = ['pc_code', 'pc_name', 'owner_emails', 'manager_email', 'default_budget_key', 'note'];

function readMasterPc_(ss) {
  var out = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.MASTER_PC).forEach(function (p) {
    if (!p.pc_code) return;
    out[String(p.pc_code)] = {
      pc_code: String(p.pc_code), pc_name: String(p.pc_name || ''),
      owner_emails: splitEmails_(p.owner_emails), manager_email: String(p.manager_email || '').trim().toLowerCase(),
      default_budget_key: String(p.default_budget_key || ''), note: String(p.note || '')
    };
  });
  return out;
}

function splitEmails_(s) {
  return String(s || '').split(/[,;\s]+/).map(function (e) { return e.trim().toLowerCase(); }).filter(function (e) { return /@/.test(e); });
}

function listMasterPc() {
  requireActive_();
  var map = readMasterPc_(db_());
  return Object.keys(map).sort().map(function (k) {
    var p = map[k];
    return { pc_code: p.pc_code, pc_name: p.pc_name, owner_emails: p.owner_emails.join(', '),
      manager_email: p.manager_email, default_budget_key: p.default_budget_key, note: p.note };
  });
}

/** Upsert one PC row (admin). */
function saveMasterPc(input) {
  var me = requireAdmin_();
  var code = String(input.pc_code || '').trim();
  if (!code) throw new Error('ไม่มีรหัส PC');
  var owners = splitEmails_(input.owner_emails);
  var manager = String(input.manager_email || '').trim().toLowerCase();
  if (manager && !/@/.test(manager)) throw new Error('อีเมลหัวหน้าไม่ถูกต้อง');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = db_();
    var h = ensureHeader_(ss, CONFIG.SHEETS.MASTER_PC, MASTER_PC_HEADER);
    var values = h.sheet.getDataRange().getValues();
    var col = {};
    h.header.forEach(function (k, i) { col[k] = i; });
    var obj = { pc_code: code, pc_name: String(input.pc_name || ''), owner_emails: owners.join(', '),
      manager_email: manager, default_budget_key: String(input.default_budget_key || ''), note: String(input.note || '') };
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][col.pc_code]) === code) {
        MASTER_PC_HEADER.forEach(function (k) { h.sheet.getRange(r + 1, col[k] + 1).setValue(obj[k]); });
        logActivity_('saveMasterPc', 'ok', me.email + ' ' + code);
        return listMasterPc();
      }
    }
    appendObject_(ss, CONFIG.SHEETS.MASTER_PC, MASTER_PC_HEADER, obj);
  } finally {
    lock.releaseLock();
  }
  logActivity_('saveMasterPc', 'ok', me.email + ' ' + code);
  return listMasterPc();
}
