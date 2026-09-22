/**
 * BudgetImport.js
 * Server-side budget import, so the budget can be loaded without the browser:
 * finds the "Budget STT ... Revise-Budget" file in Google Drive (an .xlsx
 * upload or a Google Sheet), reads it, aggregates it exactly like
 * web/parsers.js parseBudget(), and writes budget_master.
 *
 *   - From the web app: Admin tab > "นำเข้าจาก Google Drive"  (admin only)
 *   - From the editor:  Run ▶ importBudgetFromDrive              (owner only)
 *
 * .xlsx files are converted with the Drive advanced service (enabled in
 * appsscript.json); the temporary Google Sheet copy is trashed afterwards.
 */

var BUDGET_FILE_QUERY = "name contains 'Budget STT' and trashed = false and " +
  "(mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.google-apps.spreadsheet')";

function importBudgetFromDrive(fileId) {
  if (CURRENT_USER_) requireAdmin_(); else ownerOnly_();
  var file = fileId ? { id: fileId, name: DriveApp.getFileById(fileId).getName() } : latestBudgetFile_();
  if (!file) throw new Error('ไม่พบไฟล์ Budget ใน Google Drive (ชื่อต้องมีคำว่า "Budget STT") — อัปโหลดไฟล์ .xlsx ขึ้น Drive ก่อน');

  var sheetId = file.id, tempId = null;
  if (file.mimeType !== 'application/vnd.google-apps.spreadsheet') {
    var copy = Drive.Files.copy({ name: '(temp import) ' + file.name, mimeType: 'application/vnd.google-apps.spreadsheet' }, file.id);
    sheetId = tempId = copy.id;
  }
  var values;
  try {
    values = SpreadsheetApp.openById(sheetId).getSheets()[0].getDataRange().getValues();
  } finally {
    if (tempId) DriveApp.getFileById(tempId).setTrashed(true); // only the temporary converted copy
  }

  var agg = aggregateBudgetValues_(values);
  if (!agg.rows.length) throw new Error('อ่านไฟล์ได้แต่ไม่พบรายการ Budget (ต้องมีคอลัมน์ Media Location / GL Code / Month Number)');
  writeBudgetMaster_(agg.rows);
  var msg = 'นำเข้า Budget จาก "' + file.name + '": ' + agg.rows.length + ' แถวรายเดือน' +
    (agg.skipped ? ' · ข้าม ' + agg.skipped + ' แถวที่ไม่มี GL Code/เลขเดือน' : '');
  logActivity_('importBudgetFromDrive', 'ok', msg);
  Logger.log(msg);
  return msg;
}

function latestBudgetFile_() {
  var res = Drive.Files.list({ q: BUDGET_FILE_QUERY, orderBy: 'modifiedTime desc', pageSize: 1, fields: 'files(id,name,mimeType,modifiedTime)' });
  var f = res.files && res.files[0];
  return f ? { id: f.id, name: f.name, mimeType: f.mimeType } : null;
}

/** Same rules as web/parsers.js parseBudget (keep the two in step). */
function aggregateBudgetValues_(values) {
  var h = -1;
  for (var r = 0; r < Math.min(values.length, 60); r++) {
    var cells = values[r].map(function (c) { return String(c).trim(); });
    if (cells.indexOf('Media Location') !== -1 && cells.indexOf('GL Code') !== -1) { h = r; break; }
  }
  if (h < 0) return { rows: [], skipped: 0 };
  var head = values[h].map(function (c) { return String(c).replace(/\s+/g, ' ').trim(); });
  var col = function (name) { return head.indexOf(name); };
  var c = {
    calc: col('Calculation'), company: col('Company'), division: col('Division'), mediaType: col('Media Type'),
    mediaGroup: col('Media Group'), location: col('Media Location'), expenseGroup: col('Expense Group'),
    gl: col('GL Code'), year: col('Year'), month: col('Month Number'), budget: col('Budget'),
    revise: col('Revise Budget'), actual: col('Actual'), remark: col('Remark'),
    glName: head.findIndex(function (x) { return /^ประเภทของค่าใช้จ่าย/.test(x); })
  };
  var s = function (row, i) { return i < 0 || row[i] == null ? '' : String(row[i]).replace(/\s+/g, ' ').trim(); };
  var n = function (row, i) { if (i < 0) return 0; var v = row[i]; if (typeof v === 'number') return v; var x = parseFloat(String(v || '').replace(/[^\d.\-]/g, '')); return isNaN(x) ? 0 : x; };

  var lines = {}, order = [], skipped = 0;
  for (var i = h + 1; i < values.length; i++) {
    var row = values[i];
    var gl = s(row, c.gl), month = n(row, c.month);
    var hasData = n(row, c.budget) || n(row, c.actual) || n(row, c.revise);
    var calcLoc = (s(row, c.calc).match(/\/\s*(.+?)\s+-\s+[^-]+$/) || [])[1] || '';
    var location = s(row, c.location) || calcLoc || s(row, c.mediaGroup);
    if (!gl || !location || !(month >= 1 && month <= 12)) { if (hasData) skipped++; continue; }
    var company = s(row, c.company);
    var key = company + '|' + location + '|' + gl;
    var k = key + '|' + month;
    if (!lines[k]) {
      lines[k] = { key: key, company: company, division: s(row, c.division), media_type: s(row, c.mediaType),
        media_group: s(row, c.mediaGroup), media_location: location, expense_group: s(row, c.expenseGroup),
        gl_code: gl, gl_name: s(row, c.glName), year: n(row, c.year), month_number: month,
        budget: 0, revise_budget: 0, actual: 0, plan: 0, remark: '', remarks: [] };
      order.push(k);
    }
    var l = lines[k];
    var reviseCell = s(row, c.revise);
    l.budget += n(row, c.budget);
    l.revise_budget += n(row, c.revise);
    l.actual += n(row, c.actual);
    l.plan += reviseCell !== '' ? n(row, c.revise) : n(row, c.budget);
    var remark = s(row, c.remark);
    if (remark && l.remarks.indexOf(remark) === -1) l.remarks.push(remark);
  }
  return {
    rows: order.map(function (k) {
      var l = lines[k];
      ['budget', 'revise_budget', 'actual', 'plan'].forEach(function (f) { l[f] = round2_(l[f]); });
      l.remark = l.remarks.join(' / ').slice(0, 300);
      delete l.remarks;
      return l;
    }),
    skipped: skipped
  };
}
