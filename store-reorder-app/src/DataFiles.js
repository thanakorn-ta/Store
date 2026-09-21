/**
 * DataFiles.js
 * Shared Store data behind the 4 file slots in the web UI. Instead of each
 * browser keeping its own copy, parsed rows live in the database Sheet so
 * every member sees the same data, and anyone with edit access to the Sheet
 * can fix individual rows directly (then press "โหลดใหม่" in the app).
 *
 *   data_minmax | data_usage | data_balance | data_reorder
 *     column A = source_file (which uploaded file the row came from),
 *     remaining columns = the record fields produced by web/parsers.js
 *   data_files  = upload log: slot | file_name | rows | uploaded_by | uploaded_at
 *
 * Admins add/replace/remove files from the app; users read only.
 * "usage" keeps several files (one per month); the other slots hold one file.
 */

var DATA_SLOTS = {
  minmax:  { sheet: 'data_minmax',  multi: false, fields: ['sku', 'name', 'group', 'avg', 'unit', 'lead', 'safety', 'min', 'max', 'category'] },
  usage:   { sheet: 'data_usage',   multi: true,  fields: ['sku', 'name', 'group', 'month', 'year', 'qty', 'unit', 'price', 'total'] },
  balance: { sheet: 'data_balance', multi: false, fields: ['sku', 'name', 'pc', 'pcName', 'location', 'balance', 'unit', 'cost', 'value'] },
  reorder: { sheet: 'data_reorder', multi: false, fields: ['sku', 'name', 'pc', 'pcName', 'balance', 'min', 'need', 'unit', 'lot'] }
};
// Fields read back as numbers; minmax min/max stay null when blank (engine treats null as "not set")
var DATA_NUMERIC = ['avg', 'lead', 'safety', 'min', 'max', 'month', 'year', 'qty', 'price', 'total', 'balance', 'cost', 'value', 'need'];
var DATA_FILES_HEADER = ['slot', 'file_name', 'rows', 'uploaded_by', 'uploaded_at'];

function dataSheet_(ss, slot) {
  var def = DATA_SLOTS[slot];
  if (!def) throw new Error('ไม่รู้จักช่องข้อมูล ' + slot);
  return ensureSheet_(ss, def.sheet, ['source_file'].concat(def.fields));
}

/** Everything the page needs to fill the 4 slots. */
function getStoreData() {
  requireActive_();
  var ss = db_();
  var meta = {};
  readSheetAsObjects_(ss, 'data_files').forEach(function (m) {
    meta[m.slot + '|' + m.file_name] = {
      uploadedBy: String(m.uploaded_by || ''),
      uploadedAt: m.uploaded_at instanceof Date ? Utilities.formatDate(m.uploaded_at, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : String(m.uploaded_at || '')
    };
  });

  var out = {};
  Object.keys(DATA_SLOTS).forEach(function (slot) {
    var def = DATA_SLOTS[slot];
    var files = {};
    var order = [];
    readSheetAsObjects_(ss, def.sheet).forEach(function (row) {
      if (!row.sku) return;
      var src = String(row.source_file || '(แก้ในชีต)');
      if (!files[src]) { files[src] = []; order.push(src); }
      var rec = {};
      def.fields.forEach(function (f) {
        var v = row[f];
        if (DATA_NUMERIC.indexOf(f) !== -1) {
          rec[f] = (v === '' || v == null) ? ((f === 'min' || f === 'max') && slot === 'minmax' ? null : 0) : Number(v) || 0;
        } else {
          rec[f] = v instanceof Date ? Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd') : String(v == null ? '' : v);
        }
      });
      files[src].push(rec);
    });
    out[slot] = order.map(function (name) {
      var m = meta[slot + '|' + name] || {};
      return { name: name, count: files[name].length, records: files[name], uploadedBy: m.uploadedBy || '', uploadedAt: m.uploadedAt || '' };
    });
  });
  return { files: out, sheetUrl: isAdminEmail_() ? ss.getUrl() : '' };
}

function isAdminEmail_() {
  try { requireAdmin_(); return true; } catch (e) { return false; }
}

/** Add or replace one uploaded file's rows. Single-file slots are cleared first. */
function saveDataFile(slot, fileName, records) {
  var me = requireAdmin_();
  var def = DATA_SLOTS[slot];
  if (!def) throw new Error('ไม่รู้จักช่องข้อมูล ' + slot);
  fileName = String(fileName || '').slice(0, 200);
  if (!fileName) throw new Error('ไม่มีชื่อไฟล์');
  if (!records || !records.length) throw new Error('ไฟล์นี้ไม่มีข้อมูล');
  if (records.length > 50000) throw new Error('ข้อมูลมากเกินไป');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = db_();
    var removed = def.multi ? [fileName] : null; // null = replace everything in this slot
    rewriteDataSheet_(ss, slot, function (src) { return removed ? removed.indexOf(src) === -1 : false; }, records.map(function (r) {
      return [fileName].concat(def.fields.map(function (f) { return r[f] == null ? '' : r[f]; }));
    }));
    updateDataIndex_(ss, slot, def.multi ? fileName : null, { file_name: fileName, rows: records.length, uploaded_by: me.email, uploaded_at: new Date() });
  } finally {
    lock.releaseLock();
  }
  logActivity_('saveDataFile', 'ok', me.email + ' ' + slot + ' ' + fileName + ' (' + records.length + ' rows)');
  return getStoreData();
}

function removeDataFile(slot, fileName) {
  var me = requireAdmin_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = db_();
    rewriteDataSheet_(ss, slot, function (src) { return src !== fileName; }, []);
    updateDataIndex_(ss, slot, fileName, null);
  } finally {
    lock.releaseLock();
  }
  logActivity_('removeDataFile', 'ok', me.email + ' ' + slot + ' ' + fileName);
  return getStoreData();
}

/** Keeps rows whose source_file passes keepFn, then appends newRows. */
function rewriteDataSheet_(ss, slot, keepFn, newRows) {
  var sheet = dataSheet_(ss, slot);
  var values = sheet.getDataRange().getValues();
  var header = ['source_file'].concat(DATA_SLOTS[slot].fields);
  var kept = values.slice(1).filter(function (r) { return r[0] !== '' && keepFn(String(r[0])); })
    .map(function (r) { return header.map(function (_, i) { return r[i] == null ? '' : r[i]; }); });
  var all = kept.concat(newRows);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (all.length) sheet.getRange(2, 1, all.length, header.length).setValues(all);
}

/** Drops index rows for (slot, fileName) — or the whole slot when fileName is null — then adds entry. */
function updateDataIndex_(ss, slot, fileName, entry) {
  var sheet = ensureSheet_(ss, 'data_files', DATA_FILES_HEADER);
  var rows = sheet.getDataRange().getValues().slice(1).filter(function (r) {
    return r[0] && !(r[0] === slot && (fileName === null || r[1] === fileName));
  });
  if (entry) rows.push([slot, entry.file_name, entry.rows, entry.uploaded_by, entry.uploaded_at]);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, DATA_FILES_HEADER.length).setValues([DATA_FILES_HEADER]).setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, DATA_FILES_HEADER.length).setValues(rows);
}
