/**
 * Quotes.js
 * Quotations (ใบเสนอราคา) attached to a purchase request.
 *
 * The page sends each file as base64 inside the submit payload; files are kept
 * in a Drive folder owned by the script owner (Script Property QUOTE_FOLDER_ID,
 * created on first use) and listed in the `request_files` sheet. Nobody needs
 * Drive access: the page downloads through getQuoteFile(), which checks that
 * the viewer is an admin, the requester or an approver of that request.
 * Admin approval emails attach them (Approval.js).
 */

var FILE_HEADER = ['request_id', 'file_id', 'name', 'mime', 'size', 'uploaded_by', 'uploaded_at'];
var QUOTE_MAX_FILES = 5;
var QUOTE_MAX_BYTES = 5 * 1024 * 1024;        // per file
var QUOTE_MAX_TOTAL = 10 * 1024 * 1024;       // per submit
var QUOTE_EXT = /\.(pdf|png|jpe?g|gif|webp|xlsx?|docx?|csv)$/i;

/** Validates { name, type, data(base64 or data: URL) } uploads before anything is written. */
function checkUploads_(files) {
  files = files || [];
  if (files.length > QUOTE_MAX_FILES) throw new Error('แนบใบเสนอราคาได้สูงสุด ' + QUOTE_MAX_FILES + ' ไฟล์');
  var total = 0;
  files.forEach(function (f) {
    var name = String(f && f.name || '');
    if (!QUOTE_EXT.test(name)) throw new Error('ไฟล์ ' + name + ' ไม่รองรับ (PDF, รูป, Excel, Word)');
    var bytes = Math.floor(String(f.data || '').replace(/^data:[^,]*,/, '').length * 3 / 4);
    if (!bytes) throw new Error('ไฟล์ ' + name + ' ว่างเปล่า');
    if (bytes > QUOTE_MAX_BYTES) throw new Error('ไฟล์ ' + name + ' ใหญ่เกิน ' + (QUOTE_MAX_BYTES / 1048576) + ' MB');
    total += bytes;
  });
  if (total > QUOTE_MAX_TOTAL) throw new Error('ไฟล์แนบรวมเกิน ' + (QUOTE_MAX_TOTAL / 1048576) + ' MB');
}

function quoteFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('QUOTE_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* deleted: make a new one */ }
  }
  var folder = DriveApp.createFolder('Store Reorder - ใบเสนอราคา');
  props.setProperty('QUOTE_FOLDER_ID', folder.getId());
  return folder;
}

function saveQuoteFiles_(ss, id, files, me) {
  files = files || [];
  if (!files.length) return;
  var folder = quoteFolder_();
  var rows = files.map(function (f) {
    var name = String(f.name).replace(/[\\/:*?"<>|]/g, '_').slice(0, 150);
    var bytes = Utilities.base64Decode(String(f.data).replace(/^data:[^,]*,/, ''));
    var file = folder.createFile(Utilities.newBlob(bytes, String(f.type || 'application/octet-stream'), id + ' - ' + name));
    return { request_id: id, file_id: file.getId(), name: name, mime: String(f.type || ''), size: bytes.length,
      uploaded_by: me.email, uploaded_at: new Date() };
  });
  appendObjects_(ss, CONFIG.SHEETS.REQUEST_FILES, FILE_HEADER, rows);
}

/** { request_id: [{ id, name, mime, size }] } */
function quoteFilesByRequest_(ss) {
  var out = {};
  readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_FILES).forEach(function (f) {
    if (!f.request_id || !f.file_id) return;
    (out[f.request_id] = out[f.request_id] || []).push({ id: String(f.file_id), name: String(f.name), mime: String(f.mime || ''), size: Number(f.size) || 0 });
  });
  return out;
}

/** Drops this request's files that are not in keepIds (moved to Drive trash, so still recoverable). */
function pruneQuoteFiles_(ss, id, keepIds) {
  keepIds = (keepIds || []).map(String);
  var drop = (quoteFilesByRequest_(ss)[id] || []).filter(function (f) { return keepIds.indexOf(f.id) === -1; });
  if (!drop.length) return;
  var ids = drop.map(function (f) { return f.id; });
  rewriteSheetRows_(ss, CONFIG.SHEETS.REQUEST_FILES, FILE_HEADER, function (row, col) {
    return !(row[col.request_id] === id && ids.indexOf(String(row[col.file_id])) !== -1);
  });
  ids.forEach(function (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) { /* already gone */ } });
}

/** Page download: { name, mime, data(base64) } for admins, the requester and approvers. */
function getQuoteFile(fileId) {
  var me = requireActive_();
  var ss = db_();
  var f = readSheetAsObjects_(ss, CONFIG.SHEETS.REQUEST_FILES).filter(function (x) { return String(x.file_id) === String(fileId); })[0];
  if (!f) throw new Error('ไม่พบไฟล์');
  var req = findRequest_(ss, f.request_id);
  var allowed = me.role === 'admin' || req.requester_email === me.email ||
    splitEmails_(req.approval_to).indexOf(me.email) !== -1 || splitEmails_(req.assigned_to).indexOf(me.email) !== -1;
  if (!allowed) throw new Error('ไม่มีสิทธิ์เปิดไฟล์นี้');
  var blob = DriveApp.getFileById(String(f.file_id)).getBlob();
  return { name: String(f.name), mime: String(f.mime || blob.getContentType() || ''), data: Utilities.base64Encode(blob.getBytes()) };
}

/** Blobs for email attachments; ids limited to this request's files. */
function quoteBlobs_(ss, id, ids) {
  var files = quoteFilesByRequest_(ss)[id] || [];
  if (ids) files = files.filter(function (f) { return ids.map(String).indexOf(f.id) !== -1; });
  return files.map(function (f) { return DriveApp.getFileById(f.id).getBlob().setName(f.name); });
}
