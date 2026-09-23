// Demo / trial mode: in-browser stand-ins for the Google services used by the server code
// (demo/server.js = src/*.js). Everything lives in this browser's localStorage — no real
// Google Sheet, Drive or email is touched. Also used by the local test harness.
(function () {
  const KEY = 'store-reorder-ai:demo-db';
  const SHEETS = {};
  const DRIVE = {};            // id -> { name, mime, data(base64), trashed, folder }
  const PROPS = {};
  const MAIL = [];             // sent emails, newest last
  const CACHE = {};
  window.__SHEETS = SHEETS; window.__MAIL = MAIL; window.__DRIVE = DRIVE;

  // ---------------------------------------------------------------- bytes / base64
  const enc = new TextEncoder();
  const toBytes = s => Array.from(enc.encode(String(s)));
  const b64 = bytes => {
    let s = '';
    const u = Uint8Array.from(bytes, x => x & 255);
    for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  };
  const unb64 = s => Array.from(atob(String(s)), c => c.charCodeAt(0));
  function Blob_(bytes, type, name) { this.bytes = bytes || []; this.type = type || 'application/octet-stream'; this.name = name || ''; }
  Blob_.prototype.getBytes = function () { return this.bytes.slice(); };
  Blob_.prototype.getContentType = function () { return this.type; };
  Blob_.prototype.getName = function () { return this.name; };
  Blob_.prototype.setName = function (n) { this.name = n; return this; };
  Blob_.prototype.getDataAsString = function () { return new TextDecoder().decode(Uint8Array.from(this.bytes, x => x & 255)); };

  // ---------------------------------------------------------------- Spreadsheet
  function Sheet(name) { this.name = name; this.data = []; }
  Sheet.prototype.getName = function () { return this.name; };
  Sheet.prototype.setName = function (n) { this.name = n; return this; };
  Sheet.prototype.getLastRow = function () { return this.data.length; };
  Sheet.prototype.getLastColumn = function () { return this.data.reduce((a, r) => Math.max(a, r.length), 0); };
  Sheet.prototype.appendRow = function (row) { this.data.push(row.slice()); return this; };
  Sheet.prototype.clearContents = function () { this.data = []; return this; };
  Sheet.prototype.clear = Sheet.prototype.clearContents;
  Sheet.prototype.setFrozenRows = function () { return this; };
  Sheet.prototype.autoResizeColumns = function () { return this; };
  Sheet.prototype.deleteRow = function (r) { this.data.splice(r - 1, 1); return this; };
  Sheet.prototype.getDataRange = function () { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); };
  Sheet.prototype.getRange = function (r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); };
  function Range(s, r, c, nr, nc) { Object.assign(this, { s, r, c, nr, nc }); }
  Range.prototype.getValues = function () {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.s.data[this.r - 1 + i] || [];
      const o = [];
      for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; o.push(v == null ? '' : v); }
      out.push(o);
    }
    return out;
  };
  Range.prototype.getValue = function () { return this.getValues()[0][0]; };
  Range.prototype.setValues = function (v) {
    for (let i = 0; i < v.length; i++) {
      const ri = this.r - 1 + i;
      while (this.s.data.length <= ri) this.s.data.push([]);
      for (let j = 0; j < v[i].length; j++) this.s.data[ri][this.c - 1 + j] = v[i][j];
    }
    return this;
  };
  Range.prototype.setValue = function (x) { return this.setValues([[x]]); };
  ['setFontWeight', 'setNumberFormat', 'setBackground', 'clearContent'].forEach(m => { Range.prototype[m] = function () { return this; }; });
  const SS = {
    getSheetByName: n => SHEETS[n] || null,
    insertSheet: n => (SHEETS[n] = new Sheet(n)),
    getSheets: () => Object.values(SHEETS),
    getUrl: () => '#demo', getId: () => 'demo'
  };
  window.SpreadsheetApp = {
    openById: () => SS, getActiveSpreadsheet: () => SS, flush() {},
    create: () => { throw new Error('โหมดทดลอง: ไม่สร้าง Google Sheet จริง'); } // → server falls back to CSV
  };

  // ---------------------------------------------------------------- Drive
  const newId = () => 'demo-' + Math.random().toString(36).slice(2, 12);
  const fileObj = id => ({
    getId: () => id, getName: () => DRIVE[id].name, getUrl: () => '#',
    getBlob: () => new Blob_(unb64(DRIVE[id].data), DRIVE[id].mime, DRIVE[id].name),
    setTrashed: t => { DRIVE[id].trashed = !!t; }
  });
  const folderObj = id => ({
    getId: () => id,
    createFile: blob => { const fid = newId(); DRIVE[fid] = { name: blob.getName(), mime: blob.getContentType(), data: b64(blob.getBytes()), folder: id }; return fileObj(fid); }
  });
  window.DriveApp = {
    createFolder: () => { const id = newId(); DRIVE[id] = { folder: true, name: 'folder' }; return folderObj(id); },
    getFolderById: id => { if (!DRIVE[id] || !DRIVE[id].folder) throw new Error('no folder'); return folderObj(id); },
    getFileById: id => { if (!DRIVE[id]) throw new Error('ไม่พบไฟล์ ' + id); return fileObj(id); },
    getFilesByName: () => ({ hasNext: () => false })
  };

  // ---------------------------------------------------------------- the rest
  window.LockService = { getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }) };
  window.CacheService = { getScriptCache: () => ({ get: k => CACHE[k] || null, put: (k, v) => { CACHE[k] = String(v); }, remove: k => { delete CACHE[k]; } }) };
  const pad = n => String(n).padStart(2, '0');
  window.Utilities = {
    formatDate: (d, tz, f) => f.replace('yyyy', d.getFullYear()).replace('yy', String(d.getFullYear()).slice(2)).replace('MM', pad(d.getMonth() + 1))
      .replace('dd', pad(d.getDate())).replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes())),
    getUuid: () => (crypto.randomUUID ? crypto.randomUUID() : newId() + newId()),
    DigestAlgorithm: { SHA_256: 'sha' }, Charset: { UTF_8: 'utf8' },
    // not SHA-256 — only needs to be deterministic for demo sign-in
    computeDigest: (a, v) => { const s = Array.isArray(v) ? v.join(',') : String(v); const out = []; let h = 2166136261; for (let i = 0; i < 32; i++) { for (let k = 0; k < s.length; k++) h = Math.imul(h ^ s.charCodeAt(k) ^ i, 16777619); out.push((h & 255) - 128); } return out; },
    newBlob: (data, type, name) => new Blob_(Array.isArray(data) ? data : toBytes(data == null ? '' : data), type, name),
    base64Encode: v => b64(Array.isArray(v) ? v : toBytes(v)),
    base64Decode: s => unb64(s),
    sleep: () => {}
  };
  window.MailApp = {
    sendEmail: o => {
      MAIL.push({ at: new Date().toISOString(), to: o.to, cc: o.cc || '', replyTo: o.replyTo || '', name: o.name || '', subject: o.subject, htmlBody: o.htmlBody || o.body || '',
        attachments: (o.attachments || []).map(b => ({ name: b.getName(), type: b.getContentType(), size: b.getBytes().length,
          data: b.getBytes().length < 400000 ? b64(b.getBytes()) : '' })) });
      while (MAIL.length > 30) MAIL.shift();
    },
    getRemainingDailyQuota: () => 100
  };
  window.PropertiesService = { getScriptProperties: () => ({ getProperty: k => PROPS[k] || null, setProperty: (k, v) => { PROPS[k] = String(v); }, deleteProperty: k => { delete PROPS[k]; } }) };
  window.Session = { getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => 'owner@demo' }) };
  window.ScriptApp = { getService: () => ({ getUrl: () => '' }), getProjectTriggers: () => [], getOAuthToken: () => 'demo', WeekDay: {},
    newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create() {} }) }) }) }), deleteTrigger() {} };
  window.UrlFetchApp = { fetch: () => { throw new Error('โหมดทดลอง: ไม่เชื่อมต่อภายนอก'); } };
  window.Logger = { log: () => {} };
  window.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: t => ({ setMimeType() { return this; }, t }) };
  window.HtmlService = { createHtmlOutput: () => ({ setTitle() { return this; }, addMetaTag() { return this; } }) };

  // ---------------------------------------------------------------- persistence
  const packSheets = () => Object.fromEntries(Object.entries(SHEETS).map(([k, v]) =>
    [k, v.data.map(r => r.map(c => c instanceof Date ? { __d: c.toISOString() } : c))]));
  window.GASMOCK = {
    save() {
      const db = { sheets: packSheets(), drive: DRIVE, props: PROPS, mail: MAIL, cache: CACHE };
      try { localStorage.setItem(KEY, JSON.stringify(db)); return true; }
      catch (e) {
        // quota: keep the data, drop email attachment copies then old emails
        MAIL.forEach(m => m.attachments.forEach(a => { a.data = ''; }));
        while (MAIL.length > 5) MAIL.shift();
        try { localStorage.setItem(KEY, JSON.stringify({ ...db, mail: MAIL })); return true; } catch (e2) { console.warn('demo: storage full', e2); return false; }
      }
    },
    load() {
      let db = null;
      try { db = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { db = null; }
      if (!db) return false;
      for (const [k, rows] of Object.entries(db.sheets || {})) {
        (SHEETS[k] = new Sheet(k)).data = rows.map(r => r.map(c => c && c.__d ? new Date(c.__d) : c));
      }
      Object.assign(DRIVE, db.drive || {}); Object.assign(PROPS, db.props || {}); Object.assign(CACHE, db.cache || {});
      (db.mail || []).forEach(m => MAIL.push(m));
      return true;
    },
    reset() { try { localStorage.removeItem(KEY); } catch (e) { /* ignore */ } },
    sheets: SHEETS, mail: MAIL, drive: DRIVE, cache: CACHE
  };
})();
