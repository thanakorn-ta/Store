/**
 * app.js
 * UI for the Store reorder comparison. Runs unchanged in two places:
 *   - GitHub Pages (static; files parsed in the browser)
 *   - Apps Script Web App (same page served by HtmlService; see src/Ui.js).
 *     There, AI calls and "save to Google Sheet" go through google.script.run
 *     so the API key stays in Script Properties.
 * Depends on globals from parsers.js / engine.js and SheetJS (XLSX).
 */
(() => {
// demo mode (web/demo.js) keeps its own copy so trial data never mixes with the real system
const STORE_KEY = typeof window.DEMO_API === 'function' ? 'store-reorder-ai:demo:v1' : 'store-reorder-ai:v1';
const AI_KEY = 'store-reorder-ai:apikey';
const MERGED_SHEET = 'รวมข้อมูล';

const SLOTS = [
  { type: 'minmax', no: 1, title: 'MIN/MAX (กรอกเอง)', desc: 'MIN_MAX_Calculated.xlsx', multi: false },
  { type: 'usage', no: 2, title: 'รวมการใช้ของ', desc: 'รวมการใช้ของ.xlsx หรือ การใช้งานเดือน*.xls ทีละเดือน (หลายไฟล์ได้)', multi: true },
  { type: 'balance', no: 3, title: 'ยอดคงเหลือสินค้า', desc: 'ยอดคงเหลือสินค้า.xls จาก Store', multi: false },
  { type: 'reorder', no: 4, title: 'รายงานสินค้าถึงจุดสั่งซื้อ', desc: 'รายงานสินค้าถึงจุดสั่งซื้อ.xls จาก Store', multi: false },
  { type: 'baseline', no: '±', title: 'ไฟล์รวมรอบก่อน (ไม่บังคับ)', desc: 'ไฟล์ที่ส่งออกจากระบบนี้รอบก่อน เพื่อเทียบคงเหลือ/ยอดแนะนำกับรอบนี้', multi: false }
];

const THAI_MONTH_FULL = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const PARSERS = { minmax: parseMinMax, usage: parseUsage, balance: parseBalance, reorder: parseReorder };

let state = load() || { files: {}, opt: { ...DEFAULTS }, cart: {}, qtySource: 'auto' };
state.opt = { ...DEFAULTS, ...state.opt };
let merged = { rows: [], months: [] };
let bySku = new Map();
let activeTab = 'home';
let pendingSlot = null;

// ------------------------------------------------------------ persistence

function load() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY)); } catch { return null; }
}
function save() {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); } catch { /* quota / private mode: keep in memory */ }
}

// ------------------------------------------------------------ helpers

// Server = the Apps Script project (src/). Reached two ways:
//   - page served by Apps Script itself  -> google.script.run.api(...)
//   - page on GitHub Pages / Vercel       -> HTTPS POST to the Web App URL (doPost in src/Api.js)
const GAS = typeof google !== 'undefined' && !!(google.script && google.script.run);
const DEFAULT_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby9p8MjJMZIeOufAyn3C6G9j4W3Q5iJZlgOrzfgt0vSOZTBCzyDjuI8a0GaXQZQlW4D/exec';
const API_URL_KEY = 'store-reorder-ai:api';
// A new Apps Script deployment gets a new /exec URL — it can be set from the sign-in screen
// (stored per browser) or with ?api=<url>, without editing this file.
let APPS_SCRIPT_URL = (() => {
  try {
    const q = new URLSearchParams(location.search).get('api');
    if (q && /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(q)) localStorage.setItem(API_URL_KEY, q);
    return localStorage.getItem(API_URL_KEY) || DEFAULT_APPS_SCRIPT_URL;
  } catch { return DEFAULT_APPS_SCRIPT_URL; }
})();
const SERVER = GAS || !!APPS_SCRIPT_URL;
const TOKEN_KEY = typeof window.DEMO_API === 'function' ? 'store-reorder-ai:demo:token' : 'store-reorder-ai:token';
let token = (() => { try { return localStorage.getItem(TOKEN_KEY) || ''; } catch { return ''; } })();
const setToken = t => { token = t || ''; try { t ? localStorage.setItem(TOKEN_KEY, t) : localStorage.removeItem(TOKEN_KEY); } catch { /* private mode */ } };

const DRAFT_OWNER_KEY = 'store-reorder-ai:draft-owner';
// Demo / trial mode (web/demo.js): the server code runs inside this page with sample data
const DEMO = !GAS && typeof window.DEMO_API === 'function';

function rawCall(fn, args) {
  if (DEMO) {
    return new Promise((resolve, reject) => setTimeout(() => {
      try { resolve(window.DEMO_API(token, fn, args)); renderDemoBar(); }
      catch (e) { reject(new Error(String((e && e.message) || e))); }
    }, 80));
  }
  if (GAS) {
    return new Promise((resolve, reject) =>
      google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).api(token, fn, args));
  }
  // text/plain keeps this a "simple" request (no CORS preflight, which Apps Script can't answer)
  return fetch(APPS_SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ token, fn, args }) })
    .catch(() => {
      // The browser hides the reason; in practice it is Google's HTML error page (no CORS header)
      const e = new Error('NETWORK'); e.network = true; throw e;
    })
    .then(r => r.text())
    .then(t => {
      let j;
      try { j = JSON.parse(t); } catch { const e = new Error('NETWORK'); e.network = true; throw e; }
      if (!j.ok) throw new Error(j.error);
      return j.result;
    });
}

// true when the Apps Script deployment still runs an older Code.gs than this page
let serverOutdated = false;

const gasCall = (fn, ...args) => rawCall(fn, args).catch(e => {
  const msg = String((e && e.message) || e);
  if (/SESSION_EXPIRED/.test(msg)) { setToken(''); applySession(null); throw new Error('หมดเวลาเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่'); }
  // Api.js answers "ไม่รู้จักคำสั่ง <fn>" when the deployed Code.gs predates this feature
  if (/ไม่รู้จักคำสั่ง/.test(msg)) {
    serverOutdated = true;
    if (activeTab === 'home') renderHome(); else if (activeTab === 'requests') renderRequests();
    throw new Error(`คำสั่ง "${fn}" ยังไม่มีในฝั่ง Apps Script — ต้องวางโค้ด Code.gs เวอร์ชันใหม่แล้ว Deploy เวอร์ชันใหม่ก่อน`);
  }
  throw e;
});

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (n, d = 0) => (n == null || n === '' || isNaN(n)) ? '' : Number(n).toLocaleString('th-TH', { minimumFractionDigits: d, maximumFractionDigits: d });
const fmt1 = n => (n == null ? '' : (Math.abs(n) < 10 && n % 1 ? fmt(n, 1) : fmt(Math.round(n))));
const money = n => '฿' + fmt(n);
const monthLabel = k => { const m = k % 100, y = Math.floor(k / 100); return MONTH_SHORT[m - 1] + (y ? ' ' + String(y).slice(-2) : ''); };

function toast(msg) {
  const t = $('#toast');
  t.textContent = msg; t.hidden = false;
  clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), 3200);
}

// ------------------------------------------------------------ file import

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      try { resolve(XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: false })); }
      catch (e) { reject(e); }
    };
    fr.onerror = () => reject(fr.error);
    fr.readAsArrayBuffer(file);
  });
}

const sheetRows = ws => XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });

async function importFiles(fileList, forcedType = null) {
  const results = [];
  for (const file of fileList) {
    try {
      const wb = await readWorkbook(file);
      let type = forcedType, records = null;

      if (wb.SheetNames.includes(MERGED_SHEET) && (!type || type === 'baseline')) {
        type = 'baseline';
        records = parseMergedExport(sheetRows(wb.Sheets[MERGED_SHEET]));
      } else {
        for (const name of wb.SheetNames) {
          const rows = sheetRows(wb.Sheets[name]);
          const t = detectType(rows, file.name);
          if (t === 'budget' && !forcedType) { results.push(await uploadBudget(rows, file.name)); type = 'budget'; break; }
          if (!t && !type) continue;
          const useType = type && type !== 'baseline' ? type : t;
          if (!PARSERS[useType]) continue;
          const recs = PARSERS[useType](rows, file.name);
          if (recs.length) { type = useType; records = (records || []).concat(recs); }
        }
      }

      if (type === 'budget') continue;
      if (!type || !records || !records.length) {
        results.push(`✗ ${file.name}: อ่านไม่ออกว่าเป็นไฟล์ประเภทไหน หรือไม่พบรหัสสินค้า`);
        continue;
      }
      if (type === 'usage' && records.some(r => !r.month)) {
        results.push(`⚠ ${file.name}: ${records.filter(r => !r.month).length} แถวไม่รู้ว่าเป็นเดือนไหน (ตั้งชื่อไฟล์ให้มีชื่อเดือน เช่น "การใช้งานเดือน ก.ค.")`);
      }
      const entry = { name: file.name, count: records.length, loadedAt: Date.now(), records };
      const slot = SLOTS.find(s => s.type === type);
      if (isShared(type)) {
        // Apps Script: the 4 Store slots are shared data in the database Sheet
        if (!isAdmin()) { results.push(`✗ ${file.name}: ข้อมูลกลาง — เฉพาะ Admin เพิ่ม/เปลี่ยนไฟล์ได้`); continue; }
        applyServerData(await gasCall('saveDataFile', type, file.name, records));
        results.push(`✓ ${file.name} → ${slot.title} (${records.length} แถว) บันทึกลง Google Sheet แล้ว — ทุกคนเห็นข้อมูลนี้`);
        continue;
      }
      if (slot.multi) {
        state.files[type] = (state.files[type] || []).filter(f => f.name !== file.name).concat(entry);
      } else {
        state.files[type] = [entry];
      }
      results.push(`✓ ${file.name} → ${slot.title} (${records.length} แถว)`);
    } catch (e) {
      console.error(e);
      results.push(`✗ ${file.name}: ${e.message}`);
    }
  }
  save();
  recompute();
  toast(results.join('\n'));
}

async function removeFile(type, name) {
  if (isShared(type)) {
    if (!confirm(`นำไฟล์ "${name}" ออกจากข้อมูลกลาง?\nทุกคนจะไม่เห็นข้อมูลจากไฟล์นี้อีก (ลบแถวในชีตด้วย)`)) return;
    const slotEl = document.querySelector(`[data-slot="${type}"]`);
    slotEl && slotEl.classList.add('busy');
    try { applyServerData(await gasCall('removeDataFile', type, name)); toast(`นำ "${name}" ออกแล้ว`); }
    catch (e) { toast('นำออกไม่สำเร็จ: ' + errMsg(e)); slotEl && slotEl.classList.remove('busy'); }
    return;
  }
  state.files[type] = (state.files[type] || []).filter(f => f.name !== name);
  save(); recompute();
}

// ------------------------------------------------------------ shared data (Apps Script; server: src/DataFiles.js)

const SHARED_SLOTS = ['minmax', 'usage', 'balance', 'reorder'];
let sharedLoaded = false;
const isShared = type => SERVER && SHARED_SLOTS.includes(type);
const canEditSlot = type => !isShared(type) || isAdmin();

function applyServerData(data) {
  for (const t of SHARED_SLOTS) state.files[t] = data.files[t] || [];
  sharedLoaded = true;
  if (data.sheetUrl) { $('#lnkSheet').href = data.sheetUrl; $('#lnkSheet').hidden = false; }
  save();
  recompute();
}

async function loadServerData(showToast = false) {
  const btn = $('#btnReloadData');
  btn.disabled = true;
  try {
    applyServerData(await gasCall('getStoreData'));
    if (showToast) toast('โหลดข้อมูลล่าสุดจาก Google Sheet แล้ว');
  } catch (e) {
    toast('โหลดข้อมูลกลางไม่ได้: ' + errMsg(e));
  } finally {
    btn.disabled = false;
  }
}
$('#btnReloadData').addEventListener('click', () => loadServerData(true));

// ------------------------------------------------------------ slots UI

function renderSlots() {
  if (SERVER) {
    $('#filesHint').textContent = isAdmin()
      ? 'ข้อมูลกลางเก็บใน Google Sheet — ทุกคนเห็นชุดเดียวกัน · เพิ่ม/เปลี่ยน/นำไฟล์ออกได้ตลอด · แก้รายแถวได้ในชีต data_* แล้วกด "โหลดข้อมูลใหม่" · ช่อง ± เป็นของเครื่องนี้เท่านั้น'
      : 'ข้อมูลกลางจาก Google Sheet (Admin เป็นผู้อัปเดต) · ช่อง ± ใส่ไฟล์รอบก่อนเพื่อเทียบได้เอง (เก็บในเครื่องนี้)';
    $('#btnReloadData').hidden = !isMember();
  }
  $('#slots').innerHTML = SLOTS.map(s => {
    const files = state.files[s.type] || [];
    const edit = canEditSlot(s.type);
    const waiting = isShared(s.type) && !sharedLoaded;
    return `<div class="slot ${files.length ? 'filled' : ''}" data-slot="${s.type}">
      <div class="slot-title"><span class="slot-no">${s.no}</span>${esc(s.title)}</div>
      <div class="slot-desc">${esc(s.desc)}</div>
      <ul class="slot-files">${files.map(f => `
        <li><span class="fname" title="${esc(f.name)}">${esc(f.name)}${f.uploadedAt ? `<span class="fby">${esc(f.uploadedAt)}${f.uploadedBy ? ' · ' + esc(f.uploadedBy.split('@')[0]) : ''}</span>` : ''}</span>
          <span class="fmeta">${fmt(f.count)} แถว</span>
          ${edit ? `<button class="icon-btn" data-remove="${esc(f.name)}" title="นำไฟล์ออก" aria-label="นำไฟล์ออก">×</button>` : ''}</li>`).join('')}
      </ul>
      ${waiting ? '<div class="slot-note">กำลังโหลดข้อมูลกลาง…</div>'
        : edit ? `<button class="btn sm ghost slot-add" data-pick="${s.type}">${s.multi && files.length ? 'เพิ่มไฟล์' : files.length ? 'เปลี่ยนไฟล์' : 'เลือกไฟล์'}</button>`
        : `<div class="slot-note">${files.length ? 'อัปเดตโดย Admin' : 'ยังไม่มีข้อมูล — รอ Admin อัปโหลด'}</div>`}
    </div>`;
  }).join('');
}

$('#slots').addEventListener('click', e => {
  const pick = e.target.closest('[data-pick]');
  if (pick) { pendingSlot = pick.dataset.pick; $('#fileInput').multiple = pendingSlot === 'usage'; $('#fileInput').click(); return; }
  const rm = e.target.closest('[data-remove]');
  if (rm) removeFile(rm.closest('[data-slot]').dataset.slot, rm.dataset.remove);
});
$('#fileInput').addEventListener('change', e => {
  const files = [...e.target.files];
  e.target.value = '';
  if (files.length) importFiles(files, pendingSlot);
  pendingSlot = null;
});

// Drop anywhere = auto-detect; drop on a slot = force that slot's type
document.addEventListener('dragover', e => {
  e.preventDefault();
  document.querySelectorAll('.slot').forEach(s => s.classList.toggle('drag', s.contains(e.target)));
});
document.addEventListener('dragleave', e => { if (!e.relatedTarget) document.querySelectorAll('.slot').forEach(s => s.classList.remove('drag')); });
document.addEventListener('drop', e => {
  e.preventDefault();
  document.querySelectorAll('.slot').forEach(s => s.classList.remove('drag'));
  const slot = e.target.closest && e.target.closest('[data-slot]');
  const files = [...(e.dataTransfer?.files || [])];
  if (files.length) importFiles(files, slot ? slot.dataset.slot : null);
});

// ------------------------------------------------------------ compute

function recompute() {
  const f = t => (state.files[t] || []).flatMap(x => x.records);
  merged = mergeAll({
    minmax: f('minmax'), usage: f('usage'), balance: f('balance'), reorder: f('reorder'), baseline: f('baseline')
  }, state.opt);
  bySku = new Map(merged.rows.map(r => [r.sku, r]));

  renderSlots();
  const has = merged.rows.length > 0;
  $('#summary').hidden = !has;
  $('#work').hidden = SERVER && !isMember();
  $('#btnExportAll').disabled = !has;
  if (!has) { renderActive(); return; }
  renderSettings();
  renderCards();
  renderPcFilter();
  renderActive();
}

function renderSettings() {
  document.querySelectorAll('[data-opt]').forEach(i => { i.value = state.opt[i.dataset.opt]; });
}
document.querySelectorAll('[data-opt]').forEach(i => i.addEventListener('change', () => {
  const v = parseFloat(i.value);
  if (!isNaN(v) && v >= 0) { state.opt[i.dataset.opt] = v; save(); recompute(); }
}));

function renderCards() {
  const s = summarize(merged.rows);
  const nM = merged.months.length;
  const dt = draftTotals();
  const missing = SLOTS.filter(x => x.type !== 'baseline' && !(state.files[x.type] || []).length).map(x => x.title);
  $('#cards').innerHTML = `
    <div class="card"><div class="k">สินค้าทั้งหมด (รวม 4 ไฟล์)</div><div class="v">${fmt(s.skus)}</div>
      <div class="s">มูลค่าคงเหลือ ${money(s.stockValue)}</div></div>
    <div class="card"><div class="k">ข้อมูลการใช้</div><div class="v">${nM} เดือน</div>
      <div class="s">${nM ? monthLabel(merged.months[0]) + ' – ' + monthLabel(merged.months[nM - 1]) : 'ยังไม่มีไฟล์การใช้'} · ประจำ ${s.regular} / ตามงาน ${s.job}</div></div>
    <div class="card store"><div class="k">Store แจ้งสั่ง</div><div class="v">${fmt(s.storeCount)}</div>
      <div class="s">ประมาณ ${money(s.storeValue)}</div></div>
    <div class="card manual"><div class="k">ตาม MIN/MAX เดิม (กรอกเอง)</div><div class="v">${fmt(s.manualCount)}</div>
      <div class="s">ประมาณ ${money(s.manualValue)} · Avg สูงเกินจริง ${s.overEstimated} รายการ</div></div>
    <div class="card actual"><div class="k">ตามยอดใช้จริง</div><div class="v">${fmt(s.actualCount)}</div>
      <div class="s">ประมาณ ${money(s.actualValue)}</div></div>
    <div class="card dead"><div class="k">Dead stock</div><div class="v">${fmt(s.dead)}</div>
      <div class="s">มูลค่าค้าง ${money(s.deadValue)}</div></div>
    <div class="card"><div class="k">ในคำขอที่กำลังทำ</div><div class="v">${fmt(dt.n)}</div>
      <div class="s">ประมาณ ${money(dt.v)}</div></div>
    ${missing.length ? `<div class="card"><div class="k">ยังไม่ได้นำเข้า</div><div class="s" style="color:var(--warn)">${missing.map(esc).join('<br>')}</div></div>` : ''}`;
  updateCartBar();
}

// ------------------------------------------------------------ sticky basket (always shows the next step)

function updateCartBar() {
  const { n, v } = draftTotals();
  $('#cartBadge').textContent = n;
  $('#cartBadge').hidden = !n;
  const show = n > 0 && ['compare', 'merged', 'home', 'budget'].includes(activeTab) && !$('#work').hidden && isMember();
  $('#cartBar').hidden = !show;
  document.body.classList.toggle('has-cartbar', show);
  $('#cartBarText').innerHTML = `${state.draft.editId ? 'กำลังแก้ไข ' + esc(state.draft.editId) + ' · ' : 'ในคำขอ '}<b>${fmt(n)}</b> รายการ · ประมาณ <b>${money(v)}</b>`;
}
$('#cartBarBtn').addEventListener('click', () => goTab('order'));

function goTab(name) {
  // "requests" has one menu item per role — use the visible one
  const t = [...document.querySelectorAll(`.tab[data-tab="${name}"]`)].find(x => !x.hidden);
  if (t) t.click();
}

function renderPcFilter() {
  const sel = $('#fPc'), cur = sel.value;
  const pcs = new Map();
  merged.rows.forEach(r => { const k = r.pc || r.group || ''; if (k && !pcs.has(k)) pcs.set(k, r.pc ? r.pcName : ''); });
  sel.innerHTML = '<option value="">ทุก PC</option>' +
    [...pcs].sort().map(([k, n]) => `<option value="${esc(k)}">${esc(k)} ${esc(n)}</option>`).join('');
  sel.value = cur;
}

// ------------------------------------------------------------ filtering

function filtered(mode) {
  const q = $('#fSearch').value.trim().toLowerCase();
  const pc = $('#fPc').value, cat = $('#fCat').value, show = mode || $('#fShow').value;
  return merged.rows.filter(r => {
    if (q && !(r.sku.toLowerCase().includes(q) || r.name.toLowerCase().includes(q))) return false;
    if (pc && (r.pc || r.group) !== pc) return false;
    if (cat && r.category !== cat) return false;
    const any = r.orderStore > 0 || r.orderManual > 0 || r.orderActual > 0;
    switch (show) {
      case 'any': return any || draftHas(r.sku);
      case 'diff': return any && !(r.orderStore === r.orderManual && r.orderManual === r.orderActual);
      case 'store': return r.orderStore > 0;
      case 'actual': return r.orderActual > 0;
      case 'flag': return r.flags.length > 0;
      default: return true;
    }
  });
}
['#fSearch', '#fPc', '#fCat', '#fShow'].forEach(s => $(s).addEventListener('input', () => renderActive()));
$('#qtySource').value = state.qtySource;
$('#qtySource').addEventListener('change', e => { state.qtySource = e.target.value; save(); renderActive(); });

// ------------------------------------------------------------ tabs

document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
  activeTab = t.dataset.tab;
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === t));
  document.querySelectorAll('.tabpane').forEach(p => (p.hidden = p.id !== 'tab-' + activeTab));
  $('#filters').hidden = !['compare', 'merged'].includes(activeTab);
  // filters sit above the table of the active view
  const pane = $('#tab-' + activeTab);
  if (!$('#filters').hidden) pane.querySelector('.page-head').after($('#filters'));
  window.scrollTo(0, 0);
  updateCartBar();
  renderActive();
}));

function renderActive() {
  if (activeTab === 'home') renderHome();
  else if (activeTab === 'compare') renderCompare();
  else if (activeTab === 'merged') renderMerged();
  else if (activeTab === 'order') renderOrder();
  else if (activeTab === 'requests') loadRequests();
  else if (activeTab === 'admin') { members = null; renderAdmin(); } // always fresh: sign-ups arrive any time
  else if (activeTab === 'budget') renderBudget();
  else if (activeTab === 'delivery') { renderDelivery(); loadRequests(); }
}

// ------------------------------------------------------------ compare table

const catTag = c => c ? `<span class="tag ${{ [CATEGORY.REGULAR]: 'reg', [CATEGORY.JOB]: 'job', [CATEGORY.DEAD]: 'dead' }[c] || 'none'}">${esc(c)}</span>` : '';
const o = n => `<span class="${n > 0 ? '' : 'zero'}">${n > 0 ? fmt(n) : '0'}</span>`;

function spark(r) {
  if (!r.months.length) return '';
  const vals = r.months.map(k => r.usageByMonth[k] || 0);
  const mx = Math.max(...vals, 1);
  const title = r.months.map((k, i) => `${monthLabel(k)}: ${fmt(vals[i])}`).join('\n');
  return `<span class="spark" title="${esc(title)}">${vals.map(v => `<b class="${v ? '' : 'z'}" style="height:${Math.max(1, Math.round(v / mx * 18))}px"></b>`).join('')}</span>`;
}

/** ต่อท้ายแถวที่ติ๊กแล้ว: ของชิ้นนี้ไปอยู่ใน Budget ไหน เดือนไหน เป็นเงินเท่าไร */
function rowBudgetCell(c) {
  if (!c) return '<td class="bud"><span class="dim">–</span></td>';
  const line = budgetLine(c.budgetKey);
  const m = Number(c.month);
  if (!line || !m) return '<td class="bud"><span class="warn-txt">ยังไม่เลือกงบ</span></td>';
  return `<td class="bud"><b>${money(lineAmt(c))}</b>
    <div class="pc">${esc(line.location)} · ${esc(line.glCode)} · ${esc(THAI_MONTH_FULL[m] || '')}</div></td>`;
}

/** สรุปใต้ตาราง: แต่ละ Budget + เดือน เลือกไปกี่รายการ เป็นเงินเท่าไร และเหลือเท่าไร */
function renderCompareBudget() {
  const groups = draftGroups();
  const box = $('#compareBudget');
  if (!groups.length) { box.innerHTML = ''; return; }
  const total = groups.reduce((a, g) => a + g.amount, 0);
  box.innerHTML = `<div class="ps-head">ที่ติ๊กไว้ทั้งหมด ${fmt(state.draft.lines.length)} รายการ · รวม <b>${money(total)}</b>
      <button class="btn sm" data-go="order">ไปหน้าสั่งซื้อ</button></div>
    ${groups.map(g => {
      if (!g.line || !g.month) return `<div class="ps-row warn"><span>${fmt(g.n)} รายการ ยังไม่ได้เลือก Budget/เดือน</span><span class="num">${money(g.amount)}</span></div>`;
      return `<div class="ps-row ${g.over ? 'over' : ''}">
        <span><b>${esc(g.line.location)} · ${esc(g.line.glCode)} ${esc(g.line.glName)}</b>
          <span class="pc">${esc(g.line.company)} · เดือน${esc(THAI_MONTH_FULL[g.month])} · ${fmt(g.n)} รายการ</span></span>
        <span class="num">${money(g.amount)}
          <span class="pc">งบเหลือ ${money(g.b.available)} → ${g.over ? 'เกิน ' + money(-g.after) : 'คงเหลือ ' + money(g.after)}</span></span></div>`;
    }).join('')}`;
}

function renderCompare() {
  const pl = budgetLine(state.pick.key);
  $('#compareHint').innerHTML = pl
    ? `ติ๊กรายการที่จะสั่ง → ใส่ใน <b>${esc(pl.label)}</b> · เดือน <b>${THAI_MONTH_FULL[Number(state.pick.month)] || ''}</b> <button class="link-btn" data-go="order">เปลี่ยน Budget</button>`
    : 'ติ๊กรายการที่จะสั่ง — ยังไม่ได้เลือก Budget: <button class="link-btn" data-go="order">เลือก Budget ก่อน</button> (หรือเลือกทีหลังในหน้าสั่งซื้อ)';
  const rows = filtered();
  renderCompareBudget();
  const hasBase = rows.some(r => r.base);
  $('#countInfo').textContent = `แสดง ${fmt(rows.length)} จาก ${fmt(merged.rows.length)} รายการ`;
  if (!rows.length) {
    const msg = merged.rows.length ? 'ไม่มีรายการตามเงื่อนไข' : 'ยังไม่มีข้อมูล — นำเข้าไฟล์ Store ที่ช่อง 1–4 ด้านบนก่อน';
    $('#tblCompare').innerHTML = `<tbody><tr><td class="empty">${msg}</td></tr></tbody>`;
    return;
  }
  document.body.classList.toggle('detail-mode', !!state.detail);
  if (!state.detail) { renderCompareSimple(rows); return; }

  $('#tblCompare').innerHTML = `
    <thead>
      <tr class="grp"><th colspan="4"></th><th colspan="3">การใช้ / คงเหลือ</th>
        <th colspan="2" class="col-manual">MIN / MAX</th>
        <th colspan="3">ยอดแนะนำสั่ง</th>${hasBase ? '<th colspan="2">เทียบรอบก่อน</th>' : ''}<th colspan="3"></th></tr>
      <tr class="sub">
        <th title="เลือกเพื่อสั่งซื้อ">สั่ง</th><th>รหัส</th><th>สินค้า</th><th>ประเภท</th>
        <th class="n">คงเหลือ</th><th class="n" title="Avg/เดือน ที่กรอกเอง → ใช้จริงเฉลี่ย">Avg เดิม → จริง</th><th>รายเดือน</th>
        <th class="n col-manual">เดิม</th><th class="n col-actual">จากใช้จริง</th>
        <th class="n col-store">Store</th><th class="n col-manual">MIN/MAX เดิม</th><th class="n col-actual">ใช้จริง</th>
        ${hasBase ? '<th class="n">คงเหลือก่อน</th><th class="n">Δ</th>' : ''}
        <th class="n">จำนวนสั่ง</th><th>ข้อสังเกต</th><th class="n">ใส่ในงบ</th>
      </tr>
    </thead>
    <tbody>${rows.map(r => {
      const c = draftLineOf(r.sku);
      const m = r.manual;
      const d = r.base ? r.balance - r.base.balance : null;
      return `<tr data-sku="${r.sku}" class="${c ? 'sel' : ''}">
        <td><input type="checkbox" data-pick-sku ${c ? 'checked' : ''} aria-label="เลือก ${esc(r.sku)}"></td>
        <td class="sku">${r.sku}<div class="pc">${esc(r.pc || r.group)}</div></td>
        <td class="name">${esc(r.name)}<div class="pc">${esc(r.unit)}${r.cost ? ' · ' + money(r.cost) + '/หน่วย' : ''}</div></td>
        <td>${catTag(r.category)}</td>
        <td class="n">${fmt(r.balance)}</td>
        <td class="n"><span class="dim">${m ? fmt1(m.avg) : '–'}</span> → ${fmt1(r.avgActual)}</td>
        <td>${spark(r)}</td>
        <td class="n col-manual">${m && m.min != null ? `${fmt(m.min)} / ${fmt(m.max)}` : '<span class="dim">–</span>'}</td>
        <td class="n col-actual">${r.minNew != null ? `${fmt(r.minNew)} / ${fmt(r.maxNew)}` : '<span class="dim">–</span>'}</td>
        <td class="n o col-store">${r.store ? o(r.orderStore) : '<span class="dim">–</span>'}</td>
        <td class="n o col-manual">${m ? o(r.orderManual) : '<span class="dim">–</span>'}</td>
        <td class="n o col-actual">${r.nMonths ? o(r.orderActual) : '<span class="dim">–</span>'}</td>
        ${hasBase ? `<td class="n dim">${r.base ? fmt(r.base.balance) : ''}</td>
          <td class="n delta ${d > 0 ? 'up' : d < 0 ? 'down' : ''}">${d == null ? '' : (d > 0 ? '+' : '') + fmt(d)}</td>` : ''}
        <td class="n"><input class="qty" type="number" min="0" step="1" data-qty value="${c ? c.qty : ''}" ${c ? '' : 'disabled'} aria-label="จำนวนสั่ง"></td>
        <td class="flags">${r.flags.map(f => `<div>${esc(f)}</div>`).join('')}</td>
        ${rowBudgetCell(c)}
      </tr>`;
    }).join('')}</tbody>`;
}

const SOURCE_LABEL = { actual: 'ใช้จริง', store: 'Store', manual: 'MIN/MAX เดิม' };

/** Which method the suggested qty comes from, following the "ยอดแนะนำจาก" setting. */
function suggestion(r) {
  const src = state.qtySource;
  if (src !== 'auto') return { qty: defaultQty(r, src), from: src };
  if (r.orderActual) return { qty: r.orderActual, from: 'actual' };
  if (r.orderStore) return { qty: r.orderStore, from: 'store' };
  return { qty: r.orderManual, from: r.orderManual ? 'manual' : '' };
}

/** Default view: one suggested number per item; the 3-method comparison is behind the toggle. */
function renderCompareSimple(rows) {
  $('#tblCompare').innerHTML = `
    <thead><tr>
      <th title="เลือกเพื่อสั่งซื้อ">สั่ง</th><th>สินค้า</th><th>PC</th>
      <th class="n">คงเหลือ</th><th class="n">ใช้/เดือน</th><th>แนวโน้ม</th>
      <th class="n">แนะนำสั่ง</th><th class="n">จำนวนสั่ง</th><th>ข้อควรรู้</th><th class="n">ใส่ในงบ</th>
    </tr></thead>
    <tbody>${rows.map(r => {
      const c = draftLineOf(r.sku);
      const s = suggestion(r);
      const warn = r.flags.find(f => /ต้องถาม|ไม่ถูกเบิก|ไม่ถึงจุด|≠/.test(f)) || '';
      return `<tr data-sku="${r.sku}" class="${c ? 'sel' : ''}">
        <td><input type="checkbox" data-pick-sku ${c ? 'checked' : ''} aria-label="เลือก ${esc(r.sku)}"></td>
        <td class="name"><b class="item-name">${esc(r.name)}</b><div class="pc">${r.sku} · ${esc(r.unit)}${r.cost ? ' · ' + money(r.cost) + '/หน่วย' : ''}</div></td>
        <td class="pc">${esc(r.pc || r.group)}</td>
        <td class="n">${fmt(r.balance)}</td>
        <td class="n">${fmt1(r.avgActual)}</td>
        <td>${spark(r)}</td>
        <td class="n">${s.qty > 0 ? `<b>${fmt(s.qty)}</b><div class="src-tag">${esc(SOURCE_LABEL[s.from] || '')}</div>` : '<span class="zero">–</span>'}</td>
        <td class="n"><input class="qty" type="number" min="0" step="1" data-qty value="${c ? c.qty : ''}" ${c ? '' : 'disabled'} aria-label="จำนวนสั่ง"></td>
        <td class="flags">${esc(warn)}</td>
        ${rowBudgetCell(c)}
      </tr>`;
    }).join('')}</tbody>`;
}

$('#detailToggle').checked = !!state.detail;
$('#detailToggle').addEventListener('change', e => { state.detail = e.target.checked; save(); renderCompare(); });

$('#tblCompare').addEventListener('change', e => {
  const tr = e.target.closest('tr[data-sku]');
  if (!tr) return;
  const sku = tr.dataset.sku;
  if (e.target.matches('[data-pick-sku]')) {
    setPicked(sku, e.target.checked);
    tr.classList.toggle('sel', e.target.checked);
    const q = tr.querySelector('[data-qty]');
    q.disabled = !e.target.checked;
    q.value = e.target.checked ? draftLineOf(sku).qty : '';
    if (e.target.checked) q.select();
    if (e.target.checked && !state.pick.key) toast('เพิ่มแล้ว — ยังไม่ได้เลือก Budget: เลือกได้ในหน้า "สั่งซื้อตาม Budget"');
  } else if (e.target.matches('[data-qty]') && draftLineOf(sku)) {
    draftLineOf(sku).qty = Math.max(0, Math.round(Number(e.target.value) || 0));
  }
  const bud = tr.querySelector('td.bud'); // ช่อง "ใส่ในงบ" ท้ายแถว + สรุปใต้ตาราง
  if (bud) bud.outerHTML = rowBudgetCell(draftLineOf(sku));
  renderCompareBudget();
  save(); renderCards();
});

document.querySelectorAll('[data-bulk]').forEach(b => b.addEventListener('click', () => {
  const mode = b.dataset.bulk;
  const rows = filtered();
  for (const r of rows) {
    if (mode === 'none') setPicked(r.sku, false);
    else if (mode === 'suggested' && suggestion(r).qty > 0) setPicked(r.sku, true);
    else if (mode === 'visible') setPicked(r.sku, true);
    else if (mode === 'actual' && r.orderActual > 0) setPicked(r.sku, true);
    else if (mode === 'store' && r.orderStore > 0) setPicked(r.sku, true);
  }
  save(); renderCards(); renderCompare();
}));

// ------------------------------------------------------------ merged table

function mergedColumns() {
  const cols = [
    ['รหัสสินค้า', r => r.sku], ['สินค้า', r => r.name], ['PC', r => r.pc], ['ชื่อ PC', r => r.pcName],
    ['กลุ่ม (ไฟล์การใช้/MIN-MAX)', r => r.group], ['หน่วย', r => r.unit], ['ที่เก็บ', r => r.location],
    ['คงเหลือ', r => r.balance, 1], ['ต้นทุน/หน่วย', r => r.cost, 1], ['มูลค่าคงเหลือ', r => r.value, 1]
  ];
  for (const k of merged.months) cols.push([monthLabel(k), r => r.usageByMonth[k] || 0, 1, 'm']);
  cols.push(
    ['ใช้รวม', r => r.usageTotal, 1], ['เดือนที่มีการเบิก', r => r.monthsUsed, 1],
    ['เฉลี่ยใช้จริง/เดือน', r => Math.round(r.avgActual * 100) / 100, 1],
    ['Avg กรอกเอง', r => r.manual ? r.manual.avg : '', 1],
    ['Avg เดิม ÷ จริง', r => r.ratio == null ? '' : (r.ratio === Infinity ? 'ใช้จริง 0' : Math.round(r.ratio * 10) / 10), 1],
    ['ประเภท', r => r.category], ['Lead time', r => r.lead, 1],
    ['MIN เดิม', r => r.manual?.min ?? '', 1], ['MAX เดิม', r => r.manual?.max ?? '', 1],
    ['Safety ใหม่', r => r.safetyNew ?? '', 1], ['MIN ใหม่', r => r.minNew ?? '', 1], ['MAX ใหม่', r => r.maxNew ?? '', 1],
    ['คงพอใช้ (เดือน)', r => r.coverMonths == null ? '' : Math.round(r.coverMonths * 10) / 10, 1],
    ['Store: จุดต่ำสุด', r => r.store ? r.store.min : '', 1],
    ['Store แจ้งสั่ง', r => r.store ? r.orderStore : '', 1],
    ['Store: สั่งต่อครั้ง', r => r.store ? r.store.lot : ''],
    ['แนะนำสั่ง (MIN/MAX เดิม)', r => r.orderManual, 1],
    ['แนะนำสั่ง (ใช้จริง)', r => r.orderActual, 1],
    ['ในคำขอ', r => draftLineOf(r.sku)?.qty ?? '', 1],
    ['มีในไฟล์', r => ['minmax', 'usage', 'balance', 'reorder'].filter(k => r.sources[k]).map(k => ({ minmax: '1', usage: '2', balance: '3', reorder: '4' }[k])).join(',')],
    ['ข้อสังเกต', r => r.flags.join(' | ')]
  );
  return cols;
}

function renderMerged() {
  const rows = filtered();
  const cols = mergedColumns();
  const show = rows.slice(0, 1500);
  $('#tblMerged').innerHTML = `<thead><tr class="sub" style="top:0">${cols.map(c => `<th class="${c[2] ? 'n' : ''}" style="top:0">${esc(c[0])}</th>`).join('')}</tr></thead>
    <tbody>${show.map(r => `<tr>${cols.map(c => {
      const v = c[1](r);
      return `<td class="${c[2] ? 'n' : ''} ${c[0] === 'สินค้า' ? 'name' : ''} ${c[3] === 'm' && !v ? 'zero' : ''}">${esc(typeof v === 'number' ? fmt(v, v % 1 ? 1 : 0) : v)}</td>`;
    }).join('')}</tr>`).join('')}</tbody>`;
}

// ------------------------------------------------------------ order by budget: pick a budget line + month, then the items
//
// state.draft = { lines: [{ uid, sku, name, unit, cost, qty, note, budgetKey, month, pc, pcName, balance, custom }],
//                 note, over: { "key|month": { on, reason } }, editId, keepFiles: [{ id, name, size }] }
//   editId    = request called back for editing (resubmitted to the same id)
//   keepFiles = its quotations already on the server that stay attached
// state.pick  = { key, month } — the budget line + month the page is filling now
// New quotation files stay in memory (draftFiles) until the request is sent.

const uid = () => Math.random().toString(36).slice(2, 10);
state.draft = state.draft || {};
state.draft = { lines: [], note: '', over: {}, editId: null, keepFiles: [], ...state.draft };
state.pick = state.pick || { key: '', month: new Date().getMonth() + 1 };
if (state.cart && Object.keys(state.cart).length) { // older versions kept a flat cart keyed by SKU
  for (const [sku, c] of Object.entries(state.cart)) {
    state.draft.lines.push({ uid: uid(), sku, qty: Math.max(1, Number(c.qty) || 1), note: c.note || '', budgetKey: c.budgetKey || '', month: Number(c.month) || 0 });
  }
}
state.cart = {};
let draftFiles = [];

// a budget line's key itself contains "|" (company|location|gl), so split "key|month" from the right
const splitKM = k => ({ key: String(k).slice(0, String(k).lastIndexOf('|')), month: Number(String(k).slice(String(k).lastIndexOf('|') + 1)) || 0 });

const lineStore = l => bySku.get(l.sku);
const lineName = l => l.name || (lineStore(l) || {}).name || l.sku;
const lineUnit = l => l.unit || (lineStore(l) || {}).unit || '';
const lineCost = l => Number(l.cost != null && l.cost !== '' ? l.cost : (lineStore(l) || {}).cost) || 0;
const lineAmt = l => (Number(l.qty) || 0) * lineCost(l);
const lineBalance = l => { const r = lineStore(l); return r ? r.balance : (l.balance ?? ''); };
const snapshot = r => ({ sku: r.sku, name: r.name, unit: r.unit, cost: r.cost || 0, pc: r.pc || r.group || '', pcName: r.pcName || r.group || '', balance: r.balance });
const budgetLine = key => budgetOpts && budgetOpts.lines ? budgetOpts.lines.find(l => l.key === key) : null;
const hasBudget = () => !!(budgetOpts && budgetOpts.lines && budgetOpts.lines.length);

// Store สั่งซื้อจากงบ 3 หมวดนี้เท่านั้น — หมวดอื่น (เช่น ค่าไฟฟ้าป้ายโฆษณา) ซ่อนไว้
// จับจากชื่อหมวด ไม่ใช่เลข GL เพราะแต่ละบริษัทใช้เลขไม่เหมือนกัน
const GL_WORDS = ['ซ่อมแซมบำรุงรักษา', 'วัสดุสิ้นเปลือง', 'ซ่อมภาพ'];
const GL_ALL_KEY = 'store-reorder-ai:gl-all';
let showAllGl = false;
try { showAllGl = localStorage.getItem(GL_ALL_KEY) === '1'; } catch { /* private mode */ }
const glWanted = l => GL_WORDS.some(w => `${l.glName || ''} ${l.expenseGroup || ''} ${l.label || ''}`.includes(w));
/** Budget ที่เอามาแสดง/ให้เลือก · budgetLine() ยังหาได้ทุกบรรทัด เพื่อให้คำขอเก่าแสดงผลถูก */
function budgetLines() {
  const all = (budgetOpts && budgetOpts.lines) || [];
  if (showAllGl) return all;
  const keep = all.filter(glWanted);
  return keep.length ? keep : all; // ชื่อหมวดในไฟล์ไม่ตรงคำที่รู้จัก → แสดงทั้งหมดดีกว่าหน้าว่าง
}
const glHidden = () => ((budgetOpts && budgetOpts.lines) || []).length - budgetLines().length;

function draftTotals() {
  const ls = state.draft.lines;
  return { n: ls.length, v: ls.reduce((a, l) => a + lineAmt(l), 0) };
}
const cartTotals = draftTotals; // older name used by the sticky bar / home
const draftHas = sku => state.draft.lines.some(l => l.sku === sku);
const draftLineOf = sku => state.draft.lines.find(l => l.sku === sku);
const draftAmtFor = (key, m) => state.draft.lines.filter(l => l.budgetKey === key && Number(l.month) === Number(m)).reduce((a, l) => a + lineAmt(l), 0);

/** Adds an item to the budget line + month being filled (same item there again = +qty). */
function addLine(item, qty) {
  const key = state.pick.key || '', month = key ? Number(state.pick.month) || 0 : 0;
  const same = !item.custom && state.draft.lines.find(l => l.sku === item.sku && l.budgetKey === key && Number(l.month) === month);
  if (same) { same.qty = (Number(same.qty) || 0) + Math.max(1, Math.round(qty || 1)); return same; }
  const l = { uid: uid(), ...item, qty: Math.max(1, Math.round(qty || 1)), note: '', budgetKey: key, month };
  state.draft.lines.push(l);
  return l;
}

/** Store suggestion page: tick = add to the budget being filled, untick = remove everywhere. */
function setPicked(sku, on) {
  const r = bySku.get(sku);
  if (!r) return;
  if (on) { if (!draftHas(sku)) addLine(snapshot(r), Math.max(1, defaultQty(r, state.qtySource) || suggestion(r).qty || 1)); }
  else state.draft.lines = state.draft.lines.filter(l => l.sku !== sku);
}

/** Draft grouped by budget line x month: [{ k, key, month, line, amount, n, b, after, over }]. */
function draftGroups() {
  const g = new Map();
  for (const l of state.draft.lines) {
    const key = l.budgetKey || '', month = key ? Number(l.month) || 0 : 0, k = key + '|' + month;
    if (!g.has(k)) g.set(k, { k, key, month, amount: 0, n: 0 });
    const x = g.get(k); x.amount += lineAmt(l); x.n++;
  }
  return [...g.values()].map(x => {
    x.line = budgetLine(x.key);
    if (x.line && x.month) { x.b = budgetAvail(x.line, x.month); x.after = x.b.available - x.amount; x.over = x.after < -0.004; }
    return x;
  }).sort((a, b) => (a.line ? 0 : 1) - (b.line ? 0 : 1) || (a.line && b.line ? a.line.label.localeCompare(b.line.label) : 0) || a.month - b.month);
}

function renderOrder() {
  updateCartBar();
  const body = $('#orderBody'), empty = $('#orderEmpty');
  const ed = state.draft.editId;
  $('#orderEditing').hidden = !ed;
  $('#orderEditing').innerHTML = ed ? `<b>กำลังแก้ไขคำขอ ${esc(ed)}</b> — แก้รายการ / Budget / ใบเสนอราคา แล้วกด "บันทึกและส่งใหม่" · คำขอจะกลับไปรอ Admin ตรวจ` : '';
  $('#btnSubmitReq').textContent = ed ? 'บันทึกและส่งใหม่ให้ Admin ตรวจ' : 'ส่งให้ Admin ตรวจ';
  $('#btnCancelEdit').hidden = !ed;
  $('#roundBox').hidden = !isAdmin() || !!ed;
  if (!isMember()) { body.hidden = true; empty.innerHTML = ''; return; }
  if (!budgetOpts) { body.hidden = true; empty.innerHTML = '<div class="empty">กำลังโหลด Budget…</div>'; return; }
  if (budgetOpts.error || !hasBudget()) {
    body.hidden = true;
    empty.innerHTML = `<div class="alert-row"><span>${budgetOpts.error ? 'โหลด Budget ไม่ได้: ' + esc(budgetOpts.error)
      : 'ยังไม่มีข้อมูล Budget — ' + (isAdmin() ? 'นำเข้าไฟล์ Budget ที่ สมาชิก &amp; ตั้งค่า ก่อน' : 'รอ Admin นำเข้าไฟล์ Budget')}</span>
      ${isAdmin() ? '<button class="btn sm" data-go="admin">ไปนำเข้า</button>' : ''}</div>`;
    return;
  }
  empty.innerHTML = '';
  body.hidden = false;
  if (state.pick.key && !budgetLine(state.pick.key)) state.pick.key = '';
  renderBudgetPicker();
  renderItemsCard();
  renderOrderSummary();
  renderFileList();
  if (document.activeElement !== $('#reqNote')) $('#reqNote').value = state.draft.note || '';
}

function lineYearLeft(line) {
  let left = 0;
  for (let m = 1; m <= 12; m++) left += budgetAvail(line, m).available - draftAmtFor(line.key, m);
  return left;
}

function renderBudgetPicker() {
  const words = $('#obSearch').value.trim().toLowerCase().split(/\s+/).filter(Boolean);
  const used = new Set(state.draft.lines.map(l => l.budgetKey));
  const list = budgetLines().filter(l => words.every(w => (l.label + ' ' + l.expenseGroup + ' ' + (l.mediaType || '')).toLowerCase().includes(w)));
  $('#obList').innerHTML = list.length ? list.map(l => {
    const left = lineYearLeft(l);
    const n = state.draft.lines.filter(x => x.budgetKey === l.key).length;
    return `<button class="bl-item ${l.key === state.pick.key ? 'sel' : ''}" data-pickkey="${esc(l.key)}">
      <span class="bl-main"><b>${esc(l.location)}</b> · ${esc(l.glCode)} ${esc(l.glName)}<span class="bl-sub">${esc(l.company)}${l.expenseGroup ? ' · ' + esc(l.expenseGroup) : ''}</span></span>
      <span class="bl-side">${n ? `<span class="badge">${n}</span>` : ''}<span class="num ${left < 0 ? 'over' : 'ok'}">${money(left)}</span><span class="bl-sub">เหลือทั้งปี</span></span>
      ${used.has(l.key) && l.key !== state.pick.key ? '<span class="bl-dot" title="มีรายการในคำขอนี้"></span>' : ''}</button>`;
  }).join('') : '<div class="empty-sm">ไม่พบ Budget ที่ค้นหา</div>';

  const line = budgetLine(state.pick.key);
  if (!line) {
    $('#obMonths').innerHTML = '<p class="hint">เลือก Budget ด้านบนก่อน</p>';
    $('#obPickInfo').innerHTML = '';
    return;
  }
  const cur = new Date().getMonth() + 1;
  $('#obMonths').innerHTML = MONTHS_1_12.map(m => {
    const b = budgetAvail(line, m);
    const left = b.available - draftAmtFor(line.key, m);
    const n = state.draft.lines.filter(x => x.budgetKey === line.key && Number(x.month) === m).length;
    return `<button class="mp ${Number(state.pick.month) === m ? 'sel' : ''} ${left < 0 ? 'over' : ''} ${b.plan ? '' : 'nobudget'}" data-pickmonth="${m}"
        title="${THAI_MONTH_FULL[m]}: งบ ${money(b.plan)} · เหลือ ${money(left)}">
      <span class="mp-m">${MONTH_SHORT[m - 1]}${m === cur ? ' •' : ''}</span><span class="mp-v">${fmtK(left)}</span>${n ? `<span class="mp-n">${n}</span>` : ''}</button>`;
  }).join('');
  const m = Number(state.pick.month) || cur;
  const b = budgetAvail(line, m);
  const mine = draftAmtFor(line.key, m);
  const after = b.available - mine;
  const remark = (line.months[m] || {}).remark;
  $('#obPickInfo').innerHTML = `
    <div class="pi-title">${esc(line.label)} · <b>${THAI_MONTH_FULL[m]} ${budgetOpts.year || ''}</b></div>
    <div class="pi-grid">
      <div><span>งบตาม BG</span><b>${money(b.plan)}</b></div>
      <div><span>ใช้จริงแล้ว</span><b>${money(b.actual)}</b></div>
      <div><span>คำขออื่นในระบบ</span><b>${money(b.reserved)}</b></div>
      <div><span>คำขอนี้</span><b>${money(mine)}</b></div>
      <div class="${after < 0 ? 'over' : 'ok'}"><span>คงเหลือหลังคำขอนี้</span><b>${money(after)}</b></div>
    </div>
    ${useBar({ plan: b.plan, actual: b.actual, committed: b.reserved, pending: mine, left: after })}
    ${remark ? `<div class="hint">Remark ตาม Budget: ${esc(remark)}</div>` : ''}
    ${after < 0 ? '<div class="over-note">งบเดือนนี้ไม่พอ — ส่งได้โดยติ๊ก "ขอ Over Budget" พร้อมเหตุผลในสรุปด้านขวา หรือเลือกเดือน/Budget อื่น</div>' : ''}`;
}

$('#obSearch').addEventListener('input', () => { if (hasBudget()) renderBudgetPicker(); });
$('#obList').addEventListener('click', e => {
  const b = e.target.closest('[data-pickkey]');
  if (!b) return;
  state.pick.key = b.dataset.pickkey;
  state.pick.month = Number(state.pick.month) || new Date().getMonth() + 1;
  save(); renderOrder();
});
$('#obMonths').addEventListener('click', e => {
  const b = e.target.closest('[data-pickmonth]');
  if (!b) return;
  state.pick.month = Number(b.dataset.pickmonth);
  save(); renderOrder();
});

function renderItemsCard() {
  const line = budgetLine(state.pick.key);
  const m = Number(state.pick.month);
  const ready = !!(line && m);
  $('#itemsCard').classList.toggle('disabled', !ready);
  $('#oiSearch').disabled = !ready;
  $('#itemsFor').innerHTML = ready ? `กำลังใส่รายการใน <b>${esc(line.label)}</b> · เดือน <b>${THAI_MONTH_FULL[m]}</b>`
    : 'เลือก Budget และเดือนในขั้นที่ 1–2 ก่อน แล้วค้นหาสินค้าที่จะสั่ง';

  // suggested items from Store that are not in this request yet
  const sug = ready && merged.rows.length ? merged.rows.filter(r => suggestion(r).qty > 0 && !draftHas(r.sku))
    .sort((a, b) => suggestion(b).qty * b.cost - suggestion(a).qty * a.cost).slice(0, 10) : [];
  $('#oiSuggest').innerHTML = sug.length ? `<span class="hint">แนะนำจาก Store:</span> ${sug.map(r =>
    `<button class="chip" data-addsku="${esc(r.sku)}" data-qty="${suggestion(r).qty}" title="${esc(r.sku)} · คงเหลือ ${fmt(r.balance)} ${esc(r.unit)}">+ ${esc(r.name)} <b>${fmt(suggestion(r).qty)}</b></button>`).join('')}` : '';

  const lines = ready ? state.draft.lines.filter(l => l.budgetKey === line.key && Number(l.month) === m) : [];
  const monthOpts = sel => MONTHS_1_12.map(x => `<option value="${x}" ${x === sel ? 'selected' : ''}>${MONTH_SHORT[x - 1]}</option>`).join('');
  $('#oiLines').innerHTML = !ready ? '' : !lines.length
    ? '<div class="empty-sm">ยังไม่มีรายการในงบนี้ — ค้นหาสินค้าด้านบน กดรายการแนะนำ หรือเพิ่มรายการที่ไม่มีใน Store</div>'
    : `<div class="table-wrap" style="max-height:none"><table class="grid order-lines"><thead><tr>
        <th>สินค้า</th><th class="n">Store คงเหลือ</th><th class="n">จำนวน</th><th>หน่วย</th><th class="n">ราคา/หน่วย</th><th class="n">รวม</th>
        <th>ขอบเขตการใช้งาน / หมายเหตุ</th><th>เดือน</th><th></th></tr></thead>
      <tbody>${lines.map(l => `<tr data-uid="${l.uid}">
        <td class="name"><b class="item-name">${esc(lineName(l))}</b><div class="pc">${l.custom ? 'รายการใหม่ (ไม่มีใน Store)' : esc(l.sku) + (l.pc ? ' · ' + esc(l.pc) : '')}</div></td>
        <td class="n">${lineBalance(l) === '' ? '–' : fmt(lineBalance(l))}</td>
        <td class="n"><input class="qty" type="number" min="0" step="1" data-f="qty" value="${esc(l.qty)}"></td>
        <td>${l.custom ? `<input class="unit-in" data-f="unit" value="${esc(l.unit || '')}">` : esc(lineUnit(l))}</td>
        <td class="n"><input class="qty" type="number" min="0" step="0.01" data-f="cost" value="${esc(lineCost(l))}" title="แก้ราคาได้ เช่น ตามใบเสนอราคา"></td>
        <td class="n"><b>${fmt(lineAmt(l), 2)}</b></td>
        <td><input class="note" data-f="note" value="${esc(l.note || '')}" placeholder="เช่น ใช้ล้างแกนมอเตอร์ป้าย…"></td>
        <td><select data-f="month" title="ย้ายไปเดือนอื่น">${monthOpts(Number(l.month))}</select></td>
        <td><button class="icon-btn" data-rmline title="เอาออก" aria-label="เอาออก">×</button></td>
      </tr>`).join('')}</tbody>
      <tfoot><tr><th colspan="5">รวมในงบนี้ (${lines.length} รายการ)</th><th class="n">${fmt(lines.reduce((a, l) => a + lineAmt(l), 0), 2)}</th><th colspan="3"></th></tr></tfoot></table></div>`;
}

// search the Store items
function renderItemSearch() {
  const q = $('#oiSearch').value.trim().toLowerCase();
  const box = $('#oiResults');
  if (!q) { box.hidden = true; box.innerHTML = ''; return; }
  const words = q.split(/\s+/);
  const hits = merged.rows.filter(r => words.every(w => (r.sku + ' ' + r.name + ' ' + (r.pc || '') + ' ' + (r.pcName || r.group || '')).toLowerCase().includes(w))).slice(0, 12);
  box.hidden = false;
  box.innerHTML = hits.length ? hits.map((r, i) => {
    const s = suggestion(r);
    return `<button class="sg-item ${i === 0 ? 'first' : ''}" data-addsku="${esc(r.sku)}" data-qty="${s.qty > 0 ? s.qty : 1}">
      <span><b>${esc(r.name)}</b><span class="pc">${esc(r.sku)} · ${esc(r.pc || r.group || '')} · ${money(r.cost)}/${esc(r.unit)}</span></span>
      <span class="sg-side">คงเหลือ ${fmt(r.balance)}${s.qty > 0 ? ` · แนะนำ <b>${fmt(s.qty)}</b>` : ''}${draftHas(r.sku) ? ' · <span class="ok">มีในคำขอแล้ว</span>' : ''}</span></button>`;
  }).join('') : `<div class="empty-sm">ไม่พบ "${esc(q)}" ใน Store — ${merged.rows.length ? 'ใช้ "+ เพิ่มรายการที่ไม่มีใน Store"' : 'ยังไม่มีข้อมูล Store'}</div>`;
}
$('#oiSearch').addEventListener('input', renderItemSearch);
$('#oiSearch').addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); const f = $('#oiResults [data-addsku]'); if (f) f.click(); }
  if (e.key === 'Escape') { e.target.value = ''; renderItemSearch(); }
});
document.addEventListener('click', e => {
  const add = e.target.closest('#tab-order [data-addsku]');
  if (!add) return;
  const r = bySku.get(add.dataset.addsku);
  if (!r) return;
  addLine(snapshot(r), Number(add.dataset.qty) || 1);
  $('#oiSearch').value = ''; renderItemSearch();
  save(); renderOrder(); renderCards();
  toast(`เพิ่ม ${r.name} แล้ว`);
  $('#oiSearch').focus();
});
$('#oiNewForm').addEventListener('submit', e => {
  e.preventDefault();
  const f = e.target;
  addLine({ sku: '', name: f.name.value.trim(), unit: f.unit.value.trim(), cost: Number(f.cost.value) || 0, custom: true }, Number(f.qty.value) || 1);
  f.reset(); save(); renderOrder(); renderCards();
});
$('#oiLines').addEventListener('change', e => {
  const tr = e.target.closest('tr[data-uid]');
  const l = tr && state.draft.lines.find(x => x.uid === tr.dataset.uid);
  if (!l || !e.target.dataset.f) return;
  const f = e.target.dataset.f;
  if (f === 'qty') l.qty = Math.max(0, Math.round(Number(e.target.value) || 0));
  else if (f === 'cost') l.cost = Math.max(0, Number(e.target.value) || 0);
  else if (f === 'month') l.month = Number(e.target.value);
  else l[f] = e.target.value;
  save(); renderOrder(); renderCards();
});
$('#oiLines').addEventListener('click', e => {
  if (!e.target.matches('[data-rmline]')) return;
  const id = e.target.closest('tr').dataset.uid;
  state.draft.lines = state.draft.lines.filter(l => l.uid !== id);
  save(); renderOrder(); renderCards();
});

function renderOrderSummary() {
  const groups = draftGroups();
  const total = groups.reduce((a, x) => a + x.amount, 0);
  const over = state.draft.over || (state.draft.over = {});
  if (!groups.length) {
    $('#oSummary').innerHTML = '<div class="empty-sm">ยังไม่มีรายการ<br>เลือก Budget + เดือน แล้วใส่สินค้าที่จะสั่ง</div>';
    return;
  }
  $('#oSummary').innerHTML = groups.map(x => {
    if (!x.line || !x.month) {
      return `<div class="bline warn"><div class="bline-head"><b>${x.n} รายการยังไม่ได้เลือก Budget</b><span class="num">${money(x.amount)}</span></div>
        <button class="btn sm" data-assign="${esc(x.k)}" ${budgetLine(state.pick.key) ? '' : 'disabled'}>ใส่ใน Budget ที่เลือกอยู่</button></div>`;
    }
    const ob = over[x.k] || {};
    const cur = x.key === state.pick.key && x.month === Number(state.pick.month);
    return `<div class="bline ${x.over ? 'is-over' : ''} ${cur ? 'cur' : ''}">
      <button class="bline-head link" data-gopick="${esc(x.k)}" title="เปิดงบนี้"><b>${esc(x.line.location)} · ${esc(x.line.glCode)}</b>
        <span class="hint">${THAI_MONTH_FULL[x.month]} · ${x.n} รายการ</span></button>
      <div class="row"><span>คำขอนี้</span><b>${money(x.amount)}</b></div>
      <div class="row"><span>คงเหลือหลังส่ง</span><b class="${x.over ? 'over' : 'ok'}">${money(x.after)}</b></div>
      ${x.over ? `<label class="over-req"><input type="checkbox" data-over-on="${esc(x.k)}" ${ob.on ? 'checked' : ''}> ขอ Over Budget (เกิน ${money(-x.after)})</label>
        <input class="over-reason" data-over-reason="${esc(x.k)}" value="${esc(ob.reason || '')}" placeholder="เหตุผลที่ขอเกินงบ (จำเป็น)" ${ob.on ? '' : 'hidden'}>` : ''}
    </div>`;
  }).join('') + `<div class="row total"><span>รวม ${state.draft.lines.length} รายการ (ไม่รวม VAT)</span><b>${money(total)}</b></div>`;
}

$('#oSummary').addEventListener('click', e => {
  const go = e.target.closest('[data-gopick]');
  if (go) { state.pick = splitKM(go.dataset.gopick); save(); renderOrder(); $('#itemsCard').scrollIntoView({ block: 'start', behavior: 'smooth' }); return; }
  const as = e.target.closest('[data-assign]');
  if (as && budgetLine(state.pick.key)) {
    state.draft.lines.forEach(l => { if (!budgetLine(l.budgetKey) || !Number(l.month)) { l.budgetKey = state.pick.key; l.month = Number(state.pick.month); } });
    save(); renderOrder();
  }
});
$('#oSummary').addEventListener('change', e => {
  const k = e.target.dataset.overOn;
  if (k == null) return;
  state.draft.over[k] = { ...(state.draft.over[k] || {}), on: e.target.checked };
  save(); renderOrderSummary();
  if (e.target.checked) { const inp = $('#oSummary').querySelector(`[data-over-reason="${CSS.escape(k)}"]`); if (inp) inp.focus(); }
});
$('#oSummary').addEventListener('input', e => {
  const k = e.target.dataset.overReason;
  if (k == null) return;
  state.draft.over[k] = { ...(state.draft.over[k] || {}), on: true, reason: e.target.value };
  save();
});
$('#reqNote').addEventListener('input', e => { state.draft.note = e.target.value; save(); });

// ---- quotations
const MAX_FILE = 5 * 1024 * 1024, MAX_TOTAL = 10 * 1024 * 1024, MAX_FILES = 5;
const kb = n => n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
function renderFileList() {
  const keep = state.draft.keepFiles || [];
  $('#oFileList').innerHTML = keep.map(f => `<li><span class="file-chip" data-file="${esc(f.id)}">📎 ${esc(f.name)}</span><span class="hint">${kb(f.size)} · แนบไว้แล้ว</span>
      <button class="icon-btn" data-rmkeep="${esc(f.id)}" title="เอาออก" aria-label="เอาออก">×</button></li>`).join('') +
    draftFiles.map((f, i) => `<li><span class="file-chip new">📎 ${esc(f.name)}</span><span class="hint">${kb(f.size)}</span>
      <button class="icon-btn" data-rmfile="${i}" title="เอาออก" aria-label="เอาออก">×</button></li>`).join('');
}
const readDataUrl = file => new Promise((resolve, reject) => { const fr = new FileReader(); fr.onload = () => resolve(fr.result); fr.onerror = () => reject(fr.error); fr.readAsDataURL(file); });
$('#oFiles').addEventListener('change', async e => {
  const files = [...e.target.files];
  e.target.value = '';
  for (const f of files) {
    const count = draftFiles.length + (state.draft.keepFiles || []).length;
    const total = draftFiles.reduce((a, x) => a + x.size, 0) + (state.draft.keepFiles || []).reduce((a, x) => a + x.size, 0);
    if (count >= MAX_FILES) { toast(`แนบได้สูงสุด ${MAX_FILES} ไฟล์`); break; }
    if (f.size > MAX_FILE) { toast(`${f.name} ใหญ่เกิน 5 MB`); continue; }
    if (total + f.size > MAX_TOTAL) { toast('ไฟล์แนบรวมเกิน 10 MB'); break; }
    draftFiles.push({ name: f.name, type: f.type || 'application/octet-stream', size: f.size, data: await readDataUrl(f) });
  }
  renderFileList();
});
$('#oFileList').addEventListener('click', e => {
  const rm = e.target.closest('[data-rmfile]');
  if (rm) { draftFiles.splice(Number(rm.dataset.rmfile), 1); renderFileList(); return; }
  const rk = e.target.closest('[data-rmkeep]');
  if (rk) { state.draft.keepFiles = state.draft.keepFiles.filter(f => f.id !== rk.dataset.rmkeep); save(); renderFileList(); }
});

function resetDraft() {
  state.draft = { lines: [], note: '', over: {}, editId: null, keepFiles: [] };
  draftFiles = [];
  save();
}

$('#btnSubmitReq').addEventListener('click', async () => {
  const lines = state.draft.lines.filter(l => Number(l.qty) > 0);
  if (!lines.length) return toast('ยังไม่มีรายการ หรือจำนวนเป็น 0');
  const groups = draftGroups().filter(x => x.amount > 0 || x.n);
  if (groups.some(x => !x.line || !x.month)) return toast('ยังมีรายการที่ไม่ได้เลือก Budget / เดือน');
  const over = groups.filter(x => x.over);
  const notAsked = over.filter(x => !(state.draft.over[x.k] && state.draft.over[x.k].on && String(state.draft.over[x.k].reason || '').trim()));
  if (notAsked.length) {
    renderOrderSummary();
    return toast(`งบไม่พอ ${notAsked.length} Budget — ติ๊ก "ขอ Over Budget" และใส่เหตุผล หรือเปลี่ยน Budget / เดือน / จำนวน`);
  }
  const total = groups.reduce((a, x) => a + x.amount, 0);
  const nFiles = draftFiles.length + (state.draft.keepFiles || []).length;
  const ed = state.draft.editId;
  if (!confirm(`${ed ? 'ส่งคำขอ ' + ed + ' ที่แก้ไขแล้ว' : 'ส่งคำขอ'} ${lines.length} รายการ รวม ${money(total)} (ไม่รวม VAT)\n` +
    groups.map(x => `• ${x.line.label} · ${THAI_MONTH_FULL[x.month]} · ${money(x.amount)}${x.over ? ' (ขอ Over Budget)' : ''}`).join('\n') +
    (nFiles ? `\nแนบใบเสนอราคา ${nFiles} ไฟล์` : '') + '\n\nส่งให้ Admin ตรวจ — ยืนยัน?')) return;
  const overReasons = {};
  over.forEach(x => { overReasons[x.k] = String(state.draft.over[x.k].reason).trim(); });
  $('#btnSubmitReq').disabled = true;
  $('#submitHint').textContent = nFiles ? 'กำลังส่ง (อัปโหลดใบเสนอราคา)…' : 'กำลังส่ง…';
  try {
    const res = await gasCall('submitOrderRequest', {
      editId: ed || '', keepFileIds: (state.draft.keepFiles || []).map(f => f.id),
      note: state.draft.note || '', overReasons,
      files: draftFiles.map(f => ({ name: f.name, type: f.type, data: f.data })),
      items: lines.map(l => ({ sku: l.custom ? '' : l.sku, name: lineName(l), pc: l.pc || (lineStore(l) || {}).pc || '', pcName: l.pcName || '',
        unit: lineUnit(l), qty: Number(l.qty), unitCost: lineCost(l), note: l.note || '', budgetKey: l.budgetKey, month: Number(l.month),
        storeBalance: lineBalance(l) }))
    });
    resetDraft();
    toast(`${res.resubmitted ? 'ส่งคำขอ ' + res.id + ' ที่แก้ไขแล้ว' : 'ส่งคำขอ ' + res.id + ' แล้ว'}${res.overBudget ? ' (มีขอ Over Budget)' : ''} — Admin ได้รับอีเมลแจ้ง`);
    renderCards(); refreshBudget();
    $('#reqFilter').value = '';
    focusReq = res.id;
    await loadRequests(true);
    goTab('requests');
  } catch (e) {
    toast('ส่งไม่สำเร็จ: ' + errMsg(e));
  } finally {
    $('#btnSubmitReq').disabled = false;
    $('#submitHint').textContent = '';
  }
});

$('#btnCancelEdit').addEventListener('click', () => {
  if (!confirm(`ยกเลิกการแก้ไข ${state.draft.editId}?\nรายการในหน้านี้จะถูกล้าง — คำขอยังอยู่ในสถานะ "รอผู้ขอแก้ไข" กดแก้ไขใหม่ได้ภายหลัง`)) return;
  resetDraft(); renderOrder(); renderCards();
});
$('#btnClearCart').addEventListener('click', () => {
  if (!state.draft.lines.length && !draftFiles.length) return;
  if (!confirm('ล้างรายการในคำขอนี้ทั้งหมด?')) return;
  const ed = state.draft.editId, keep = state.draft.keepFiles;
  resetDraft();
  if (ed) { state.draft.editId = ed; state.draft.keepFiles = keep; save(); }
  renderOrder(); renderCards();
});

/** Loads a request (called back / sent back) into the order page to edit and resubmit. */
function loadIntoDraft(r) {
  if (state.draft.lines.length && state.draft.editId !== r.request_id &&
    !confirm('มีรายการที่ยังไม่ได้ส่งอยู่ในหน้าสั่งซื้อ — แทนที่ด้วยคำขอ ' + r.request_id + '?')) return false;
  const over = {};
  const lines = r.items.filter(i => i.qty > 0).map(i => {
    const key = i.budget_key || r.budget_key, month = Number(i.budget_month || r.budget_month) || 0;
    if (i.over_budget) over[key + '|' + month] = { on: true, reason: i.over_reason };
    return { uid: uid(), sku: i.sku, name: i.name, unit: i.unit, cost: i.unit_cost, qty: i.qty, note: i.note, budgetKey: key, month,
      pc: i.pc_code, pcName: i.pc_name, balance: i.store_balance, custom: !bySku.has(i.sku) };
  });
  state.draft = { lines, note: r.note || '', over, editId: r.request_id, keepFiles: (r.files || []).slice() };
  draftFiles = [];
  if (lines.length) state.pick = { key: lines[0].budgetKey, month: lines[0].month };
  save();
  return true;
}

function cartSheetRows() {
  return state.draft.lines.map(l => {
    const line = budgetLine(l.budgetKey);
    return { 'Budget': line ? line.label : '', 'เดือน': THAI_MONTH_FULL[Number(l.month)] || '', 'PC': l.pc || '', 'รหัสสินค้า': l.custom ? '' : l.sku,
      'สินค้า': lineName(l), 'หน่วย': lineUnit(l), 'Store คงเหลือ': lineBalance(l), 'จำนวนสั่ง': l.qty, 'ราคา/หน่วย': lineCost(l),
      'รวมเงิน': Math.round(lineAmt(l) * 100) / 100, 'ขอบเขตการใช้งาน': l.note || '' };
  });
}

const stamp = () => new Date().toISOString().slice(0, 10);

$('#btnExportCart').addEventListener('click', () => {
  const rows = cartSheetRows();
  if (!rows.length) return toast('ยังไม่มีรายการ');
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'รายการสั่งซื้อ');
  XLSX.writeFile(wb, `รายการสั่งซื้อ_${stamp()}.xlsx`);
});

// ------------------------------------------------------------ export all

$('#btnExportAll').addEventListener('click', () => {
  const cols = mergedColumns();
  const toObj = r => Object.fromEntries(cols.map(c => [c[0], c[1](r)]));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(merged.rows.map(toObj)), MERGED_SHEET);
  const cmp = merged.rows.filter(r => r.orderStore > 0 || r.orderManual > 0 || r.orderActual > 0).map(r => ({
    'รหัสสินค้า': r.sku, 'สินค้า': r.name, 'PC': r.pc || r.group, 'ประเภท': r.category, 'คงเหลือ': r.balance,
    'Store แจ้งสั่ง': r.orderStore, 'MIN/MAX เดิม': r.orderManual, 'ใช้จริง': r.orderActual,
    'ต่างกัน (Store - ใช้จริง)': r.orderStore - r.orderActual,
    'มูลค่า Store': Math.round(r.orderStore * r.cost), 'มูลค่าใช้จริง': Math.round(r.orderActual * r.cost),
    'ข้อสังเกต': r.flags.join(' | ')
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cmp), 'เทียบยอดสั่ง');
  const cart = cartSheetRows();
  if (cart.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cart), 'รายการที่เลือกสั่งซื้อ');
  const src = SLOTS.flatMap(s => (state.files[s.type] || []).map(f => ({ 'ช่อง': s.title, 'ไฟล์': f.name, 'แถว': f.count })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(src), 'ไฟล์ต้นทาง');
  XLSX.writeFile(wb, `Store_รวมข้อมูล_${stamp()}.xlsx`);
});

$('#btnReset').addEventListener('click', () => {
  if (!confirm('ล้างไฟล์ที่นำเข้าและรายการที่เลือกทั้งหมด?')) return;
  state = { files: {}, opt: { ...DEFAULTS }, cart: {}, qtySource: 'auto', draft: { lines: [], note: '', over: {}, editId: null, keepFiles: [] }, pick: { key: '', month: new Date().getMonth() + 1 } };
  draftFiles = [];
  save(); recompute();
});

// ------------------------------------------------------------ AI assistant

$('.ai-key').hidden = SERVER; // Apps Script uses ANTHROPIC_API_KEY from Script Properties
try { $('#aiKey').value = localStorage.getItem(AI_KEY) || ''; } catch { /* ignore */ }
$('#aiKey').addEventListener('change', e => { try { localStorage.setItem(AI_KEY, e.target.value.trim()); } catch { /* ignore */ } });
$('#aiKeyClear').addEventListener('click', () => { $('#aiKey').value = ''; try { localStorage.removeItem(AI_KEY); } catch { /* ignore */ } });
document.querySelectorAll('[data-ask]').forEach(b => b.addEventListener('click', () => { $('#aiQ').value = b.dataset.ask; askAI(); }));
$('#aiAsk').addEventListener('click', askAI);

function aiContext() {
  const s = summarize(merged.rows);
  const line = r => [r.sku, r.name, r.pc || r.group, r.category, r.balance, r.unit, r.cost,
    r.manual ? r.manual.avg : '', Math.round(r.avgActual * 10) / 10, r.orderStore, r.orderManual, r.orderActual,
    draftLineOf(r.sku)?.qty ?? '', r.flags.join('; ')].join('\t');
  const head = 'sku\tname\tpc\tcategory\tbalance\tunit\tunit_cost\tavg_manual\tavg_actual\torder_store\torder_minmax_old\torder_actual\tselected_qty\tflags';
  const interesting = merged.rows
    .filter(r => r.flags.length || r.orderStore || r.orderActual || r.orderManual || draftHas(r.sku))
    .sort((a, b) => Math.max(b.orderStore, b.orderActual, b.orderManual) * b.cost - Math.max(a.orderStore, a.orderActual, a.orderManual) * a.cost)
    .slice(0, 250);
  return `เดือนที่มีข้อมูลการใช้: ${merged.months.map(monthLabel).join(', ')}
สรุป: ${JSON.stringify(s)}
สูตร: MIN=${state.opt.factorMin}×(avg/30)×lead, MAX=${state.opt.factorMax}×(avg/30)×lead, เบิกประจำ=ใช้≥${state.opt.regularMinMonths} เดือน
รายการ (${interesting.length} จาก ${merged.rows.length}, เรียงตามมูลค่าที่ต้องสั่ง):
${head}
${interesting.map(line).join('\n')}`;
}

const AI_SYSTEM = 'คุณคือผู้ช่วยฝ่ายจัดซื้อ/Store ของ Plan B Media ทีม Static Media (PG44) ตอบเป็นภาษาไทย กระชับ ' +
  'ใช้เฉพาะตัวเลขที่ให้มา ห้ามแต่งตัวเลขเพิ่ม ถ้าข้อมูลไม่พอให้บอกตรงๆ ' +
  'order_store = ยอดที่ระบบ Store แจ้ง, order_minmax_old = จาก MIN/MAX ที่กรอกเอง, order_actual = จาก MIN/MAX ที่คำนวณจากยอดเบิกจริง. ' +
  'เบิกตามงาน = ไม่ควรตั้งจุดสั่งซื้อตายตัว ต้องถามทีมก่อน. คำแนะนำของคุณเป็นข้อเสนอให้คนตัดสินใจเท่านั้น';

/** Browser-direct call (GitHub Pages): uses the viewer's own key. */
async function askClaudeBrowser(key, userContent) {
  const { default: Anthropic } = await import('https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk/+esm');
  const client = new Anthropic({ apiKey: key, dangerouslyAllowBrowser: true });
  const res = await client.beta.messages.create({
    model: 'claude-opus-5',
    max_tokens: 16000,
    betas: ['server-side-fallback-2026-07-01'],
    fallbacks: 'default',
    system: AI_SYSTEM,
    messages: [{ role: 'user', content: userContent }]
  });
  if (res.stop_reason === 'refusal') throw new Error('AI ปฏิเสธคำขอนี้');
  return res.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
}

async function askAI() {
  const key = $('#aiKey').value.trim();
  const q = $('#aiQ').value.trim();
  if (!SERVER && !key) return toast('ใส่ Anthropic API key ก่อน');
  if (!q) return toast('พิมพ์คำถามก่อน');
  if (!merged.rows.length) return toast('ยังไม่มีข้อมูล');
  $('#aiAsk').disabled = true; $('#aiStatus').textContent = 'กำลังวิเคราะห์…'; $('#aiOut').textContent = '';
  const userContent = `ข้อมูลรอบนี้:\n${aiContext()}\n\nคำถาม: ${q}`;
  try {
    // Apps Script: key lives in Script Properties, request goes out via UrlFetchApp (src/Ui.js)
    $('#aiOut').textContent = SERVER
      ? await gasCall('askClaudeFromUi', AI_SYSTEM, userContent)
      : await askClaudeBrowser(key, userContent);
    $('#aiStatus').textContent = '';
  } catch (e) {
    console.error(e);
    $('#aiStatus').textContent = 'ผิดพลาด: ' + (e.status === 401 ? 'API key ไม่ถูกต้อง' : e.message);
  } finally {
    $('#aiAsk').disabled = false;
  }
}

// ------------------------------------------------------------ membership: ID + password (server: src/Members.js, src/Api.js)

let me = null;          // { id, name, email, role, status, mustChange } from getSessionInfo()/login()
let budgetOpts = null;  // { year, lines:[{key,label,months:{m:{plan,actual}}}], reserved:{"key|m":amt} }
let requests = [];

const isMember = () => !!me && me.status === 'active' && !me.mustChange;
const isAdmin = () => isMember() && me.role === 'admin';
const errMsg = e => (e && e.network) ? 'เชื่อมต่อ Apps Script ไม่ได้ (ยังไม่ได้ Deploy เวอร์ชันใหม่ หรือสิทธิ์ไม่ใช่ "ทุกคน")' : String((e && e.message) || e).replace(/^Error:\s*/, '').replace(/^SESSION_EXPIRED:\s*/, '');

function applySession(info) {
  // identity used to match Master PC owners / managers: email when set, else the ID (same rule as the server)
  const prevId = me && me.id;
  me = info && info.id ? { ...info, email: String(info.email || info.id).toLowerCase() } : info;
  members = null; masterPc = []; requests = []; appSettings = null; // never show the previous user's data
  editingReq = null; composing = null; reqFilterSet = false; selectedReqs.clear(); visibleReqs = [];
  const active = isMember();
  const admin = isAdmin();
  // a different person on this browser starts with an empty request
  if (active && prevId && prevId !== me.id) resetDraft();
  if (active) { try { const owner = localStorage.getItem(DRAFT_OWNER_KEY); if (owner && owner !== me.id) resetDraft(); localStorage.setItem(DRAFT_OWNER_KEY, me.id); } catch { /* private mode */ } }
  $('#whoami').hidden = !active;
  if (active) {
    $('#whoami').innerHTML = `<b>${esc(me.name || me.id)}</b> <span class="hint">(${esc(me.id)})</span>
      <button class="link-btn" data-acct="profile">บัญชีของฉัน</button> <button class="link-btn" data-acct="logout">ออกจากระบบ</button>`;
  }
  // Admin and User see different menus and home pages
  $('#rolePill').hidden = !active;
  $('#rolePill').textContent = admin ? 'Admin · ผู้ดูแลระบบ' : 'User · ผู้ใช้งาน';
  $('#rolePill').className = 'role-pill ' + (admin ? 'admin' : 'user');
  document.body.dataset.role = active ? (admin ? 'admin' : 'user') : '';
  document.querySelectorAll('#sidenav [data-role]').forEach(el => {
    const r = el.dataset.role;
    el.hidden = !(r === 'all' || (r === 'admin' ? admin : !admin));
  });
  document.querySelectorAll('#sidenav [data-admin]').forEach(el => { el.textContent = admin ? el.dataset.admin : el.dataset.user; });
  const cur = document.querySelector('.tab.active');
  if (!cur || cur.hidden) { activeTab = 'home'; document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab === 'home' && !x.hidden)); document.querySelectorAll('.tabpane').forEach(p => (p.hidden = p.id !== 'tab-home')); }
  $('#roundBox').hidden = !admin;
  renderDemoBar();
  renderGate();
  const blocked = SERVER && !active;
  $('#sidenav').hidden = blocked;
  $('.top-actions').querySelectorAll('.btn').forEach(b => { b.hidden = blocked; });
  if (blocked) { $('#summary').hidden = true; $('#work').hidden = true; return; }
  recompute();
  if (active) {
    loadServerData(); refreshBudget(); loadMasterPc();
    if (admin) loadSettings();
    loadRequests(true).then(openDeepLink);
  }
}

// Links in emails are <url>?req=REQ-…
function openDeepLink() {
  const go = id => {
    if (!id) return;
    $('#reqFilter').value = '';
    focusReq = id;
    document.querySelector('.tab[data-tab="requests"]').click();
  };
  if (GAS && google.script.url) google.script.url.getLocation(loc => go(loc && loc.parameter && loc.parameter.req));
  else go(new URLSearchParams(location.search).get('req'));
}

function renderGate() {
  const g = $('#gate');
  if (!SERVER || isMember()) { g.hidden = true; return; }
  g.hidden = false;
  if (me && me.status === 'loading') { g.innerHTML = '<div class="gate"><p class="hint">กำลังตรวจสอบการเข้าสู่ระบบ…</p></div>'; return; }

  if (me && me.mustChange) {
    g.innerHTML = `<div class="gate"><h2>ตั้งรหัสผ่านใหม่</h2>
      <p class="hint">${esc(me.id)} ใช้รหัสผ่านเริ่มต้นอยู่ — ตั้งรหัสใหม่ (อย่างน้อย 8 ตัวอักษร) ก่อนใช้งาน</p>
      <form id="pwForm" class="gate-form">
        <input name="old" type="password" placeholder="รหัสผ่านปัจจุบัน" autocomplete="current-password" required>
        <input name="pw1" type="password" placeholder="รหัสผ่านใหม่" autocomplete="new-password" minlength="8" required>
        <input name="pw2" type="password" placeholder="ยืนยันรหัสผ่านใหม่" autocomplete="new-password" minlength="8" required>
        <button class="btn">บันทึกรหัสผ่าน</button> <button type="button" class="btn ghost" data-acct="logout">ออกจากระบบ</button>
      </form></div>`;
    $('#pwForm').addEventListener('submit', async e => {
      e.preventDefault();
      const f = e.target;
      if (f.pw1.value !== f.pw2.value) return toast('รหัสผ่านใหม่ไม่ตรงกัน');
      try { applySession(await gasCall('changePassword', f.old.value, f.pw1.value)); toast('เปลี่ยนรหัสผ่านแล้ว'); }
      catch (err) { toast(errMsg(err)); }
    });
    return;
  }

  const registering = g.dataset.mode === 'register';
  g.innerHTML = registering
    ? `<div class="gate"><h2>สมัครสมาชิก</h2>
        <p class="hint">สมัครแล้วรอ Admin อนุมัติ · อีเมลใช้รับแจ้งเตือนและจับคู่ผู้รับผิดชอบ PC</p>
        <form id="regForm" class="gate-form">
          <input name="username" placeholder="ID (a-z 0-9 . _ -)" autocomplete="username" required>
          <input name="password" type="password" placeholder="รหัสผ่าน (อย่างน้อย 8 ตัว)" autocomplete="new-password" minlength="8" required>
          <input name="name" placeholder="ชื่อ-นามสกุล / ทีม" required>
          <input name="email" type="email" placeholder="อีเมล @planbmedia.co.th">
          <button class="btn">สมัคร</button> <button type="button" class="btn ghost" data-gate="login">มีบัญชีแล้ว — เข้าสู่ระบบ</button>
        </form></div>`
    : `<div class="gate"><h2>เข้าสู่ระบบ</h2>
        ${DEMO ? `<div class="demo-login"><p><b>โหมดทดลอง</b> — เลือกบทบาทเพื่อลองใช้งาน (ข้อมูลตัวอย่าง ไม่กระทบระบบจริง)</p>
          <div class="demo-roles">
            <button type="button" class="role-card admin" data-demo="admin"><b>เข้าเป็น Admin</b><span>ตรวจคำขอ · แก้ไข/ลบ · ส่งเมลรายการ + Budget</span></button>
            <button type="button" class="role-card user" data-demo="user"><b>เข้าเป็น User</b><span>เลือก Budget · ใส่รายการ · แนบใบเสนอราคา · ส่งให้ Admin</span></button>
          </div>
          <p class="hint">หรือใส่ ID เอง: admin / user / user2 / head · รหัสผ่าน demo1234</p></div>` : ''}
        <form id="loginForm" class="gate-form">
          <input name="username" placeholder="ID" autocomplete="username" required ${DEMO ? '' : 'autofocus'}>
          <input name="password" type="password" placeholder="รหัสผ่าน" autocomplete="current-password" required>
          <button class="btn">เข้าสู่ระบบ</button> <button type="button" class="btn ghost" data-gate="register">สมัครสมาชิก</button>
        </form>
        <p id="gateErr" class="hint over"></p>
        ${!DEMO && !GAS ? '<p class="hint try-demo">อยากลองก่อน? <a href="?demo=1">ทดลองใช้งานด้วยข้อมูลตัวอย่าง</a> (Admin และ User — ไม่กระทบระบบจริง)</p>' : ''}</div>`;
  const f = $('#loginForm') || $('#regForm');
  f.addEventListener('submit', async e => {
    e.preventDefault();
    const btn = f.querySelector('button.btn:not(.ghost)');
    btn.disabled = true;
    try {
      if (registering) {
        toast(await gasCall('register', { username: f.username.value, password: f.password.value, name: f.name.value, email: f.email.value }));
        g.dataset.mode = 'login'; renderGate();
      } else {
        const res = await gasCall('login', f.username.value, f.password.value);
        setToken(res.token);
        applySession(res.user);
      }
    } catch (err) {
      btn.disabled = false;
      if (err && err.network) { showConnectHelp(); return; }
      if ($('#gateErr')) $('#gateErr').textContent = errMsg(err);
      else toast(errMsg(err));
    }
  });
}

/** Shown when the Apps Script backend can't be reached from a GitHub/Vercel page. */
function showConnectHelp() {
  const box = $('#gateErr');
  if (!box) return;
  const health = APPS_SCRIPT_URL + '?health=1';
  box.classList.remove('over');
  box.innerHTML = `<div class="connect-help">
    <b class="over">เชื่อมต่อ Apps Script ไม่ได้</b> — ลิงก์ backend ยังไม่ได้รันโค้ดเวอร์ชันนี้ มักเกิดจาก:
    <ol>
      <li>บันทึกโค้ดแล้วแต่ยังไม่ได้ Deploy ใหม่: Apps Script &gt; การทำให้ใช้งานได้ &gt; จัดการการทำให้ใช้งานได้ &gt; ✏ &gt; เวอร์ชัน: <b>เวอร์ชันใหม่</b> &gt; ทำให้ใช้งานได้</li>
      <li>ผู้มีสิทธิ์เข้าถึงไม่ใช่ <b>"ทุกคน"</b> (Anyone) · เรียกใช้ในฐานะ: <b>ฉัน</b></li>
      <li>สร้าง deployment ใหม่ทำให้ลิงก์ /exec เปลี่ยน — ใส่ลิงก์ใหม่ด้านล่าง</li>
    </ol>
    ตรวจได้ที่ <a href="${esc(health)}" target="_blank" rel="noopener">${esc(health)}</a> — ต้องเห็น <code>"app":"store-reorder-ai"</code>
    <form id="apiForm" class="inline-form" style="margin-top:8px">
      <input name="url" value="${esc(APPS_SCRIPT_URL)}" style="flex:1 1 360px" placeholder="https://script.google.com/macros/s/…/exec">
      <button class="btn sm">บันทึกลิงก์</button>
    </form></div>`;
  $('#apiForm').addEventListener('submit', e => {
    e.preventDefault();
    const url = e.target.url.value.trim();
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) return toast('ลิงก์ต้องเป็น https://script.google.com/macros/s/…/exec');
    APPS_SCRIPT_URL = url;
    try { url === DEFAULT_APPS_SCRIPT_URL ? localStorage.removeItem(API_URL_KEY) : localStorage.setItem(API_URL_KEY, url); } catch { /* ignore */ }
    toast('บันทึกลิงก์แล้ว ลองเข้าสู่ระบบอีกครั้ง');
  });
}

document.addEventListener('click', async e => {
  const gateBtn = e.target.closest('[data-gate]');
  if (gateBtn) { $('#gate').dataset.mode = gateBtn.dataset.gate; renderGate(); return; }
  const acct = e.target.closest('[data-acct]');
  if (!acct) return;
  if (acct.dataset.acct === 'logout') {
    try { await gasCall('logout'); } catch { /* token may already be gone */ }
    setToken(''); $('#gate').dataset.mode = 'login'; applySession(null);
  } else if (acct.dataset.acct === 'profile') {
    const email = prompt('อีเมลของคุณ (ใช้รับแจ้งเตือน / จับคู่ผู้รับผิดชอบ PC):', me.id === me.email ? '' : me.email);
    if (email === null) return;
    const change = confirm('ต้องการเปลี่ยนรหัสผ่านด้วยหรือไม่?');
    try {
      await gasCall('updateProfile', { name: me.name, email });
      if (change) {
        const old = prompt('รหัสผ่านปัจจุบัน:'); if (old === null) return;
        const pw = prompt('รหัสผ่านใหม่ (อย่างน้อย 8 ตัว):'); if (!pw) return;
        await gasCall('changePassword', old, pw);
      }
      applySession(await gasCall('getSessionInfo'));
      toast('บันทึกแล้ว');
    } catch (err) { toast(errMsg(err)); }
  }
});

// ------------------------------------------------------------ demo / trial mode (web/demo.js)

function renderDemoBar() {
  $('#demoBar').hidden = !DEMO;
  document.body.classList.toggle('demo', DEMO);
  if (!DEMO) return;
  $('#demoMailCount').textContent = window.DEMO.outbox().length;
  document.querySelectorAll('#demoBar [data-demo="admin"], #demoBar [data-demo="user"]')
    .forEach(b => b.classList.toggle('on', !!me && me.id === b.dataset.demo));
}

async function demoLogin(username) {
  const acct = window.DEMO.accounts.find(a => a.username === username);
  if (!acct) return;
  try { if (token) await gasCall('logout'); } catch { /* already signed out */ }
  setToken('');
  try {
    const res = await gasCall('login', acct.username, acct.password);
    setToken(res.token);
    applySession(res.user);
    toast(`เข้าสู่ระบบเป็น ${acct.name}`);
  } catch (e) { toast(errMsg(e)); }
}

function renderOutbox(sel = 0) {
  const mails = window.DEMO.outbox();
  $('#demoMailList').innerHTML = mails.length ? mails.map((m, i) => `<li class="${i === sel ? 'sel' : ''}" data-mail="${i}">
      <b>${esc(m.subject)}</b><span class="hint">ถึง ${esc(m.to)}${m.cc ? ' · สำเนา ' + esc(m.cc) : ''}</span>
      <span class="hint">${esc(new Date(m.at).toLocaleString('th-TH'))}${m.attachments.length ? ' · 📎 ' + m.attachments.length : ''}</span></li>`).join('')
    : '<li class="empty-sm">ยังไม่มีอีเมล — ลองส่งคำขอ แล้วให้ Admin กด "ส่งเมล"</li>';
  const m = mails[sel];
  $('#demoMailView').innerHTML = m ? `<div class="mail-meta"><div><b>เรื่อง:</b> ${esc(m.subject)}</div><div><b>ถึง:</b> ${esc(m.to)}</div>
      ${m.cc ? `<div><b>สำเนา:</b> ${esc(m.cc)}</div>` : ''}
      ${m.attachments.length ? `<div class="mail-att"><b>ไฟล์แนบ:</b> ${m.attachments.map(a => a.data
        ? `<a href="data:${esc(a.type)};base64,${a.data}" download="${esc(a.name)}">📎 ${esc(a.name)}</a>` : `<span>📎 ${esc(a.name)}</span>`).join(' ')}</div>` : ''}</div>
    <iframe class="mail-preview" title="อีเมล" sandbox srcdoc="${esc(m.htmlBody)}"></iframe>` : '';
}

document.addEventListener('click', e => {
  const b = e.target.closest('[data-demo]');
  if (b && DEMO) {
    const a = b.dataset.demo;
    if (a === 'admin' || a === 'user') demoLogin(a);
    else if (a === 'outbox') { renderOutbox(0); $('#demoOutbox').showModal(); }
    else if (a === 'close') $('#demoOutbox').close();
    else if (a === 'reset') {
      if (!confirm('เริ่มข้อมูลทดลองใหม่?\nคำขอและการแก้ไขในโหมดทดลองจะหายทั้งหมด (ระบบจริงไม่เกี่ยว)')) return;
      window.DEMO.reset(); setToken('');
      try { localStorage.removeItem(STORE_KEY); } catch { /* ignore */ }
      location.reload();
    } else if (a === 'exit') location.href = location.pathname + '?demo=0';
    return;
  }
  const li = e.target.closest('[data-mail]');
  if (li) renderOutbox(Number(li.dataset.mail));
});

// ------------------------------------------------------------ budget options (server: getBudgetOptions)

async function refreshBudget() {
  try { budgetOpts = await gasCall('getBudgetOptions'); }
  catch (e) { budgetOpts = { year: 0, lines: [], reserved: {}, error: errMsg(e) }; }
  if (activeTab === 'order') renderOrder();
  if (activeTab === 'home') renderHome();
  if (activeTab === 'admin') renderAdmin();
  if (activeTab === 'budget') renderBudget();
  if (activeTab === 'requests') renderRequests();
}

function budgetAvail(line, m) {
  const mb = (line && line.months[m]) || { plan: 0, actual: 0 };
  const reserved = (budgetOpts.reserved || {})[line.key + '|' + m] || 0;
  return { plan: mb.plan, actual: mb.actual, reserved, available: mb.plan - mb.actual - reserved };
}

// Admin: budget upload (Admin tab button, or drop the budget file anywhere)
async function uploadBudget(rows, fileName) {
  if (!SERVER) return `✗ ${fileName}: นำเข้า Budget ได้เฉพาะในเวอร์ชัน Apps Script`;
  if (!isAdmin()) return `✗ ${fileName}: เฉพาะ Admin นำเข้าไฟล์ Budget ได้`;
  const lines = parseBudget(rows);
  if (!lines.length) return `✗ ${fileName}: ไม่พบรายการ Budget`;
  budgetOpts = await gasCall('importBudget', lines);
  if (activeTab === 'admin') renderAdmin();
  return `✓ ${fileName} → Budget ${budgetOpts.lines.length} รายการ (${lines.length} แถวรายเดือน)` +
    (lines.skipped ? `\n⚠ ข้าม ${lines.skipped} แถวที่ไม่มี GL Code หรือเลขเดือน — ตรวจในไฟล์ต้นทาง` : '');
}
$('#btnBudgetFile').addEventListener('click', () => $('#budgetFile').click());
$('#budgetFile').addEventListener('change', async e => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const wb = await readWorkbook(file);
    const name = wb.SheetNames.find(n => detectType(sheetRows(wb.Sheets[n]), file.name) === 'budget');
    if (!name) throw new Error('ไฟล์นี้ไม่ใช่ไฟล์ Budget (ไม่พบคอลัมน์ Media Location / GL Code)');
    $('#budgetStatus').textContent = 'กำลังนำเข้า…';
    toast(await uploadBudget(sheetRows(wb.Sheets[name]), file.name));
  } catch (err) {
    toast('นำเข้าไม่สำเร็จ: ' + errMsg(err));
    renderAdmin();
  }
});

// ------------------------------------------------------------ requests (server: src/Requests.js, Approval.js, Quotes.js)
//   User:  คำขอของฉัน — track, call back to edit (รอผู้ขอแก้ไข) and resubmit, see quotations
//   Admin: ตรวจคำขอ — check items + quotations, fix / send back / delete, send the order list + budget email
//          (with the Excel file + quotations attached), record the approval, Purchasing, PR/PO, received

const ST = {
  WAIT_PC: 'รอ PC ยืนยัน', PC_REJECTED: 'PC ปฏิเสธ', PENDING: 'รอ Admin ตรวจ', APPROVAL_WAIT: 'รอผู้บริหารอนุมัติ',
  OVER_WAIT: 'รอหัวหน้าอนุมัติเกินงบ', APPROVED: 'อนุมัติ', REJECTED: 'ไม่อนุมัติ', CANCELLED: 'ยกเลิก',
  PO: 'ออก PR/PO แล้ว', RECEIVED: 'รับของแล้ว', EDIT: 'รอผู้ขอแก้ไข'
};
const STATUS_CLASS = {
  [ST.WAIT_PC]: 'st-pending', [ST.PENDING]: 'st-pending', [ST.APPROVAL_WAIT]: 'st-pending', [ST.OVER_WAIT]: 'st-pending',
  [ST.EDIT]: 'st-edit', [ST.APPROVED]: 'st-approved', [ST.PO]: 'st-approved', [ST.RECEIVED]: 'st-approved',
  [ST.REJECTED]: 'st-rejected', [ST.PC_REJECTED]: 'st-cancelled', [ST.CANCELLED]: 'st-cancelled'
};
const FLOW = [ST.PENDING, ST.APPROVAL_WAIT, ST.APPROVED, ST.PO, ST.RECEIVED];
const FLOW_LABEL = { [ST.PENDING]: 'Admin ตรวจ', [ST.APPROVAL_WAIT]: 'ส่งเมล / รออนุมัติ', [ST.APPROVED]: 'อนุมัติ', [ST.PO]: 'PR/PO', [ST.RECEIVED]: 'รับของ' };
let focusReq = null;       // request id to scroll to (deep link / just sent)
let editingReq = null;     // { id, mode: 'pc' | 'admin' } whose inline form is open
let composing = null;      // { id, kind, draft } — email being prepared
let masterPc = [];
let reqFilterSet = false;  // first visit picks a sensible filter per role

const OUTDATED_MSG = 'ระบบหลังบ้าน (Apps Script) ยังเป็นโค้ดเวอร์ชันเก่า — ปุ่มลบ / แก้ไข / ส่งกลับ / แนบใบเสนอราคา จะยังใช้ไม่ได้ ' +
  'จนกว่าจะวาง build/apps-script/Code.gs เวอร์ชันใหม่ แล้ว Deploy → จัดการการทำให้ใช้งานได้ → แก้ไข → เวอร์ชันใหม่';
const selectedReqs = new Set();  // admin: requests ticked for bulk delete
let visibleReqs = [];            // ids currently listed (bulk actions only touch these)

const mine = r => r.assigned.includes(me.email);
const isRequester = r => String(r.requester_email || '').toLowerCase() === me.email;
const isManagerOf = r => String(r.manager_email || '').toLowerCase() === me.email;
const isApproverOf = r => (r.approvers || []).includes(me.email);

/** Requests that are waiting on me right now. */
function needsMe(r) {
  if (r.status === ST.WAIT_PC) return mine(r) || (isAdmin() && !r.assigned.length);
  if (r.status === ST.EDIT) return isRequester(r);
  if (r.status === ST.OVER_WAIT) return isManagerOf(r) || isAdmin();
  if (r.status === ST.APPROVAL_WAIT) return isApproverOf(r) || isAdmin();
  if (isAdmin()) return [ST.PENDING, ST.APPROVED, ST.PO].includes(r.status);
  return false;
}

async function loadRequests(badgeOnly = false) {
  if (!isMember()) return;
  if (!badgeOnly) $('#reqInfo').textContent = 'กำลังโหลด…';
  try { requests = await gasCall('listMyRequests'); }
  catch (e) { $('#reqInfo').textContent = 'โหลดไม่ได้: ' + errMsg(e); return; }
  updateReqBadge();
  updateDlvBadge();
  if (!badgeOnly || activeTab === 'requests') renderRequests();
  if (activeTab === 'delivery') renderDelivery();
}

function updateReqBadge() {
  const n = requests.filter(needsMe).length;
  document.querySelectorAll('.reqBadge').forEach(b => { b.hidden = !n; b.textContent = n; });
  if (activeTab === 'home') renderHome();
}

// ------------------------------------------------------------ home: different for Admin and User

function renderHome() {
  const admin = isAdmin();
  const hasData = merged.rows.length > 0;
  const { n: draftN, v: draftV } = draftTotals();
  const todo = isMember() ? requests.filter(needsMe) : [];
  const hour = new Date().getHours();
  const hello = hour < 12 ? 'สวัสดีตอนเช้า' : hour < 17 ? 'สวัสดีตอนบ่าย' : 'สวัสดีตอนเย็น';
  $('#homeTitle').textContent = me && me.name ? `${hello}, ${me.name}` : 'หน้าแรก';
  $('#homeSub').textContent = new Date().toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }) +
    (me && me.role ? ` · ${admin ? 'ผู้ดูแลระบบ (Admin)' : 'ผู้ใช้งาน (User)'}` : '');

  const alerts = [];
  if (serverOutdated) alerts.push([OUTDATED_MSG, '', '']);
  if (admin && !hasData) alerts.push(['ยังไม่มีข้อมูล Store — นำเข้าไฟล์ 4 ไฟล์ก่อน (รายการแนะนำ / ยอดคงเหลือ)', 'data', 'ไปนำเข้าไฟล์']);
  if (isMember() && budgetOpts && !budgetOpts.error && !budgetOpts.lines.length) alerts.push(admin
    ? ['ยังไม่มีข้อมูล Budget — User ยังส่งคำขอไม่ได้', 'admin', 'นำเข้า Budget'] : ['ยังไม่มีข้อมูล Budget — รอ Admin นำเข้า', '', '']);
  if (!admin && state.draft.editId) alerts.push([`กำลังแก้ไขคำขอ ${state.draft.editId} — ยังไม่ได้ส่งใหม่`, 'order', 'ทำต่อ']);
  $('#homeAlerts').innerHTML = alerts.map(([msg, tab, label]) =>
    `<div class="alert-row"><span>${esc(msg)}</span>${tab ? `<button class="btn sm" data-go="${tab}">${esc(label)}</button>` : ''}</div>`).join('');

  const count = (sts, fn) => requests.filter(r => sts.includes(r.status) && (!fn || fn(r)));
  const kpi = (label, list, cls, filter) => `<button class="kpi ${cls || ''}" data-kpi="${esc(filter || '')}">
    <span class="kpi-k">${esc(label)}</span><span class="kpi-v">${list.length}</span><span class="kpi-s">${money(list.reduce((a, r) => a + (Number(r.total) || 0), 0))}</span></button>`;
  const myReq = r => isRequester(r);
  $('#homeKpis').innerHTML = admin
    ? kpi('รอ Admin ตรวจ', count([ST.PENDING]), 'warn', ST.PENDING) + kpi('รอผู้บริหารอนุมัติ', count([ST.APPROVAL_WAIT]), '', ST.APPROVAL_WAIT) +
      kpi('อนุมัติแล้ว รอ PR/PO', count([ST.APPROVED]), 'good', ST.APPROVED) + kpi('ส่งกลับให้แก้ไข', count([ST.EDIT]), '', ST.EDIT)
    : kpi('ต้องแก้ไข', count([ST.EDIT], myReq), 'warn', ST.EDIT) + kpi('รอ Admin ตรวจ', count([ST.PENDING], myReq), '', ST.PENDING) +
      kpi('รออนุมัติ', count([ST.APPROVAL_WAIT], myReq), '', ST.APPROVAL_WAIT) + kpi('อนุมัติ / สั่งแล้ว', count([ST.APPROVED, ST.PO, ST.RECEIVED], myReq), 'good', ST.APPROVED);

  const card = (tab, title, desc, extra, primary) => `<button class="action-card ${primary ? 'primary' : ''}" data-go="${tab}">
    <b>${title}</b><span>${desc}</span>${extra ? `<span class="ac-extra">${extra}</span>` : ''}</button>`;
  $('#homeActions').innerHTML = admin
    ? card('requests', 'ตรวจคำขอสั่งซื้อ', 'ตรวจรายการ + ใบเสนอราคา แก้ไข/ส่งกลับ แล้วส่งเมลรายการ + Budget', todo.length ? `รอคุณ ${todo.length} รายการ` : '', true) +
      card('budget', 'งบประมาณ', 'ทุกรายการ Budget รายเดือน ใช้แล้ว / คงเหลือ') +
      card('data', 'ข้อมูล Store', hasData ? `${fmt(merged.rows.length)} รายการ · ${merged.months.length} เดือน` : 'ยังไม่มีข้อมูล') +
      card('admin', 'สมาชิก & ตั้งค่า', 'สมาชิก · Budget · อีเมลส่งรายการ · Master PC')
    : card('order', 'สั่งซื้อตาม Budget', 'เลือก Budget + เดือน แล้วใส่รายการที่จะสั่ง', draftN ? `ในคำขอ ${draftN} รายการ · ${money(draftV)}` : '', true) +
      card('requests', 'คำขอของฉัน', 'ติดตามสถานะ · เรียกกลับมาแก้ไขแล้วส่งใหม่', todo.length ? `ต้องแก้ไข ${todo.length} รายการ` : '') +
      card('budget', 'ตรวจสอบงบประมาณ', 'ดูงบแต่ละรายการ รายเดือน ใช้แล้ว / เหลือ') +
      card('compare', 'รายการแนะนำจาก Store', hasData ? 'ของที่ควรสั่งตามยอดใช้จริง' : 'รอ Admin นำเข้าข้อมูล Store');

  const step = (n, title, desc) => `<li><span class="step-no">${n}</span><div class="step-body"><b>${title}</b><div class="hint">${desc}</div></div></li>`;
  $('#homeFlowTitle').textContent = admin ? 'ขั้นตอนของ Admin' : 'ขั้นตอนการสั่งซื้อ';
  $('#homeSteps').innerHTML = admin
    ? step(1, 'ตรวจรายการสั่งซื้อ + ใบเสนอราคา', 'เปิดคำขอที่ "รอ Admin ตรวจ" ดูรายการ งบ และไฟล์แนบ') +
      step(2, 'แก้ไข / ส่งกลับ / ลบ ถ้าข้อมูลผิด', 'แก้จำนวน ราคา Budget เดือน หรือส่งกลับให้ผู้ขอแก้') +
      step(3, 'ส่งเมลรายการสั่งซื้อ + Budget', `ถึง ${esc((appSettings && appSettings.approval_to) || 'thanakorn@planbmedia.co.th')} แนบไฟล์ Excel + ใบเสนอราคา`) +
      step(4, 'บันทึกผลอนุมัติ → แจ้งจัดซื้อ → PR/PO → รับของ', 'เมื่อได้รับ "Approved"')
    : step(1, 'ตรวจสอบงบประมาณ', 'ดูว่า Budget ไหน เดือนไหนยังเหลือ') +
      step(2, 'เลือก Budget + เดือน แล้วใส่รายการ', 'ค้นหาสินค้าใน Store หรือเพิ่มรายการใหม่ · แนบใบเสนอราคาได้') +
      step(3, 'กดส่งให้ Admin', 'งบไม่พอ → ติ๊ก "ขอ Over Budget" พร้อมเหตุผล') +
      step(4, 'ติดตาม / แก้ไขแล้วส่งใหม่', 'เรียกคำขอกลับมาแก้ได้ก่อน Admin ส่งเมล');

  $('#homeTodoTitle').firstChild.textContent = admin ? 'งานที่รอคุณ ' : 'คำขอล่าสุดของฉัน ';
  $('#homeTodoCount').hidden = !todo.length;
  $('#homeTodoCount').textContent = todo.length;
  const list = admin ? todo : requests.filter(isRequester).slice(0, 6);
  $('#homeTodo').innerHTML = !list.length ? `<div class="empty-sm">${admin ? 'ไม่มีงานค้าง 🎉' : 'ยังไม่มีคำขอ — เริ่มที่ "สั่งซื้อตาม Budget"'}</div>`
    : `<ul class="todo-list">${list.slice(0, 6).map(r => `<li data-open-req="${esc(r.request_id)}">
        <span class="tag ${STATUS_CLASS[r.status] || ''}">${esc(r.status)}</span>
        <span class="todo-main"><b>${esc(r.request_id)}</b> ${esc(admin ? (r.requester_name || '') : (r.budget_months || ''))}${r.files && r.files.length ? ' 📎' : ''}</span>
        <span class="num">${money(r.total)}</span></li>`).join('')}</ul>
      ${list.length > 6 ? `<button class="btn sm ghost" data-go="requests">ดูทั้งหมด ${list.length} รายการ</button>` : ''}`;

  $('#summary').hidden = !admin || !hasData;
}

document.addEventListener('click', e => {
  const go = e.target.closest('[data-go]');
  if (go) { if (go.dataset.go === 'requests') { $('#reqFilter').value = 'todo'; reqFilterSet = true; } goTab(go.dataset.go); return; }
  const k = e.target.closest('[data-kpi]');
  if (k) { $('#reqFilter').value = k.dataset.kpi; reqFilterSet = true; goTab('requests'); return; }
  const open = e.target.closest('[data-open-req]');
  if (open) { $('#reqFilter').value = ''; reqFilterSet = true; focusReq = open.dataset.openReq; goTab('requests'); }
});

// ------------------------------------------------------------ request list + cards

/**
 * A request grouped by budget line x month, with what is still free for that
 * line/month not counting the request itself: [{ key, month, label, amount, plan, avail, overNow, over, reason }].
 */
function reqBudgetLines(r) {
  const g = new Map();
  for (const i of r.items) {
    if (!(i.qty > 0)) continue;
    const key = i.budget_key || r.budget_key, month = i.budget_month || Number(r.budget_month);
    if (!key || !month) continue;
    const k = key + '|' + month;
    if (!g.has(k)) g.set(k, { key, month, label: i.budget_label || r.budget_label, amount: 0, over: !!i.over_budget, reason: i.over_reason || '' });
    g.get(k).amount += Number(i.amount) || 0;
  }
  const holds = [ST.PENDING, ST.APPROVAL_WAIT, ST.OVER_WAIT, ST.APPROVED, ST.PO, ST.RECEIVED].includes(r.status);
  return [...g.values()].map(x => {
    const line = budgetLine(x.key);
    if (line) {
      const b = budgetAvail(line, x.month);
      x.label = line.label; x.plan = b.plan; x.avail = b.available + (holds ? x.amount : 0);
      x.overNow = x.amount > x.avail + 0.004;
    }
    return x;
  });
}

// หน้าคำขอแบ่งเป็น "ขั้นตอน" ไม่ใช่กองเดียว — แต่ละขั้นบอกว่าตอนนี้รอใคร และฝั่งเราต้องทำอะไร
const STAGES = [
  { id: 'back', no: 1, title: 'อยู่ที่ผู้ขอ', st: [ST.WAIT_PC, ST.EDIT],
    admin: 'ส่งกลับให้ผู้ขอ/PC แก้แล้ว — รอเขาส่งกลับมา', user: 'แก้ไขแล้วกด "แก้ไขแล้วส่งใหม่" หรือยืนยันจำนวนให้ Admin' },
  { id: 'check', no: 2, title: 'Admin ตรวจ', st: [ST.PENDING, ST.OVER_WAIT],
    admin: 'ตรวจรายการ + Budget + ใบเสนอราคา → แก้ไข/ส่งกลับ หรือกด "ส่งเมลรายการสั่งซื้อ + Budget"', user: 'ส่งให้ Admin แล้ว — เรียกกลับมาแก้ไขได้จนกว่า Admin จะส่งเมล' },
  { id: 'approve', no: 3, title: 'รอผู้บริหารอนุมัติ', st: [ST.APPROVAL_WAIT],
    admin: 'ส่งเมลแล้ว — เมื่อได้ "Approved" กด "บันทึก: ได้รับอนุมัติแล้ว"', user: 'รอผลอนุมัติทางอีเมล' },
  { id: 'po', no: 4, title: 'อนุมัติแล้ว — ออก PR/PO', st: [ST.APPROVED],
    admin: 'ส่งแจ้งฝ่ายจัดซื้อ → ใส่เลข PR/PO รายชิ้น + กำหนดส่ง (ตัดงบตอนได้เลข PO)', user: 'อนุมัติแล้ว รอฝ่ายจัดซื้อออก PR/PO' },
  { id: 'wait', no: 5, title: 'รอของมาส่ง', st: [ST.PO],
    admin: 'ติดตามของ → ติ๊กรายการที่มาถึงแล้วใส่วันที่รับของ', user: 'ออก PR/PO แล้ว — ดูกำหนดส่งได้ที่แท็บติดตามการส่งของ' },
  { id: 'done', no: 6, title: 'รับของแล้ว', st: [ST.RECEIVED], admin: 'จบขั้นตอน', user: 'ของเข้า Store แล้ว' },
  { id: 'closed', no: 7, title: 'ไม่อนุมัติ / ยกเลิก', st: [ST.REJECTED, ST.CANCELLED, ST.PC_REJECTED],
    admin: 'ปิดแล้ว — ผู้ขอแก้แล้วส่งใหม่ได้', user: 'ปิดแล้ว — แก้ไขแล้วส่งใหม่ได้' }
];
const stageOf = r => (STAGES.find(s => s.st.includes(r.status)) || STAGES[1]).id;
let stageFilter = '';

function renderRequests() {
  const admin = isAdmin();
  $('#reqTitle').textContent = admin ? 'ตรวจคำขอสั่งซื้อ' : 'คำขอของฉัน';
  if (!reqFilterSet) { $('#reqFilter').value = admin ? 'todo' : ''; reqFilterSet = true; }
  const f = $('#reqFilter').value;
  const words = $('#reqSearch').value.trim().toLowerCase().split(/\s+/).filter(Boolean);
  const text = r => [r.request_id, r.requester_name, r.requester_email, r.note, r.budget_label, ...r.items.map(i => i.name + ' ' + i.sku)].join(' ').toLowerCase();
  const list = requests.filter(r => (!f || (f === 'todo' ? needsMe(r) : r.status === f)) && words.every(w => text(r).includes(w))
    && (!stageFilter || stageOf(r) === stageFilter));

  // แถบขั้นตอน: นับจากคำขอที่ผ่านตัวกรองค้นหา/สถานะ แต่ยังไม่กรองขั้นตอน (กดเพื่อดูเฉพาะขั้นนั้น)
  const pool = requests.filter(r => (!f || (f === 'todo' ? needsMe(r) : r.status === f)) && words.every(w => text(r).includes(w)));
  const byStage = {};
  for (const r of pool) {
    const k = stageOf(r);
    byStage[k] = byStage[k] || { n: 0, v: 0, todo: 0 };
    byStage[k].n++; byStage[k].v += Number(r.total) || 0;
    if (needsMe(r)) byStage[k].todo++;
  }
  $('#reqDash').innerHTML = `<div class="stage-bar">${STAGES.map(s => {
    const b = byStage[s.id] || { n: 0, v: 0, todo: 0 };
    return `<button class="stage-chip ${stageFilter === s.id ? 'sel' : ''} ${b.todo ? 'todo' : ''} ${b.n ? '' : 'empty'}" data-stage="${s.id}"
      title="${esc(admin ? s.admin : s.user)}"><span class="sn">${s.no}</span>
      <span class="st-title">${esc(s.title)}</span><b>${b.n}</b>${b.todo ? `<span class="dot" title="รอคุณ ${b.todo}">${b.todo}</span>` : ''}
      <span class="st-sum">${b.n ? money(b.v) : '—'}</span></button>`;
  }).join('')}</div>${stageFilter ? `<button class="link-btn" data-stage="">← ดูทุกขั้นตอน</button>` : ''}`;

  $('#reqInfo').textContent = `${list.length} คำขอ${admin ? '' : ' ของฉัน / ที่เกี่ยวกับฉัน'}` +
    (stageFilter ? ` · เฉพาะขั้นที่ ${(STAGES.find(s => s.id === stageFilter) || {}).no}` : '');
  visibleReqs = list.map(r => r.request_id);
  updateBulkBar();
  const banner = serverOutdated ? `<div class="alert-row"><span>${esc(OUTDATED_MSG)}</span></div>` : '';
  if (!list.length) {
    $('#reqList').innerHTML = banner + `<div class="empty">${f === 'todo' ? 'ไม่มีงานที่รอคุณ 🎉' : 'ไม่มีคำขอ'}${!admin ? '<br><button class="btn sm" data-go="order">สั่งซื้อตาม Budget</button>' : ''}</div>`;
    return;
  }
  // จัดการ์ดเข้าแต่ละขั้นตอน ขั้นที่ไม่มีคำขอจะไม่แสดง
  $('#reqList').innerHTML = banner + STAGES.map(s => {
    const rows = list.filter(r => stageOf(r) === s.id);
    if (!rows.length) return '';
    const v = rows.reduce((a, r) => a + (Number(r.total) || 0), 0);
    const todo = rows.filter(needsMe).length;
    return `<details class="stage" id="stage-${s.id}" open>
      <summary class="stage-h"><span class="sn">${s.no}</span>
        <span class="stage-t"><b>${esc(s.title)}</b><span class="hint">${esc(admin ? s.admin : s.user)}</span></span>
        <span class="stage-sum">${rows.length} คำขอ · ${money(v)}${todo ? ` · <span class="over">รอคุณ ${todo}</span>` : ''}</span></summary>
      <div class="stage-body">${rows.map(renderReqCard).join('')}</div>
    </details>`;
  }).join('');
  if (focusReq) {
    const el = document.querySelector(`.req-card[data-id="${CSS.escape(focusReq)}"]`);
    if (el) {
      el.classList.add('focus'); el.scrollIntoView({ block: 'center' });
      const r = requests.find(x => x.request_id === focusReq);
      if (r && r.status === ST.WAIT_PC && (mine(r) || admin) && !editingReq) { editingReq = { id: focusReq, mode: 'pc' }; focusReq = null; renderRequests(); return; }
    }
    focusReq = null;
  }
}

$('#reqDash').addEventListener('click', e => {
  const b = e.target.closest('[data-stage]');
  if (!b) return;
  stageFilter = b.dataset.stage === stageFilter ? '' : b.dataset.stage;
  renderRequests();
});

/** Admin: tick several requests and delete them in one go (e.g. rounds nobody answered). */
function updateBulkBar() {
  const admin = isAdmin();
  const n = [...selectedReqs].filter(id => visibleReqs.includes(id)).length;
  $('#btnSelectAll').hidden = !admin || !visibleReqs.length;
  $('#btnSelectAll').textContent = n && n === visibleReqs.length ? 'ไม่เลือกทั้งหมด' : `เลือกทั้งหมดที่แสดง (${visibleReqs.length})`;
  $('#btnBulkDelete').hidden = !admin || !n;
  $('#btnBulkDelete').textContent = `ลบที่เลือก (${n})`;
}

$('#btnSelectAll').addEventListener('click', () => {
  const all = visibleReqs.every(id => selectedReqs.has(id));
  visibleReqs.forEach(id => (all ? selectedReqs.delete(id) : selectedReqs.add(id)));
  renderRequests();
});

$('#btnBulkDelete').addEventListener('click', async () => {
  const ids = visibleReqs.filter(id => selectedReqs.has(id));
  if (!ids.length) return;
  const rows = ids.map(id => requests.find(r => r.request_id === id)).filter(Boolean);
  const total = rows.reduce((a, r) => a + (Number(r.total) || 0), 0);
  if (!confirm(`ลบ ${ids.length} คำขอ (รวม ${money(total)}) ทั้งหมด?\n` + rows.slice(0, 8).map(r => `• ${r.request_id} ${r.status} ${money(r.total)}`).join('\n') +
    (rows.length > 8 ? `\n… และอีก ${rows.length - 8} รายการ` : '') + '\n\nรายการสินค้าจะถูกลบด้วย ย้อนกลับไม่ได้ในระบบ')) return;
  const btn = $('#btnBulkDelete');
  btn.disabled = true;
  let done = 0, failed = [];
  for (const id of ids) {
    try { requests = await gasCall('deleteRequest', id); selectedReqs.delete(id); done++; }
    catch (e) { failed.push(id + ': ' + errMsg(e)); }
    btn.textContent = `กำลังลบ… ${done}/${ids.length}`;
  }
  btn.disabled = false;
  updateReqBadge(); renderRequests(); refreshBudget();
  toast(`ลบแล้ว ${done} คำขอ` + (failed.length ? ` · ลบไม่ได้ ${failed.length}: ${failed[0]}` : ''));
});

function fileChips(r) {
  if (!r.files || !r.files.length) return '';
  return `<div class="req-files"><span class="hint">ใบเสนอราคา:</span> ${r.files.map(f =>
    `<button class="file-chip" data-file="${esc(f.id)}" title="เปิดดู">📎 ${esc(f.name)} <span class="hint">${kb(f.size)}</span></button>`).join('')}</div>`;
}

function renderReqCard(r) {
  const admin = isAdmin();
  const step = FLOW.indexOf(r.status === ST.OVER_WAIT ? ST.PENDING : r.status);
  const blines = reqBudgetLines(r);
  const closed = [ST.REJECTED, ST.CANCELLED, ST.PC_REJECTED].includes(r.status);
  const overAny = !closed && (r.over_budget || blines.some(x => x.over || x.overNow));
  const who = r.source === 'round' ? `ส่งให้ ${esc(r.pc_code)} ${esc(r.pc_name)} · ผู้รับผิดชอบ ${esc(r.assigned.join(', ') || '— ยังไม่ได้ตั้งใน Master PC')}`
    : `ขอโดย <b>${esc(r.requester_name || r.requester_email)}</b>${r.pc_code ? ' · PC ' + esc(r.pc_code) : ''}`;
  const B = (act, label, cls = 'ghost') => `<button class="btn sm ${cls}" data-act="${act}">${label}</button>`;
  const acts = [];
  // requester
  if (isRequester(r)) {
    if (r.status === ST.EDIT || r.status === ST.REJECTED) acts.push(B('recall', 'แก้ไขแล้วส่งใหม่', ''));
    if (r.status === ST.PENDING) acts.push(B('recall', 'เรียกกลับมาแก้ไข'));
  }
  if (r.status === ST.WAIT_PC && (mine(r) || admin)) acts.push(B('edit', 'ยืนยัน / แก้จำนวน', ''), B('pcreject', 'ไม่สั่งรอบนี้'));
  // admin
  if (admin) {
    if (r.status === ST.PENDING) acts.push(B('compose', 'ตรวจแล้ว → ส่งเมลรายการสั่งซื้อ + Budget', ''));
    if (r.status === ST.APPROVAL_WAIT) acts.push(B('approvedok', isApproverOf(r) ? 'อนุมัติ' : 'บันทึก: ได้รับอนุมัติแล้ว', ''), B('approvedno', isApproverOf(r) ? 'ไม่อนุมัติ' : 'บันทึก: ไม่อนุมัติ'), B('compose', 'ส่งเมลอีกครั้ง'));
    if ([ST.PENDING, ST.APPROVAL_WAIT, ST.EDIT].includes(r.status)) acts.push(B('aedit', 'แก้ไขข้อมูล'));
    if ([ST.PENDING, ST.APPROVAL_WAIT].includes(r.status)) acts.push(B('return', 'ส่งกลับให้ผู้ขอแก้ไข'), B('reject', 'ไม่อนุมัติ'));
    if ([ST.APPROVED, ST.PO].includes(r.status)) acts.push(B('purchasing', r.purchasing_sent_at ? 'ส่งแจ้งฝ่ายจัดซื้ออีกครั้ง' : 'ส่งแจ้งฝ่ายจัดซื้อ', r.purchasing_sent_at ? 'ghost' : ''));
    if (r.status === ST.APPROVED) acts.push(B('po', 'ออก PR/PO', ''));
    if (r.status === ST.PO) acts.push(B('received', 'รับของเข้าแล้ว', ''));
  } else if (r.status === ST.APPROVAL_WAIT && isApproverOf(r)) {
    acts.push(B('approvedok', 'อนุมัติ', ''), B('approvedno', 'ไม่อนุมัติ'));
  }
  if (r.status === ST.OVER_WAIT && (isManagerOf(r) || admin)) acts.push(B('mgrok', 'หัวหน้าอนุมัติเกินงบ', ''), B('mgrno', 'หัวหน้าไม่อนุมัติ'));
  if ([ST.WAIT_PC, ST.PENDING, ST.OVER_WAIT, ST.EDIT].includes(r.status) && (isRequester(r) || admin)) acts.push(B('cancel', 'ยกเลิกคำขอ'));
  if (admin && ![ST.PO, ST.RECEIVED].includes(r.status)) acts.push(B('delete', 'ลบ', 'ghost danger'));

  const budgetHtml = blines.length ? `<ul class="req-budget">${blines.map(x => `<li class="${x.over || x.overNow ? 'is-over' : ''}">
      <span>${esc(x.label)} · <b>${esc(THAI_MONTH_FULL[x.month] || '')}</b></span>
      <span class="num">${money(x.amount)}${x.avail != null ? ` <span class="hint">/ เหลือ ${money(x.avail)}</span>` : ''}</span>
      ${x.over ? `<div class="over">ขอ Over Budget${x.reason ? ': ' + esc(x.reason) : ''}</div>` : x.overNow && !closed ? '<div class="over">งบไม่พอแล้ว</div>' : ''}
    </li>`).join('')}</ul>` : '';
  const editing = editingReq && editingReq.id === r.request_id;
  const isComposing = composing && composing.id === r.request_id;

  return `<div class="req-card ${needsMe(r) ? 'todo' : ''}" data-id="${esc(r.request_id)}">
    <div class="req-head">
      ${admin ? `<input type="checkbox" class="req-sel" data-sel ${selectedReqs.has(r.request_id) ? 'checked' : ''} aria-label="เลือก ${esc(r.request_id)}">` : ''}
      <span class="id">${esc(r.request_id)}</span>
      <span class="tag ${STATUS_CLASS[r.status] || ''}">${esc(r.status)}</span>
      ${overAny ? '<span class="tag st-rejected">Over Budget</span>' : ''}
      ${r.files && r.files.length ? `<span class="tag none">📎 ${r.files.length}</span>` : ''}
      <span class="amt">${money(r.total)}</span>
    </div>
    ${step >= 0 ? `<ol class="flow">${FLOW.map((s, i) => `<li class="${i < step ? 'done' : i === step ? 'now' : ''}">${esc(FLOW_LABEL[s])}</li>`).join('')}</ol>` : ''}
    ${r.status === ST.EDIT ? `<div class="edit-note"><b>${isRequester(r) ? 'ต้องแก้ไข' : 'รอผู้ขอแก้ไข'}:</b> ${esc(r.admin_note || '')}</div>` : ''}
    <div class="req-meta">${who} · สร้าง ${esc(r.created_at)} · ${r.item_count} รายการ</div>
    ${budgetHtml}
    ${r.note ? `<div class="req-meta">หมายเหตุผู้ขอ: ${esc(r.note)}</div>` : ''}
    ${fileChips(r)}
    ${r.confirmed_by && r.source === 'round' ? `<div class="req-meta">ยืนยันโดย ${esc(r.confirmed_by)} ${esc(r.confirmed_at)}</div>` : ''}
    ${r.status === ST.OVER_WAIT || r.manager_decision ? `<div class="req-meta">หัวหน้า PC: ${esc(r.manager_email)}${r.manager_note ? ' — ' + esc(r.manager_note) : ''}</div>` : ''}
    ${r.approval_sent_at ? `<div class="req-meta">ส่งเมลรายการสั่งซื้อถึง ${esc(r.approval_to)}${r.approval_cc ? ' (สำเนา ' + esc(r.approval_cc) + ')' : ''} · ${esc(r.approval_sent_at)}</div>` : ''}
    ${r.approved_at ? `<div class="req-meta">ผลอนุมัติ: ${esc(r.approved_by)} ${esc(r.approved_at)}${r.approval_note ? ' — ' + esc(r.approval_note) : ''}</div>` : ''}
    ${r.purchasing_sent_at ? `<div class="req-meta">แจ้งฝ่ายจัดซื้อแล้ว ${esc(r.purchasing_sent_at)}</div>` : ''}
    ${r.status !== ST.EDIT && r.decided_by && r.admin_note ? `<div class="req-meta">Admin ${esc(r.decided_by)}${r.decided_at ? ' ' + esc(r.decided_at) : ''} — ${esc(r.admin_note)}</div>` : ''}
    ${r.po_no ? `<div class="req-meta">PR/PO: <b>${esc(r.po_no)}</b> ${esc(r.po_at)}${r.received_at ? ' · รับของ ' + esc(r.received_at) : ''}</div>` : ''}
    ${editing && editingReq.mode === 'pc' ? renderConfirmForm(r) : editing && editingReq.mode === 'admin' ? renderAdminEditForm(r)
      : `<details ${admin && r.status === ST.PENDING ? 'open' : ''}><summary>ดูรายการสินค้า (${r.items.length})</summary>${reqItemsTable(r, false)}</details>`}
    ${deliveryPanel(r)}
    ${isComposing ? renderComposer(r) : ''}
    ${acts.length && !editing && !isComposing ? `<div class="req-actions">${acts.join('')}</div>` : ''}
  </div>`;
}

function reqItemsTable(r, editable) {
  const hasBudget = r.items.some(i => i.budget_key);
  return `<div class="table-wrap" style="max-height:none;margin-top:6px"><table class="grid"><thead><tr>
    <th>PC</th><th>รหัส</th><th>สินค้า</th>${r.source === 'round' ? '<th class="n">แนะนำ</th>' : ''}<th class="n">จำนวน</th><th>หน่วย</th><th class="n">ราคา/หน่วย</th><th class="n">รวม</th>
    ${hasBudget && !editable ? '<th>Budget / เดือน</th><th class="n">Store คงเหลือ</th>' : ''}<th>ขอบเขตการใช้งาน / หมายเหตุ</th>
  </tr></thead><tbody>${r.items.map(i => `<tr data-sku="${esc(i.sku)}" class="${!editable && i.qty === 0 ? 'dim' : ''}">
    <td class="pc">${esc(i.pc_code)}</td><td class="sku">${esc(String(i.sku).startsWith('NEW-') ? 'ใหม่' : i.sku)}</td><td class="name">${esc(i.name)}</td>
    ${r.source === 'round' ? `<td class="n dim">${fmt(i.suggested_qty)}</td>` : ''}
    <td class="n">${editable ? `<input class="qty" type="number" min="0" step="1" data-eq value="${i.qty}" data-cost="${i.unit_cost}">` : fmt(i.qty)}</td>
    <td>${esc(i.unit)}</td><td class="n">${fmt(i.unit_cost, 2)}</td><td class="n" data-amt>${fmt(i.amount, 2)}</td>
    ${hasBudget && !editable ? `<td class="small">${esc(i.budget_label)} · ${esc(THAI_MONTH_FULL[i.budget_month] || '')}${i.over_budget ? ' <span class="over">Over</span>' : ''}</td><td class="n">${i.store_balance === '' ? '-' : fmt(i.store_balance)}</td>` : ''}
    <td>${editable ? `<input class="note" data-en value="${esc(i.note)}">` : esc(i.note)}</td>
  </tr>`).join('')}</tbody></table></div>`;
}

// ---- Admin: correct the data (qty, price, budget line, month, note, remove lines)

function renderAdminEditForm(r) {
  const lineOpts = key => budgetLines().map(l => `<option value="${esc(l.key)}" ${l.key === key ? 'selected' : ''}>${esc(l.label)}</option>`).join('');
  const monthOpts = m => MONTHS_1_12.map(x => `<option value="${x}" ${x === m ? 'selected' : ''}>${MONTH_SHORT[x - 1]}</option>`).join('');
  return `<div class="confirm-form admin-edit" data-aform="${esc(r.request_id)}">
    <p class="hint">แก้ไขได้ทุกช่อง · ติ๊ก "ลบ" เพื่อเอารายการออก · ระบบตรวจงบใหม่ (ไม่นับยอดเดิมของคำขอนี้) · ผู้ขอจะได้รับอีเมลแจ้ง</p>
    <div class="table-wrap" style="max-height:none"><table class="grid"><thead><tr>
      <th>สินค้า</th><th class="n">จำนวน</th><th>หน่วย</th><th class="n">ราคา/หน่วย</th><th>Budget</th><th>เดือน</th><th>ขอบเขตการใช้งาน</th><th>ลบ</th></tr></thead>
    <tbody>${r.items.map((i, n) => `<tr data-row="${n}" data-sku="${esc(i.sku)}" data-pc="${esc(i.pc_code)}" data-pcname="${esc(i.pc_name)}" data-bal="${esc(i.store_balance)}">
      <td><input class="note" data-a="name" value="${esc(i.name)}"><div class="pc">${esc(String(i.sku).startsWith('NEW-') ? 'รายการใหม่' : i.sku)}</div></td>
      <td class="n"><input class="qty" type="number" min="0" step="1" data-a="qty" value="${i.qty}"></td>
      <td><input class="unit-in" data-a="unit" value="${esc(i.unit)}"></td>
      <td class="n"><input class="qty" type="number" min="0" step="0.01" data-a="cost" value="${i.unit_cost}"></td>
      <td><select class="bsel" data-a="key">${lineOpts(i.budget_key || r.budget_key)}</select></td>
      <td><select data-a="month">${monthOpts(Number(i.budget_month || r.budget_month))}</select></td>
      <td><input class="note" data-a="note" value="${esc(i.note)}"></td>
      <td><input type="checkbox" data-a="rm" aria-label="ลบรายการนี้"></td>
    </tr>`).join('')}</tbody></table></div>
    <div class="aedit-budget" data-abudget></div>
    <div class="field"><label>หมายเหตุการแก้ไข (แจ้งผู้ขอ)</label><input data-anote placeholder="เช่น แก้ Budget เป็นเดือนมิถุนายน"></div>
    <div class="req-actions"><button class="btn" data-act="asave">บันทึกการแก้ไข</button><button class="btn ghost" data-act="closeedit">ปิด</button></div>
  </div>`;
}

/** Items as currently typed in the admin edit form. */
function adminFormItems(form) {
  return [...form.querySelectorAll('tr[data-row]')].filter(tr => !tr.querySelector('[data-a="rm"]').checked).map(tr => {
    const v = a => tr.querySelector(`[data-a="${a}"]`).value;
    return { sku: tr.dataset.sku, name: v('name'), pc: tr.dataset.pc, pcName: tr.dataset.pcname, unit: v('unit'),
      qty: Math.max(0, Math.round(Number(v('qty')) || 0)), unitCost: Math.max(0, Number(v('cost')) || 0), note: v('note'),
      budgetKey: v('key'), month: Number(v('month')), storeBalance: tr.dataset.bal };
  });
}

function updateAdminEditBudget(form) {
  const r = requests.find(x => x.request_id === form.dataset.aform);
  if (!r) return;
  const own = {}; // this request's current amounts don't count against itself
  const holds = [ST.PENDING, ST.APPROVAL_WAIT, ST.OVER_WAIT].includes(r.status);
  if (holds) r.items.forEach(i => { const k = (i.budget_key || r.budget_key) + '|' + (i.budget_month || r.budget_month); own[k] = (own[k] || 0) + (Number(i.amount) || 0); });
  const reasons = {};
  r.items.forEach(i => { if (i.over_reason) reasons[(i.budget_key || r.budget_key) + '|' + (i.budget_month || r.budget_month)] = i.over_reason; });
  form.querySelectorAll('[data-areason]').forEach(inp => { reasons[inp.dataset.areason] = inp.value; });
  const g = new Map();
  adminFormItems(form).forEach(i => { const k = i.budgetKey + '|' + i.month; g.set(k, (g.get(k) || 0) + i.qty * i.unitCost); });
  const box = form.querySelector('[data-abudget]');
  box.innerHTML = [...g].map(([k, amt]) => {
    const { key, month: m } = splitKM(k);
    const line = budgetLine(key);
    if (!line) return '';
    const avail = budgetAvail(line, m).available + (own[k] || 0);
    const over = amt > avail + 0.004;
    return `<div class="bline ${over ? 'is-over' : ''}"><div class="row"><span>${esc(line.label)} · ${THAI_MONTH_FULL[m]}</span>
      <b class="${over ? 'over' : 'ok'}">${money(amt)} / เหลือ ${money(avail)}</b></div>
      ${over ? `<input class="over-reason" data-areason="${esc(k)}" value="${esc(reasons[k] || '')}" placeholder="เหตุผล Over Budget (จำเป็น)">` : ''}</div>`;
  }).join('');
}

// ---- email composer (order list + budget → thanakorn@ first; Purchasing after approval)

function renderComposer(r) {
  const d = composing.draft;
  if (!d) return '<div class="composer"><p class="hint">กำลังเตรียมอีเมล…</p></div>';
  const purch = composing.kind === 'purchasing';
  const files = d.files || [];
  return `<div class="composer" data-compose="${esc(r.request_id)}">
    <h4>${purch ? 'ส่งแจ้งฝ่ายจัดซื้อ (ได้รับการอนุมัติแล้ว)' : 'ส่งเมลรายการสั่งซื้อ + Budget'}</h4>
    <div class="compose-grid">
      <label>ถึง<input data-c="to" value="${esc(d.to)}" placeholder="อีเมล คั่นด้วย ,"></label>
      <label>สำเนา (CC)<input data-c="cc" value="${esc(d.cc)}" placeholder="ผู้เกี่ยวข้อง"></label>
      <label>เรื่อง<input data-c="subject" value="${esc(d.subject)}"></label>
      <label>เรียน<input data-c="greeting" value="${esc(d.greeting)}" placeholder="${purch ? 'ฝ่ายจัดซื้อ' : 'เช่น ผู้บริหาร / คุณวสุ'}"></label>
      <label class="wide">ข้อความเพิ่มเติม<textarea data-c="intro" rows="2" placeholder="เช่น ****เบื้องต้นได้ปรึกษา…เรียบร้อยค่ะ">${esc(d.intro)}</textarea></label>
      <label class="wide">ลิงก์แนบ (ไม่บังคับ)<input data-c="link" value="${esc(d.link)}" placeholder="https://…"></label>
    </div>
    <div class="attach-box"><b>ไฟล์แนบ</b>
      <label><input type="checkbox" data-attach="excel" checked> 📊 ${esc(d.excelName || 'รายการสั่งซื้อและ Budget.xlsx')} <span class="hint">(สร้างจากรายการ + Budget ตอนส่ง)</span></label>
      ${files.map(f => `<label><input type="checkbox" data-attach="${esc(f.id)}" checked> 📎 ${esc(f.name)} <span class="hint">${kb(f.size)}</span>
        <button type="button" class="link-btn" data-file="${esc(f.id)}">เปิดดู</button></label>`).join('') || '<span class="hint">ผู้ขอไม่ได้แนบใบเสนอราคา</span>'}
    </div>
    <p class="hint">ค่าตั้งต้นแก้ได้ที่ สมาชิก &amp; ตั้งค่า → ตั้งค่าอีเมล · ได้รับ "Approved" แล้วกด "บันทึก: ได้รับอนุมัติแล้ว"</p>
    <iframe class="mail-preview" title="ตัวอย่างอีเมล" sandbox srcdoc="${esc(d.html)}"></iframe>
    <div class="req-actions">
      <button class="btn" data-act="sendmail">${purch ? 'ส่งถึงฝ่ายจัดซื้อ' : 'ส่งเมล'}</button>
      <button class="btn ghost" data-act="previewmail">อัปเดตตัวอย่าง</button>
      <button class="btn ghost" data-act="closecompose">ปิด</button>
    </div>
  </div>`;
}

function composerValues(id) {
  const box = document.querySelector(`[data-compose="${CSS.escape(id)}"]`);
  const v = {};
  if (!box) return v;
  box.querySelectorAll('[data-c]').forEach(el => { v[el.dataset.c] = el.value; });
  v.attachExcel = !!(box.querySelector('[data-attach="excel"]') || {}).checked;
  v.attachFileIds = [...box.querySelectorAll('[data-attach]:not([data-attach="excel"])')].filter(x => x.checked).map(x => x.dataset.attach);
  return v;
}

async function openComposer(id, kind, overrides) {
  composing = { id, kind, draft: composing && composing.id === id && composing.kind === kind ? composing.draft : null };
  renderRequests();
  try {
    const d = await gasCall('getApprovalDraft', id, kind, overrides || null);
    if (!composing || composing.id !== id) return;
    composing.draft = overrides ? { ...d, ...overrides, html: d.html, files: d.files, excelName: d.excelName } : d;
    renderRequests();
    const box = document.querySelector(`[data-compose="${CSS.escape(id)}"]`);
    if (box && !overrides) box.scrollIntoView({ block: 'start' });
  } catch (err) {
    composing = null; renderRequests(); toast('เตรียมอีเมลไม่ได้: ' + errMsg(err));
  }
}

// ---- quotation viewer (admin checks the files; requester sees their own)

async function openQuote(fileId) {
  const dlg = $('#fileViewer');
  $('#fileViewerBody').innerHTML = '<p class="hint">กำลังเปิดไฟล์…</p>';
  $('#fileViewerTitle').textContent = 'ใบเสนอราคา';
  if (!dlg.open) dlg.showModal();
  try {
    const f = await gasCall('getQuoteFile', fileId);
    const bytes = Uint8Array.from(atob(f.data), c => c.charCodeAt(0));
    const url = URL.createObjectURL(new Blob([bytes], { type: f.mime || 'application/octet-stream' }));
    $('#fileViewerTitle').textContent = f.name;
    const dl = `<a class="btn sm" href="${url}" download="${esc(f.name)}">ดาวน์โหลด</a>`;
    $('#fileViewerBody').innerHTML = /^image\//.test(f.mime) ? `<img src="${url}" alt="${esc(f.name)}">${dl}`
      : /pdf/.test(f.mime) ? `<iframe src="${url}" title="${esc(f.name)}"></iframe>${dl}`
      : `<p class="hint">ไฟล์ประเภทนี้เปิดดูในเบราว์เซอร์ไม่ได้</p>${dl}`;
  } catch (e) {
    $('#fileViewerBody').innerHTML = `<p class="over">เปิดไฟล์ไม่ได้: ${esc(errMsg(e))}</p>`;
  }
}
document.addEventListener('click', e => {
  const f = e.target.closest('[data-file]');
  if (f) { e.preventDefault(); openQuote(f.dataset.file); return; }
  if (e.target.closest('[data-closeviewer]')) $('#fileViewer').close();
});

// ---- PC round confirm form (Admin → PC owners)

function renderConfirmForm(r) {
  const pcDefault = (masterPc.find(p => p.pc_code === r.pc_code) || {}).default_budget_key || '';
  const cur = r.budget_key || pcDefault;
  const opts = hasBudget()
    ? budgetLines().map(l => `<option value="${esc(l.key)}" ${l.key === cur ? 'selected' : ''}>${esc(l.label)}</option>`).join('')
    : '<option value="">— ยังไม่มีข้อมูล Budget —</option>';
  const m = Number(r.budget_month) || new Date().getMonth() + 1;
  return `<div class="confirm-form" data-form="${esc(r.request_id)}">
    <p class="hint">แก้จำนวนได้ (ใส่ 0 = ไม่เอารายการนั้น) แล้วเลือก Budget และเดือนที่จะใช้ของ</p>
    ${reqItemsTable(r, true)}
    <div class="submit-grid" style="margin-top:10px">
      <div class="field"><label>Budget</label>
        <input type="search" data-bsearch placeholder="ค้นหา Location / GL เช่น 530090 cookies">
        <select data-bsel size="5">${opts}</select></div>
      <div class="field"><label>ใช้ของเดือน</label>
        <select data-msel>${THAI_MONTH_FULL.slice(1).map((n, i) => `<option value="${i + 1}" ${i + 1 === m ? 'selected' : ''}>${n} ${budgetOpts ? budgetOpts.year || '' : ''}</option>`).join('')}</select>
        <div class="budget-info" data-binfo></div>
        <div data-overbox hidden><label class="over-req"><input type="checkbox" data-cover> ขอ Over Budget</label>
          <input class="over-reason" data-creason placeholder="เหตุผลที่ขอเกินงบ (จำเป็น)"></div>
        <label>หมายเหตุถึง Admin</label><textarea rows="2" data-note>${esc(r.note)}</textarea>
        <div class="req-actions"><button class="btn" data-act="confirm">ยืนยันส่งให้ Admin</button><button class="btn ghost" data-act="closeedit">ปิด</button></div>
      </div>
    </div></div>`;
}

function updateConfirmInfo(form) {
  let total = 0;
  form.querySelectorAll('tr[data-sku]').forEach(tr => {
    const q = tr.querySelector('[data-eq]');
    if (!q) return;
    const amt = Math.max(0, Math.round(Number(q.value) || 0)) * Number(q.dataset.cost);
    tr.querySelector('[data-amt]').textContent = fmt(amt, 2);
    total += amt;
  });
  const line = budgetLine(form.querySelector('[data-bsel]').value);
  const mon = Number(form.querySelector('[data-msel]').value);
  const box = form.querySelector('[data-binfo]');
  const overBox = form.querySelector('[data-overbox]');
  if (!line) { overBox.hidden = true; box.innerHTML = `<div class="row total"><span>ยอดรวม</span><b>${money(total)}</b></div>`; return; }
  const b = budgetAvail(line, mon);
  overBox.hidden = !(b.available - total < -0.004);
  box.innerHTML = `<div class="row"><span>งบเดือน${THAI_MONTH_FULL[mon]}</span><b>${money(b.plan)}</b></div>
    <div class="row"><span>ใช้จริง + คำขออื่น</span><b>${money(b.actual + b.reserved)}</b></div>
    <div class="row"><span>คงเหลือ</span><b>${money(b.available)}</b></div>
    <div class="row total"><span>ยอดรวมคำขอนี้</span><b>${money(total)}</b></div>
    <div class="row"><span>หลังยืนยัน</span><b class="${b.available - total < 0 ? 'over' : 'ok'}">${money(b.available - total)}${b.available - total < 0 ? ' (งบไม่พอ — ติ๊กขอ Over Budget พร้อมเหตุผล)' : ''}</b></div>`;
}

$('#reqFilter').addEventListener('change', () => { reqFilterSet = true; renderRequests(); });
$('#reqSearch').addEventListener('input', renderRequests);
$('#btnReloadReq').addEventListener('click', () => { loadRequests(); refreshBudget(); });
$('#reqList').addEventListener('input', e => {
  const aform = e.target.closest('[data-aform]');
  if (aform && !e.target.matches('[data-areason]')) { updateAdminEditBudget(aform); return; }
  const form = e.target.closest('[data-form]');
  if (!form) return;
  if (e.target.matches('[data-bsearch]')) {
    const words = e.target.value.trim().toLowerCase().split(/\s+/).filter(Boolean);
    const sel = form.querySelector('[data-bsel]'), cur = sel.value;
    sel.innerHTML = budgetLines().filter(l => words.every(w => (l.label + ' ' + l.expenseGroup).toLowerCase().includes(w)))
      .map(l => `<option value="${esc(l.key)}" ${l.key === cur ? 'selected' : ''}>${esc(l.label)}</option>`).join('');
  }
  updateConfirmInfo(form);
});
$('#reqList').addEventListener('change', e => {
  if (e.target.matches('[data-sel]')) {
    const id = e.target.closest('[data-id]').dataset.id;
    e.target.checked ? selectedReqs.add(id) : selectedReqs.delete(id);
    updateBulkBar();
    return;
  }
  const aform = e.target.closest('[data-aform]');
  if (aform && !e.target.matches('[data-areason]')) { updateAdminEditBudget(aform); return; }
  const form = e.target.closest('[data-form]'); if (form) updateConfirmInfo(form);
});

$('#reqList').addEventListener('click', async e => {
  const btn = e.target.closest('[data-act]');
  if (!btn) return;
  const id = btn.closest('[data-id]').dataset.id, act = btn.dataset.act;
  const r = requests.find(x => x.request_id === id);
  if (act === 'edit' || act === 'aedit') {
    editingReq = { id, mode: act === 'edit' ? 'pc' : 'admin' }; composing = null; renderRequests();
    const form = document.querySelector(`[data-form="${CSS.escape(id)}"], [data-aform="${CSS.escape(id)}"]`);
    if (form) { form.matches('[data-form]') ? updateConfirmInfo(form) : updateAdminEditBudget(form); form.scrollIntoView({ block: 'start' }); }
    return;
  }
  if (act === 'closeedit') { editingReq = null; renderRequests(); return; }
  if (act === 'compose' || act === 'purchasing') { editingReq = null; return openComposer(id, act === 'purchasing' ? 'purchasing' : 'approval'); }
  if (act === 'closecompose') { composing = null; renderRequests(); return; }
  if (act === 'previewmail') { const v = composerValues(id); composing.draft = { ...composing.draft, ...v }; return openComposer(id, composing.kind, v); }
  if (act === 'recall') {
    if (r.status === ST.PENDING && !confirm(`เรียก ${id} กลับมาแก้ไข?\nคำขอจะออกจากคิว Admin จนกว่าจะกดส่งใหม่`)) return;
    btn.disabled = true;
    try {
      const full = await gasCall('recallRequest', id);
      if (loadIntoDraft(full)) { toast(`เปิด ${id} ในหน้าสั่งซื้อแล้ว — แก้ไขแล้วกด "บันทึกและส่งใหม่"`); goTab('order'); }
      loadRequests(true); refreshBudget();
    } catch (err) { toast('ไม่สำเร็จ: ' + errMsg(err)); btn.disabled = false; }
    return;
  }

  let call, done = 'บันทึกแล้ว — แจ้งผู้เกี่ยวข้องทางอีเมลแล้ว';
  if (act === 'confirm') {
    const form = btn.closest('[data-form]');
    const budgetKey = form.querySelector('[data-bsel]').value, month = Number(form.querySelector('[data-msel]').value);
    if (!budgetKey) return toast('กรุณาเลือก Budget');
    const items = [...form.querySelectorAll('tr[data-sku]')].map(tr => ({
      sku: tr.dataset.sku, qty: Math.max(0, Math.round(Number(tr.querySelector('[data-eq]').value) || 0)), note: tr.querySelector('[data-en]').value
    }));
    const overNeeded = !form.querySelector('[data-overbox]').hidden;
    const overReason = form.querySelector('[data-creason]').value.trim();
    if (overNeeded && !(form.querySelector('[data-cover]').checked && overReason)) return toast('งบไม่พอ — ติ๊ก "ขอ Over Budget" และใส่เหตุผล หรือแก้จำนวน/Budget');
    if (!confirm(`ยืนยัน ${id} ส่งให้ Admin ตรวจ?${overNeeded ? '\n(ขอ Over Budget)' : ''}`)) return;
    call = () => gasCall('confirmRequest', id, { items, budgetKey, month, note: form.querySelector('[data-note]').value, overReason: overNeeded ? overReason : '' });
  } else if (act === 'asave') {
    const form = btn.closest('[data-aform]');
    const items = adminFormItems(form).filter(i => i.qty > 0);
    if (!items.length) return toast('ต้องมีอย่างน้อย 1 รายการ — ถ้าไม่ต้องการคำขอนี้ ให้กด "ลบ"');
    const overReasons = {};
    const missing = [...form.querySelectorAll('[data-areason]')].filter(inp => { overReasons[inp.dataset.areason] = inp.value.trim(); return !inp.value.trim(); });
    if (missing.length) return toast('งบไม่พอ — ใส่เหตุผล Over Budget หรือเปลี่ยน Budget / เดือน / จำนวน');
    if (!confirm(`บันทึกการแก้ไข ${id}?\nผู้ขอจะได้รับอีเมลแจ้ง`)) return;
    call = () => gasCall('adminUpdateRequest', id, { items, overReasons, adminNote: form.querySelector('[data-anote]').value.trim() });
    done = 'บันทึกการแก้ไขแล้ว — แจ้งผู้ขอแล้ว';
  } else if (act === 'pcreject') {
    const note = prompt(`${r.pc_code}: ไม่สั่งรอบนี้ — เหตุผล (ไม่บังคับ)`, ''); if (note === null) return;
    call = () => gasCall('rejectByPc', id, note);
  } else if (act === 'sendmail') {
    const v = composerValues(id);
    if (!v.to || !v.to.includes('@')) return toast('กรุณาใส่อีเมลผู้รับ (ถึง)');
    const purch = composing.kind === 'purchasing';
    const nAtt = (v.attachExcel ? 1 : 0) + v.attachFileIds.length;
    if (!confirm(`ส่งอีเมล "${v.subject}"\nถึง: ${v.to}${v.cc ? '\nสำเนา: ' + v.cc : ''}\nไฟล์แนบ ${nAtt} ไฟล์\n\nยืนยันส่ง?`)) return;
    call = () => gasCall(purch ? 'sendPurchasingEmail' : 'sendApprovalEmail', id, v);
    done = purch ? 'ส่งแจ้งฝ่ายจัดซื้อแล้ว' : `ส่งเมลรายการสั่งซื้อ + Budget ถึง ${v.to} แล้ว`;
  } else if (act === 'approvedok' || act === 'approvedno') {
    const ok = act === 'approvedok';
    const note = prompt(ok ? `${id}: ได้รับอนุมัติ — หมายเหตุ (เช่น "Approved" ทางอีเมล วันที่…)` : `${id}: เหตุผลที่ไม่อนุมัติ`, ok ? 'Approved' : '');
    if (note === null) return;
    call = () => gasCall('recordApproval', id, ok, note);
  } else if (act === 'return') {
    const note = prompt(`ส่ง ${id} กลับให้ผู้ขอแก้ไข — บอกสิ่งที่ต้องแก้:`, ''); if (!note) return;
    call = () => gasCall('returnForEdit', id, note);
    done = 'ส่งกลับให้ผู้ขอแก้ไขแล้ว';
  } else if (act === 'reject') {
    const note = prompt(`เหตุผลที่ไม่อนุมัติ ${id}:`); if (note === null) return;
    call = () => gasCall('decideRequest', id, 'reject', note);
  } else if (act === 'delete') {
    if (!confirm(`ลบคำขอ ${id} (${r.status}, ${money(r.total)}) ทั้งหมด?\nรายการสินค้าจะถูกลบ ใบเสนอราคาย้ายไปถังขยะของ Drive — ย้อนกลับไม่ได้ในระบบ`)) return;
    call = () => gasCall('deleteRequest', id);
    done = 'ลบคำขอแล้ว';
  } else if (act === 'mgrok' || act === 'mgrno') {
    const note = prompt(act === 'mgrok' ? 'อนุมัติเกินงบ — หมายเหตุ (ไม่บังคับ)' : 'เหตุผลที่ไม่อนุมัติ:', ''); if (note === null) return;
    call = () => gasCall('managerDecision', id, act === 'mgrok', note);
  } else if (act === 'po') {
    const po = prompt(`เลขที่ PR/PO ของ ${id}:`, ''); if (!po) return;
    call = () => gasCall('issuePo', id, po);
  } else if (act === 'received') {
    const note = prompt(`รับของตาม ${id} เข้า Store แล้ว — หมายเหตุ (ไม่บังคับ)`, ''); if (note === null) return;
    call = () => gasCall('markReceived', id, note);
  } else if (act === 'cancel') {
    if (!confirm(`ยกเลิกคำขอ ${id}?`)) return;
    call = () => gasCall('cancelMyRequest', id);
  } else return;

  btn.disabled = true;
  try {
    requests = await call();
    editingReq = null; composing = null;
    if (state.draft.editId === id && [ 'delete', 'cancel' ].includes(act)) resetDraft();
    updateReqBadge(); renderRequests(); refreshBudget();
    toast(`${id}: ${done}`);
  } catch (err) {
    toast('ไม่สำเร็จ: ' + errMsg(err)); btn.disabled = false;
  }
});

// Admin: send the draft to each PC's owners for confirmation
$('#btnRound').addEventListener('click', async () => {
  const lines = state.draft.lines.filter(l => Number(l.qty) > 0);
  if (!lines.length) return toast('ยังไม่มีรายการ');
  const pcs = [...new Set(lines.map(l => l.pc || (lineStore(l) || {}).pc || ''))];
  const noOwner = pcs.filter(pc => !(masterPc.find(p => p.pc_code === pc) || {}).owner_emails);
  if (!confirm(`ส่ง ${lines.length} รายการ แยก ${pcs.length} PC ให้ผู้รับผิดชอบยืนยันทางอีเมล?` +
    (noOwner.length ? `\n\n⚠ ยังไม่มีอีเมลผู้รับผิดชอบใน Master PC: ${noOwner.join(', ')}\n(สร้างคำขอไว้ แต่ยังไม่ส่งอีเมล — ตั้งในสมาชิก & ตั้งค่า)` : ''))) return;
  $('#btnRound').disabled = true;
  try {
    const res = await gasCall('createRound', lines.map(l => ({
      sku: l.sku, name: lineName(l), pc: l.pc || (lineStore(l) || {}).pc || '', pcName: l.pcName || '', unit: lineUnit(l), qty: Number(l.qty), unitCost: lineCost(l), note: l.note || ''
    })), $('#roundNote').value);
    resetDraft(); $('#roundNote').value = ''; renderCards(); renderOrder();
    toast(`สร้างรอบ ${res.roundId}: ${res.created.length} PC` + (res.noOwner.length ? ` · ยังไม่ส่งอีเมล: ${res.noOwner.join(', ')}` : ' · ส่งอีเมลครบแล้ว'));
    $('#reqFilter').value = ST.WAIT_PC; reqFilterSet = true;
    await loadRequests(true);
    goTab('requests');
  } catch (e) {
    toast('ส่งไม่สำเร็จ: ' + errMsg(e));
  } finally {
    $('#btnRound').disabled = false;
  }
});

// ------------------------------------------------------------ Master PC (owner + manager emails per PC)

async function loadMasterPc() {
  try { masterPc = await gasCall('listMasterPc'); } catch (e) { masterPc = []; }
  if (activeTab === 'admin') renderMasterPc();
  if (activeTab === 'home') renderHome();
}

function renderMasterPc() {
  const known = new Map(masterPc.map(p => [p.pc_code, p]));
  // PCs seen in the Store data but not yet in master_pc
  for (const r of merged.rows) if (r.pc && !known.has(r.pc)) known.set(r.pc, { pc_code: r.pc, pc_name: r.pcName, owner_emails: '', manager_email: '', default_budget_key: '' });
  const lines = budgetLines();
  const list = [...known.values()].sort((a, b) => a.pc_code.localeCompare(b.pc_code));
  $('#tblMasterPc').innerHTML = `<thead><tr><th>PC</th><th>ผู้รับผิดชอบ (อีเมล, คั่นด้วย ,)</th><th>หัวหน้า (อนุมัติเกินงบ)</th><th>Budget ตั้งต้น</th><th></th></tr></thead>
    <tbody>${list.map(p => `<tr data-pc="${esc(p.pc_code)}" class="${p.owner_emails ? '' : 'sel'}">
      <td class="sku">${esc(p.pc_code)}<div class="pc">${esc(p.pc_name)}</div></td>
      <td><input data-f="owner_emails" value="${esc(p.owner_emails)}" placeholder="name@planbmedia.co.th"></td>
      <td><input data-f="manager_email" value="${esc(p.manager_email)}"></td>
      <td><select data-f="default_budget_key"><option value="">—</option>${lines.map(l => `<option value="${esc(l.key)}" ${l.key === p.default_budget_key ? 'selected' : ''}>${esc(l.label)}</option>`).join('')}</select></td>
      <td><button class="btn sm ghost" data-savepc>บันทึก</button></td>
    </tr>`).join('')}</tbody>`;
}

$('#tblMasterPc').addEventListener('click', async e => {
  const btn = e.target.closest('[data-savepc]');
  if (!btn) return;
  const tr = btn.closest('tr');
  const val = f => tr.querySelector(`[data-f="${f}"]`).value;
  btn.disabled = true;
  try {
    masterPc = await gasCall('saveMasterPc', { pc_code: tr.dataset.pc, pc_name: tr.querySelector('.pc').textContent,
      owner_emails: val('owner_emails'), manager_email: val('manager_email'), default_budget_key: val('default_budget_key') });
    toast(`บันทึก ${tr.dataset.pc} แล้ว`); renderMasterPc();
  } catch (err) { toast('บันทึกไม่สำเร็จ: ' + errMsg(err)); btn.disabled = false; }
});
$('#btnTrigger').addEventListener('click', async () => {
  try { toast(await gasCall('installReminderTrigger')); } catch (e) { toast('ตั้งไม่สำเร็จ: ' + errMsg(e)); }
});

// ------------------------------------------------------------ admin: members + budget status

let members = null;

let appSettings = null;
async function loadSettings() {
  const f = $('#settingsForm');
  if (!appSettings) {
    $('#settingsStatus').textContent = 'กำลังโหลด…';
    try { appSettings = await gasCall('getAppSettings'); } catch (e) { $('#settingsStatus').textContent = 'โหลดไม่ได้: ' + errMsg(e); return; }
  }
  [...f.elements].forEach(el => { if (el.name && appSettings[el.name] != null && document.activeElement !== el) el.value = appSettings[el.name]; });
  $('#settingsStatus').textContent = appSettings.approval_to ? '' : 'ยังไม่ได้ตั้งอีเมลผู้อนุมัติ';
}
$('#settingsForm').addEventListener('submit', async e => {
  e.preventDefault();
  const v = {};
  [...e.target.elements].forEach(el => { if (el.name) v[el.name] = el.value; });
  try { appSettings = await gasCall('saveAppSettings', v); toast('บันทึกการตั้งค่าอีเมลแล้ว'); loadSettings(); }
  catch (err) { toast('บันทึกไม่สำเร็จ: ' + errMsg(err)); }
});

async function renderAdmin() {
  if (!isAdmin()) return;
  renderMasterPc();
  loadSettings();
  $('#budgetStatus').textContent = !budgetOpts ? 'กำลังโหลด…'
    : budgetOpts.lines.length ? `มี Budget ${budgetOpts.lines.length} รายการ ปี ${budgetOpts.year || '-'}` : 'ยังไม่มีข้อมูล Budget';
  if (!members) {
    $('#tblMembers').innerHTML = '<tbody><tr><td class="empty">กำลังโหลด…</td></tr></tbody>';
    try { members = await gasCall('listMembers'); } catch (e) { toast(errMsg(e)); return; }
  }
  const order = { pending: 0, active: 1, disabled: 2 };
  const list = [...members].sort((a, b) => order[a.status] - order[b.status] || a.username.localeCompare(b.username));
  $('#tblMembers').innerHTML = `<thead><tr><th>ID</th><th>ชื่อ</th><th>อีเมล</th><th>สิทธิ์</th><th>สถานะ</th><th></th></tr></thead>
    <tbody>${list.map(m => `<tr data-user="${esc(m.username)}" class="${m.status === 'pending' ? 'sel' : ''}">
      <td class="sku">${esc(m.username)}${m.mustChange ? '<div class="pc">รอเปลี่ยนรหัส</div>' : ''}</td>
      <td><input data-f="name" value="${esc(m.name)}"></td>
      <td><input data-f="email" type="email" value="${esc(m.email)}" placeholder="ใช้รับแจ้งเตือน"></td>
      <td><select data-f="role"><option value="user" ${m.role === 'user' ? 'selected' : ''}>User</option><option value="admin" ${m.role === 'admin' ? 'selected' : ''}>Admin</option></select></td>
      <td><select data-f="status">${['pending', 'active', 'disabled'].map(s => `<option value="${s}" ${m.status === s ? 'selected' : ''}>${{ pending: 'รออนุมัติ', active: 'ใช้งาน', disabled: 'ระงับ' }[s]}</option>`).join('')}</select></td>
      <td class="nowrap">${m.status === 'pending' ? '<button class="btn sm" data-save="approve">อนุมัติ</button> ' : ''}<button class="btn sm ghost" data-save="1">บันทึก</button>
        <button class="btn sm ghost" data-save="reset">ตั้งรหัสใหม่</button></td>
    </tr>`).join('')}</tbody>`;
  renderPasswords();
}

$('#tblMembers').addEventListener('click', async e => {
  const btn = e.target.closest('[data-save]');
  if (!btn) return;
  const tr = btn.closest('tr');
  const val = f => tr.querySelector(`[data-f="${f}"]`).value;
  const input = { username: tr.dataset.user, name: val('name'), email: val('email'), role: val('role'),
    status: btn.dataset.save === 'approve' ? 'active' : val('status') };
  if (btn.dataset.save === 'reset') {
    const pw = prompt(`ตั้งรหัสผ่านชั่วคราวให้ ${tr.dataset.user} (อย่างน้อย 8 ตัว — ผู้ใช้จะถูกบังคับเปลี่ยนตอนเข้าสู่ระบบ):`);
    if (!pw) return;
    input.password = pw;
  }
  btn.disabled = true;
  try {
    members = await gasCall('saveMember', input);
    toast(btn.dataset.save === 'reset' ? 'ตั้งรหัสผ่านใหม่แล้ว — แจ้งผู้ใช้ให้เข้าสู่ระบบแล้วเปลี่ยนรหัส' : 'บันทึกแล้ว');
    renderAdmin();
  } catch (err) { toast('บันทึกไม่สำเร็จ: ' + errMsg(err)); btn.disabled = false; }
});
$('#memberAdd').addEventListener('submit', async e => {
  e.preventDefault();
  const f = e.target;
  try {
    members = await gasCall('saveMember', { username: f.username.value, password: f.password.value, name: f.name.value,
      email: f.email.value, role: f.role.value, status: 'active' });
    f.reset(); toast('เพิ่มสมาชิกแล้ว — ผู้ใช้จะถูกบังคับเปลี่ยนรหัสผ่านตอนเข้าสู่ระบบครั้งแรก'); renderAdmin();
  } catch (err) { toast('เพิ่มไม่สำเร็จ: ' + errMsg(err)); }
});

// ------------------------------------------------------------ budget page: every line, month by month, used vs left

const MONTHS_1_12 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
let budgetOpen = null; // group key whose monthly detail is expanded

/** plan / actual / pending / committed / left for one line and month. */
function budgetCell(line, m) {
  const mb = line.months[m] || { plan: 0, actual: 0 };
  const d = ((budgetOpts && budgetOpts.reservedDetail) || {})[line.key + '|' + m];
  // older servers only send the total "reserved"; count it as committed
  const pending = d ? d.pending : 0;
  const committed = d ? d.committed : (((budgetOpts && budgetOpts.reserved) || {})[line.key + '|' + m] || 0);
  return { plan: mb.plan, actual: mb.actual, pending, committed, left: mb.plan - mb.actual - pending - committed };
}

const addCell = (a, b) => { for (const k of ['plan', 'actual', 'pending', 'committed', 'left']) a[k] += b[k]; return a; };
const zeroCell = () => ({ plan: 0, actual: 0, pending: 0, committed: 0, left: 0 });

const BUDGET_GROUPS = {
  line: l => [l.key, l.label],
  gl: l => ['gl|' + l.glCode, `${l.glCode} ${l.glName}`],
  location: l => ['loc|' + l.location, l.location],
  company: l => ['co|' + l.company, l.company || '(ไม่ระบุบริษัท)']
};

/** Rows for the current grouping/filters, each with per-month cells and a total for the selected period. */
function budgetRows() {
  const words = $('#bSearch').value.trim().toLowerCase().split(/\s+/).filter(Boolean);
  const month = Number($('#bMonth').value) || 0;
  const by = BUDGET_GROUPS[$('#bGroup').value] || BUDGET_GROUPS.line;
  const groups = new Map();
  for (const l of budgetLines()) {
    const text = (l.label + ' ' + l.expenseGroup).toLowerCase();
    if (!words.every(w => text.includes(w))) continue;
    const [key, label] = by(l);
    const g = groups.get(key) || { key, label, lines: 0, months: {}, total: zeroCell() };
    g.lines++;
    for (const m of MONTHS_1_12) g.months[m] = addCell(g.months[m] || zeroCell(), budgetCell(l, m));
    groups.set(key, g);
  }
  const show = $('#bShow').value;
  return [...groups.values()].map(g => {
    for (const m of month ? [month] : MONTHS_1_12) addCell(g.total, g.months[m]);
    return g;
  }).filter(g => {
    const t = g.total;
    if (!t.plan && !t.actual && !t.pending && !t.committed) return false;
    if (show === 'over') return t.left < 0;
    if (show === 'left') return t.left > 0;
    if (show === 'used') return t.actual + t.committed > 0;
    return true;
  }).sort((a, b) => a.label.localeCompare(b.label));
}

const pct = (a, b) => b > 0 ? Math.round(a / b * 100) : (a > 0 ? 999 : 0);
function useBar(t) {
  const p = pct(t.actual + t.committed + t.pending, t.plan);
  const w = n => Math.min(100, t.plan > 0 ? n / t.plan * 100 : (n > 0 ? 100 : 0));
  return `<div class="use-bar ${t.left < 0 ? 'over' : ''}" title="ใช้ ${p}%">
    <i class="a" style="width:${w(t.actual)}%"></i><i class="c" style="left:${w(t.actual)}%;width:${Math.max(0, Math.min(100 - w(t.actual), w(t.committed + t.pending)))}%"></i></div>
    <span class="use-pct ${t.left < 0 ? 'over' : ''}">${p}%</span>`;
}

function renderBudget() {
  if (!$('#bMonth').options.length || $('#bMonth').options.length < 13) {
    $('#bMonth').innerHTML = '<option value="">ทั้งปี</option>' + MONTHS_1_12.map(m => `<option value="${m}">${THAI_MONTH_FULL[m]}</option>`).join('');
  }
  if (!budgetOpts) { $('#budgetEmpty').innerHTML = '<div class="empty">กำลังโหลด Budget…</div>'; $('#budgetBody').hidden = true; return; }
  if (budgetOpts.error || !budgetOpts.lines.length) {
    $('#budgetBody').hidden = true;
    $('#budgetEmpty').innerHTML = `<div class="alert-row"><span>${budgetOpts.error ? 'โหลด Budget ไม่ได้: ' + esc(budgetOpts.error)
      : 'ยังไม่มีข้อมูล Budget ในฐานข้อมูล' + (isAdmin() ? ' — นำเข้าที่ ตั้งค่าระบบ > ข้อมูล Budget' : ' — รอ Admin นำเข้า')}</span>
      ${isAdmin() ? '<button class="btn sm" data-go="admin">ไปนำเข้า</button>' : ''}</div>`;
    return;
  }
  $('#budgetEmpty').innerHTML = '';
  $('#budgetBody').hidden = false;
  const month = Number($('#bMonth').value) || 0;
  const rows = budgetRows();
  const total = rows.reduce((a, g) => addCell(a, g.total), zeroCell());
  const overRows = rows.filter(g => g.total.left < 0);
  $('#bAllGl').checked = showAllGl;
  const hidden = glHidden();
  $('#budgetSub').textContent = `ปีงบ ${budgetOpts.year || '-'} · ${fmt(budgetLines().length)} รายการ Budget · ${month ? THAI_MONTH_FULL[month] : 'ทั้งปี'}`
    + (hidden ? ` · ซ่อนหมวดอื่นไว้ ${fmt(hidden)} รายการ (แสดงเฉพาะ ค่าซ่อมแซมบำรุงรักษา · ค่าวัสดุสิ้นเปลืองใช้ไป · ค่าซ่อมภาพโฆษณา)` : '');

  $('#budgetCards').innerHTML = `
    <div class="card"><div class="k">งบ${month ? 'เดือน' + THAI_MONTH_FULL[month] : 'ทั้งปี'}</div><div class="v">${money(total.plan)}</div><div class="s">${fmt(rows.length)} รายการที่แสดง</div></div>
    <div class="card actual"><div class="k">ใช้จริงแล้ว (Actual)</div><div class="v">${money(total.actual)}</div><div class="s">${pct(total.actual, total.plan)}% ของงบ</div></div>
    <div class="card store"><div class="k">คำขอที่อนุมัติ / ออก PO</div><div class="v">${money(total.committed)}</div><div class="s">รออนุมัติ ${money(total.pending)}</div></div>
    <div class="card ${total.left < 0 ? 'dead' : 'actual'}"><div class="k">คงเหลือ</div><div class="v ${total.left < 0 ? 'over' : ''}">${money(total.left)}</div><div class="s">${pct(total.left, total.plan)}% ของงบ</div></div>
    <div class="card dead"><div class="k">รายการที่เกินงบ</div><div class="v">${fmt(overRows.length)}</div><div class="s">เกินรวม ${money(-overRows.reduce((a, g) => a + g.total.left, 0))}</div></div>`;

  // month chart across everything shown
  const byMonth = MONTHS_1_12.map(m => rows.reduce((a, g) => addCell(a, g.months[m]), zeroCell()));
  const max = Math.max(1, ...byMonth.map(c => Math.max(c.plan, c.actual + c.committed + c.pending)));
  $('#budgetMonths').innerHTML = byMonth.map((c, i) => {
    const h = n => Math.round(n / max * 100);
    return `<button class="mc ${month === i + 1 ? 'sel' : ''} ${c.left < 0 ? 'over' : ''}" data-month="${i + 1}" title="${THAI_MONTH_FULL[i + 1]}\nงบ ${money(c.plan)}\nใช้จริง ${money(c.actual)}\nคำขอ ${money(c.committed + c.pending)}\nคงเหลือ ${money(c.left)}">
      <span class="mc-bars"><i class="p" style="height:${h(c.plan)}%"></i><i class="u" style="height:${h(c.actual)}%"></i><i class="r" style="height:${h(c.committed + c.pending)}%;bottom:${h(c.actual)}%"></i></span>
      <span class="mc-label">${MONTH_SHORT[i]}</span><span class="mc-left ${c.left < 0 ? 'over' : ''}">${fmtK(c.left)}</span></button>`;
  }).join('');

  const top = (list, fmtRow) => list.length ? `<ul class="todo-list">${list.map(fmtRow).join('')}</ul>` : '<div class="empty-sm">ไม่มี</div>';
  $('#budgetOver').innerHTML = top([...overRows].sort((a, b) => a.total.left - b.total.left).slice(0, 6), g =>
    `<li data-bopen="${esc(g.key)}"><span class="todo-main">${esc(g.label)}</span><span class="num over">${money(g.total.left)}</span></li>`);
  $('#budgetLeft').innerHTML = top(rows.filter(g => g.total.left > 0).sort((a, b) => b.total.left - a.total.left).slice(0, 6), g =>
    `<li data-bopen="${esc(g.key)}"><span class="todo-main">${esc(g.label)}</span><span class="num ok">${money(g.total.left)}</span></li>`);

  $('#bInfo').textContent = `${fmt(rows.length)} แถว`;
  $('#tblBudget').innerHTML = `<thead><tr>
      <th>${esc($('#bGroup').selectedOptions[0].text.replace('แยกตาม', '').replace('รวมตาม', ''))}</th>
      <th class="n">งบ</th><th class="n">ใช้จริง</th><th class="n">คำขอ (อนุมัติ/รอ)</th><th class="n">คงเหลือ</th><th>ใช้ไป</th></tr></thead>
    <tbody>${rows.map(g => {
      const t = g.total;
      const open = budgetOpen === g.key;
      return `<tr class="brow ${open ? 'sel' : ''}" data-bkey="${esc(g.key)}">
          <td class="name"><span class="caret">${open ? '▾' : '▸'}</span> ${esc(g.label)}${g.lines > 1 ? `<div class="pc">${g.lines} รายการ</div>` : ''}
            ${budgetLine(g.key) ? `<button class="btn xs" data-order-line="${esc(g.key)}">สั่งซื้อในงบนี้</button>` : ''}</td>
          <td class="n">${fmt(t.plan)}</td><td class="n">${fmt(t.actual)}</td>
          <td class="n">${fmt(t.committed)}${t.pending ? `<div class="pc">รอ ${fmt(t.pending)}</div>` : ''}</td>
          <td class="n ${t.left < 0 ? 'over' : 'ok'}"><b>${fmt(t.left)}</b></td><td class="bar-cell">${useBar(t)}</td></tr>
        ${open ? `<tr class="bdetail"><td colspan="6">${budgetMonthTable(g)}</td></tr>` : ''}`;
    }).join('')}</tbody>
    <tfoot><tr><th>รวม</th><th class="n">${fmt(total.plan)}</th><th class="n">${fmt(total.actual)}</th><th class="n">${fmt(total.committed + total.pending)}</th>
      <th class="n ${total.left < 0 ? 'over' : 'ok'}">${fmt(total.left)}</th><th>${useBar(total)}</th></tr></tfoot>`;
}

function budgetMonthTable(g) {
  const cur = new Date().getMonth() + 1;
  let run = 0;
  return `<table class="grid month-table"><thead><tr><th>เดือน</th><th class="n">งบ</th><th class="n">ใช้จริง</th><th class="n">อนุมัติ/PO</th><th class="n">รออนุมัติ</th><th class="n">คงเหลือเดือนนี้</th><th class="n">คงเหลือสะสม</th><th>ใช้ไป</th></tr></thead>
    <tbody>${MONTHS_1_12.map(m => {
      const c = g.months[m];
      run += c.left;
      return `<tr class="${m === cur ? 'now' : ''}"><td>${THAI_MONTH_FULL[m]}${m === cur ? ' <span class="tag st-pending">เดือนนี้</span>' : ''}</td>
        <td class="n">${fmt(c.plan)}</td><td class="n">${fmt(c.actual)}</td><td class="n">${fmt(c.committed)}</td><td class="n">${fmt(c.pending)}</td>
        <td class="n ${c.left < 0 ? 'over' : ''}">${fmt(c.left)}</td><td class="n ${run < 0 ? 'over' : ''}">${fmt(run)}</td><td class="bar-cell">${useBar(c)}</td></tr>`;
    }).join('')}</tbody></table>`;
}

const fmtK = n => Math.abs(n) >= 1e6 ? (n / 1e6).toFixed(1) + 'M' : Math.abs(n) >= 1e3 ? Math.round(n / 1e3) + 'K' : fmt(n);

['#bSearch', '#bGroup', '#bMonth', '#bShow'].forEach(s => $(s).addEventListener('input', () => { budgetOpen = null; renderBudget(); }));
$('#tblBudget').addEventListener('click', e => {
  const ord = e.target.closest('[data-order-line]');
  if (ord) {
    state.pick = { key: ord.dataset.orderLine, month: Number($('#bMonth').value) || new Date().getMonth() + 1 };
    save(); goTab('order'); return;
  }
  const tr = e.target.closest('tr.brow');
  if (!tr) return;
  budgetOpen = budgetOpen === tr.dataset.bkey ? null : tr.dataset.bkey;
  renderBudget();
});
$('#budgetMonths').addEventListener('click', e => {
  const b = e.target.closest('[data-month]');
  if (!b) return;
  $('#bMonth').value = $('#bMonth').value === b.dataset.month ? '' : b.dataset.month;
  renderBudget();
});
document.addEventListener('click', e => {
  const o = e.target.closest('[data-bopen]');
  if (!o) return;
  budgetOpen = o.dataset.bopen;
  renderBudget();
  const row = document.querySelector(`tr.brow[data-bkey="${CSS.escape(budgetOpen)}"]`);
  row && row.scrollIntoView({ block: 'center' });
});
$('#btnBudgetReload').addEventListener('click', () => { budgetOpts = null; renderBudget(); refreshBudget(); });
$('#bAllGl').addEventListener('change', e => {
  showAllGl = e.target.checked;
  try { localStorage.setItem(GL_ALL_KEY, showAllGl ? '1' : '0'); } catch { /* private mode */ }
  renderBudget();
  if (activeTab === 'order') renderOrder(); // หน้าสั่งซื้อใช้รายการเดียวกัน
});
$('#btnBudgetExport').addEventListener('click', () => {
  if (!budgetOpts || !budgetOpts.lines.length) return toast('ยังไม่มีข้อมูล Budget');
  const out = [];
  for (const l of budgetLines()) for (const m of MONTHS_1_12) {
    const c = budgetCell(l, m);
    if (!c.plan && !c.actual && !c.pending && !c.committed) continue;
    out.push({ 'บริษัท': l.company, 'Location': l.location, 'GL Code': l.glCode, 'ประเภท': l.glName, 'Expense Group': l.expenseGroup,
      'เดือน': m, 'งบ': c.plan, 'ใช้จริง': c.actual, 'อนุมัติ/PO': c.committed, 'รออนุมัติ': c.pending, 'คงเหลือ': Math.round(c.left * 100) / 100 });
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(out), 'Budget รายเดือน');
  XLSX.writeFile(wb, `Budget_สรุป_${stamp()}.xlsx`);
});

// Admin: import the latest "Budget STT…" file straight from Google Drive (server reads it)
$('#btnBudgetDrive').addEventListener('click', async () => {
  const btn = $('#btnBudgetDrive');
  btn.disabled = true; $('#budgetStatus').textContent = 'กำลังหาไฟล์ใน Google Drive และนำเข้า…';
  try { toast(await gasCall('importBudgetFromDrive')); await refreshBudget(); }
  catch (e) { toast('นำเข้าไม่สำเร็จ: ' + errMsg(e)); renderAdmin(); }
  finally { btn.disabled = false; }
});

// ------------------------------------------------------------ admin: passwords the admin can actually read
// The sheet only keeps a salted hash, so an existing password can never be shown again.
// What the admin can do is set a new one - and that one is shown here (this page only, in memory).

const pwShown = new Map(); // username -> { password, at, mustChange }

function pwForce() { return $('#pwForce') && $('#pwForce').checked; }

function renderPasswords() {
  const t = $('#tblPasswords');
  if (!t) return;
  if (!members) { t.innerHTML = '<tbody><tr><td class="empty">กำลังโหลด…</td></tr></tbody>'; return; }
  const order = { admin: 0, user: 1 };
  const list = [...members].sort((a, b) => order[a.role] - order[b.role] || a.username.localeCompare(b.username));
  t.innerHTML = `<thead><tr><th>ID เข้าระบบ</th><th>ชื่อ</th><th>สิทธิ์</th><th>สถานะ</th><th>รหัสผ่าน</th><th></th></tr></thead>
    <tbody>${list.map(m => {
      const s = pwShown.get(m.username);
      return `<tr data-user="${esc(m.username)}">
        <td class="sku">${esc(m.username)}</td>
        <td>${esc(m.name || '')}</td>
        <td><span class="role-pill sm ${m.role === 'admin' ? 'admin' : 'user'}">${m.role === 'admin' ? 'Admin' : 'User'}</span></td>
        <td>${{ pending: 'รออนุมัติ', active: 'ใช้งาน', disabled: 'ระงับ' }[m.status] || esc(m.status)}</td>
        <td>${s ? `<code class="pw-code">${esc(s.password)}</code>${s.mustChange ? '<div class="pc">ต้องเปลี่ยนตอนเข้าครั้งแรก</div>' : ''}`
          : '<span class="hint">เข้ารหัสไว้ — ดูไม่ได้ กดสร้างรหัสใหม่เพื่อให้เห็น</span>'}</td>
        <td class="nowrap">${s ? '<button class="btn sm ghost" data-pw="copy">คัดลอก</button> ' : ''}<button class="btn sm" data-pw="new">สร้างรหัสใหม่</button></td>
      </tr>`;
    }).join('')}</tbody>`;
}

function pwLines() {
  return [...pwShown.entries()].map(([u, s]) => {
    const m = (members || []).find(x => x.username === u) || {};
    return `${u}\t${s.password}\t${m.role === 'admin' ? 'Admin' : 'User'}\t${m.name || ''}`;
  });
}

async function copyText(text, okMsg) {
  try { await navigator.clipboard.writeText(text); toast(okMsg); }
  catch { prompt('คัดลอกข้อความนี้:', text); } // clipboard blocked (http / iframe)
}

$('#tblPasswords').addEventListener('click', async e => {
  const btn = e.target.closest('[data-pw]');
  if (!btn) return;
  const username = btn.closest('tr').dataset.user;
  if (btn.dataset.pw === 'copy') {
    const s = pwShown.get(username);
    return s && copyText(`${username} / ${s.password}`, 'คัดลอก ID + รหัสผ่านแล้ว');
  }
  if (!confirm(`สร้างรหัสผ่านใหม่ให้ ${username}?\nรหัสเดิมจะใช้ไม่ได้ทันที`)) return;
  btn.disabled = true;
  try {
    const r = await gasCall('resetMemberPassword', { username, forceChange: pwForce() });
    pwShown.set(r.username, { password: r.password, mustChange: r.mustChange, at: Date.now() });
    members = null; await renderAdmin();
    toast(`รหัสใหม่ของ ${r.username} คือ ${r.password}`);
  } catch (err) { toast('สร้างไม่สำเร็จ: ' + errMsg(err)); btn.disabled = false; }
});

$('#btnPwAll').addEventListener('click', async e => {
  if (!confirm('สร้างรหัสผ่านใหม่ให้สมาชิกทุกคน?\nรหัสเดิมของทุกคน (รวมบัญชีที่คุณใช้อยู่) จะใช้ไม่ได้ทันที')) return;
  e.target.disabled = true;
  try {
    const rows = await gasCall('resetMemberPasswords', { forceChange: pwForce() });
    rows.forEach(r => pwShown.set(r.username, { password: r.password, mustChange: r.mustChange, at: Date.now() }));
    members = null; await renderAdmin();
    toast(`สร้างรหัสใหม่ให้ ${rows.length} บัญชีแล้ว — คัดลอกเก็บไว้ก่อนปิดหน้าเว็บ`);
  } catch (err) { toast('สร้างไม่สำเร็จ: ' + errMsg(err)); }
  e.target.disabled = false;
});

$('#btnPwCopy').addEventListener('click', () => {
  const lines = pwLines();
  if (!lines.length) return toast('ยังไม่มีรหัสที่แสดงอยู่ — กด "สร้างรหัสใหม่" ก่อน');
  copyText('ID\tรหัสผ่าน\tสิทธิ์\tชื่อ\n' + lines.join('\n'), `คัดลอก ${lines.length} บัญชีแล้ว`);
});

// 👁 on every password box, so what you type (or paste) can be checked before sending
function enhancePasswords(root) {
  (root || document).querySelectorAll('input[type="password"]').forEach(inp => {
    if (inp.dataset.eye) return;
    inp.dataset.eye = '1';
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'eye-btn';
    btn.title = 'แสดง/ซ่อนรหัสผ่าน';
    btn.textContent = '👁';
    btn.addEventListener('click', () => {
      const show = inp.type === 'password';
      inp.type = show ? 'text' : 'password';
      btn.classList.toggle('on', show);
      inp.focus();
    });
    const wrap = document.createElement('span');
    wrap.className = 'pw-wrap';
    inp.parentNode.insertBefore(wrap, inp);
    wrap.appendChild(inp);
    wrap.appendChild(btn);
  });
}
new MutationObserver(() => enhancePasswords()).observe(document.body, { childList: true, subtree: true });
enhancePasswords();

// ------------------------------------------------------------ PR/PO + ติดตามการส่งของ (server: src/Delivery.js)
// ของในคำขอเดียวกันอาจอยู่คนละ PR/PO และมาคนละวัน จึงเก็บ pr_no / po_no / po_at / eta_date / received_at รายชิ้น

const DAY_MS = 86400000;
const todayISO = () => new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
const parseISO = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || '')); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
/** จำนวนวันจาก a ถึง b (บวก = b อยู่หลัง a) */
function daysBetween(a, b) {
  const x = parseISO(a), y = parseISO(b);
  return x && y ? Math.round((y - x) / DAY_MS) : null;
}
const addDays = (iso, n) => { const d = parseISO(iso); if (!d) return ''; d.setDate(d.getDate() + n); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
const thaiDay = iso => { const d = parseISO(iso); return d ? `${d.getDate()} ${MONTH_SHORT[d.getMonth()]}` : ''; };

/** ทุกรายการสินค้าที่ออก PR/PO แล้ว พร้อมตัวเลขที่ใช้วิเคราะห์ */
function deliveryItems() {
  const out = [];
  for (const r of requests) {
    if (![ST.PO, ST.RECEIVED, ST.APPROVED].includes(r.status)) continue;
    (r.items || []).forEach((i, n) => {
      if (!(Number(i.qty) > 0) || !(i.po_no || i.pr_no)) return;
      const lead = daysBetween(i.po_at, i.received_at);      // ใช้เวลาส่งจริงกี่วัน
      const vsEta = daysBetween(i.eta_date, i.received_at);  // + = มาช้ากว่ากำหนด
      const lateBy = i.received_at ? null : daysBetween(i.eta_date, todayISO());
      out.push({ r, i, row: n, lead, vsEta, lateBy,
        late: i.received_at ? vsEta > 0 : lateBy > 0,
        open: !i.received_at, amount: Number(i.amount) || 0 });
    });
  }
  return out;
}

/** เฉลี่ย / กลาง / p90 ของระยะเวลาส่ง + อัตราตรงเวลา — ใช้ทั้งประมาณ ETA และหน้าวิเคราะห์ */
function deliveryStats(list) {
  const done = list.filter(x => x.lead != null && x.lead >= 0);
  const leads = done.map(x => x.lead).sort((a, b) => a - b);
  const pick = p => leads.length ? leads[Math.min(leads.length - 1, Math.floor(leads.length * p))] : null;
  const withEta = done.filter(x => x.vsEta != null);
  const onTime = withEta.filter(x => x.vsEta <= 0);
  const lateOnes = withEta.filter(x => x.vsEta > 0);
  return {
    n: list.length, done: done.length,
    open: list.filter(x => x.open).length,
    openAmt: list.filter(x => x.open).reduce((a, x) => a + x.amount, 0),
    lateOpen: list.filter(x => x.open && x.late).length,
    avg: leads.length ? leads.reduce((a, b) => a + b, 0) / leads.length : null,
    median: pick(0.5), p90: pick(0.9), max: leads.length ? leads[leads.length - 1] : null,
    onTimePct: withEta.length ? Math.round(onTime.length / withEta.length * 100) : null,
    lateN: lateOnes.length, avgLate: lateOnes.length ? lateOnes.reduce((a, x) => a + x.vsEta, 0) / lateOnes.length : null
  };
}

/** วันที่ควรได้ของ = วันออก PO + ระยะเวลาที่เคยส่งได้จริง (ยังไม่มีประวัติ → 14 วัน) */
function suggestEta(poAt) {
  const s = deliveryStats(deliveryItems());
  const days = Math.round(s.p90 != null ? s.p90 : (s.avg != null ? s.avg : 14));
  return { date: addDays(poAt || todayISO(), days), days, from: s.done ? `จากของ ${s.done} รายการที่รับมาแล้ว` : 'ค่าเริ่มต้น 14 วัน' };
}

function groupDelivery(list, keyOf, labelOf) {
  const g = new Map();
  for (const x of list) {
    const k = keyOf(x);
    if (!g.has(k)) g.set(k, { key: k, label: labelOf(x), items: [] });
    g.get(k).items.push(x);
  }
  return [...g.values()].map(x => ({ ...x, s: deliveryStats(x.items) }));
}

// ---- การ์ดคำขอ: ตาราง PR/PO + รับของ รายชิ้น (ติ๊กทีละรายการ หรือหลายรายการที่ใช้เลขเดียวกัน)

function deliveryPanel(r) {
  if (!isAdmin() || ![ST.APPROVED, ST.PO, ST.RECEIVED].includes(r.status)) return '';
  const items = (r.items || []).map((i, n) => ({ i, n })).filter(x => Number(x.i.qty) > 0);
  if (!items.length) return '';
  const withPo = items.filter(x => x.i.po_no).length;
  const got = items.filter(x => x.i.received_at).length;
  const late = items.filter(x => !x.i.received_at && daysBetween(x.i.eta_date, todayISO()) > 0).length;
  const sug = suggestEta(items.map(x => x.i.po_at).find(Boolean) || todayISO());
  return `<details class="dlv" ${r.status !== ST.RECEIVED ? 'open' : ''} data-dlv="${esc(r.request_id)}">
    <summary>PR/PO และการรับของ — มีเลข PO ${withPo}/${items.length} · รับแล้ว ${got}/${items.length}${late ? ` · <span class="over">เลยกำหนด ${late}</span>` : ''}</summary>
    <div class="table-wrap" style="max-height:none"><table class="grid dlv-tbl"><thead><tr>
      <th><input type="checkbox" data-dall aria-label="เลือกทุกรายการ"></th><th>สินค้า</th><th class="n">จำนวน</th><th class="n">รวม</th>
      <th>เลข PR</th><th>เลข PO</th><th>กำหนดส่ง</th><th>รับจริง</th><th>ระยะการส่ง</th>
    </tr></thead><tbody>${items.map(({ i, n }) => {
      const lead = daysBetween(i.po_at, i.received_at);
      const vsEta = daysBetween(i.eta_date, i.received_at);
      const lateBy = i.received_at ? null : daysBetween(i.eta_date, todayISO());
      return `<tr data-row="${n}" data-sku="${esc(i.sku)}" class="${i.received_at ? 'done' : lateBy > 0 ? 'is-late' : ''}">
        <td><input type="checkbox" data-drow aria-label="เลือก ${esc(i.name)}"></td>
        <td class="name">${esc(i.name)}<div class="pc">${esc(i.pc_code)} · ${esc(String(i.sku).startsWith('NEW-') ? 'รายการใหม่' : i.sku)}</div></td>
        <td class="n">${fmt(i.qty)} ${esc(i.unit)}</td><td class="n">${money(i.amount)}</td>
        <td class="small">${i.pr_no ? esc(i.pr_no) : '<span class="dim">–</span>'}</td>
        <td class="small">${i.po_no ? `<b>${esc(i.po_no)}</b>${i.po_at ? `<div class="pc">ออก ${esc(thaiDay(i.po_at))}</div>` : ''}` : '<span class="dim">ยังไม่ออก</span>'}</td>
        <td class="small">${i.eta_date ? esc(thaiDay(i.eta_date)) : '<span class="dim">–</span>'}</td>
        <td class="small">${i.received_at ? esc(thaiDay(i.received_at))
          : `<span class="${lateBy > 0 ? 'over' : 'dim'}">${lateBy > 0 ? `เลย ${lateBy} วัน` : lateBy != null ? `อีก ${-lateBy} วัน` : 'ยังไม่มา'}</span>`}</td>
        <td class="small">${lead != null ? `${lead} วัน ${vsEta == null ? '' : vsEta > 0 ? `<span class="over">ช้า ${vsEta} วัน</span>` : '<span class="ok-txt">ตรงเวลา</span>'}` : '<span class="dim">–</span>'}</td>
      </tr>`;
    }).join('')}</tbody></table></div>
    <div class="dlv-form">
      <label>เลข PR<input data-dpr placeholder="PR-…"></label>
      <label>เลข PO<input data-dpo placeholder="PO-…"></label>
      <label>กำหนดส่ง (ประมาณ)<input type="date" data-deta value="${esc(sug.date)}"></label>
      <button class="btn sm" data-dact="save">บันทึก PR/PO ให้รายการที่เลือก</button>
      <label>วันที่รับของ<input type="date" data-ddate value="${esc(todayISO())}"></label>
      <button class="btn sm ghost" data-dact="receive">รับของแล้ว (รายการที่เลือก)</button>
      <p class="hint">ติ๊กทีละรายการ หรือติ๊กหลายรายการที่ใช้เลข PR/PO เดียวกันแล้วกดบันทึกครั้งเดียว ·
        กำหนดส่งที่เติมให้ = วันออก PO + ${sug.days} วัน (${esc(sug.from)}) แก้เองได้ ·
        รายการที่เพิ่งได้เลข PO จะถูกตัดงบทันที · คำขอจะเป็น "รับของแล้ว" เมื่อของมาครบทุกรายการ</p>
    </div>
  </details>`;
}

document.addEventListener('change', e => {
  const all = e.target.closest('[data-dall]');
  if (!all) return;
  all.closest('table').querySelectorAll('[data-drow]').forEach(c => { c.checked = all.checked; });
});

document.addEventListener('click', async e => {
  const btn = e.target.closest('[data-dact]');
  if (!btn) return;
  const box = btn.closest('[data-dlv]');
  const id = box.dataset.dlv;
  const rows = [...box.querySelectorAll('tbody tr')].filter(tr => tr.querySelector('[data-drow]').checked)
    .map(tr => ({ row: Number(tr.dataset.row), sku: tr.dataset.sku }));
  if (!rows.length) return toast('ติ๊กเลือกรายการสินค้าก่อน');
  const v = sel => box.querySelector(sel).value.trim();
  let call, done;
  if (btn.dataset.dact === 'save') {
    const pr = v('[data-dpr]'), po = v('[data-dpo]'), eta = v('[data-deta]');
    if (!pr && !po && !eta) return toast('ใส่เลข PR หรือ PO หรือกำหนดส่งก่อน');
    if (po && !confirm(`ใส่เลข PO ${po} ให้ ${rows.length} รายการ?\nรายการที่ยังไม่มีเลข PO จะถูกตัดงบทันที`)) return;
    call = () => gasCall('savePrPo', id, { rows, pr, po, eta });
    done = `บันทึก PR/PO ให้ ${rows.length} รายการแล้ว`;
  } else {
    call = () => gasCall('receiveItems', id, { rows, date: v('[data-ddate]') || todayISO() });
    done = `บันทึกรับของ ${rows.length} รายการแล้ว`;
  }
  btn.disabled = true;
  try {
    requests = await call();
    updateReqBadge(); updateDlvBadge(); renderRequests(); refreshBudget();
    toast(done);
  } catch (err) { toast('ไม่สำเร็จ: ' + errMsg(err)); btn.disabled = false; }
});

// ---- แท็บ "ติดตามการส่งของ"

function updateDlvBadge() {
  const n = deliveryItems().filter(x => x.open && x.late).length;
  const b = $('#dlvBadge');
  b.hidden = !n;
  b.textContent = n;
}

function renderDelivery() {
  const all = deliveryItems();
  const s = deliveryStats(all);
  const d = n => n == null ? '–' : `${Math.round(n * 10) / 10} วัน`;
  $('#dlvCards').innerHTML = `
    <div class="card"><div class="k">ยังไม่ได้รับของ</div><div class="v">${fmt(s.open)}</div><div class="s">มูลค่า ${money(s.openAmt)}</div></div>
    <div class="card ${s.lateOpen ? 'dead' : ''}"><div class="k">เลยกำหนดส่ง</div><div class="v ${s.lateOpen ? 'over' : ''}">${fmt(s.lateOpen)}</div><div class="s">รายการที่ต้องตาม</div></div>
    <div class="card actual"><div class="k">รับของแล้ว</div><div class="v">${fmt(s.done)}</div><div class="s">จากที่ออก PR/PO ${fmt(s.n)} รายการ</div></div>
    <div class="card store"><div class="k">ระยะส่งเฉลี่ย</div><div class="v">${d(s.avg)}</div><div class="s">กลาง ${d(s.median)} · ช้าสุด ${d(s.max)}</div></div>
    <div class="card ${s.onTimePct != null && s.onTimePct < 80 ? 'dead' : 'actual'}"><div class="k">ส่งตรงเวลา</div><div class="v">${s.onTimePct == null ? '–' : s.onTimePct + '%'}</div><div class="s">${s.lateN ? `ช้า ${s.lateN} รายการ เฉลี่ย ${d(s.avgLate)}` : 'ยังไม่มีของที่มาช้า'}</div></div>`;

  const byPc = groupDelivery(all.filter(x => x.lead != null), x => x.i.pc_code || '—', x => `${x.i.pc_code || '—'} ${x.i.pc_name || ''}`)
    .filter(g => g.s.done).sort((a, b) => (b.s.avg || 0) - (a.s.avg || 0));
  const byPo = groupDelivery(all, x => x.i.po_no || '—', x => x.i.po_no || 'ยังไม่มีเลข PO');
  const worst = all.filter(x => x.open && x.late).sort((a, b) => b.lateBy - a.lateBy).slice(0, 5);
  $('#dlvAnalysis').innerHTML = !s.n ? '<p class="hint">ยังไม่มีของที่ออก PR/PO — ใส่เลข PR/PO ได้ที่การ์ดคำขอที่อนุมัติแล้ว (แท็บตรวจคำขอสั่งซื้อ)</p>' : `
    <div class="dlv-an">
      <div class="an-box"><h4>ควรเผื่อเวลาสั่งของ</h4>
        <p class="an-big">${d(s.p90)}</p>
        <p class="hint">${s.done ? `9 ใน 10 ครั้ง ของมาภายในนี้ (เฉลี่ยจริง ${d(s.avg)}) — ตั้งกำหนดส่งเท่านี้จะพลาดน้อยที่สุด`
          : 'ยังไม่มีของที่รับแล้ว ระบบใช้ค่าเริ่มต้น 14 วันไปก่อน'}</p></div>
      <div class="an-box"><h4>PC ที่ของมาช้าที่สุด</h4>
        ${byPc.length ? `<ol class="an-list">${byPc.slice(0, 5).map(g => `<li><span>${esc(g.label)}</span>
          <span class="num">${d(g.s.avg)}${g.s.onTimePct != null ? ` · ตรงเวลา ${g.s.onTimePct}%` : ''}</span></li>`).join('')}</ol>`
          : '<p class="hint">ยังไม่มีข้อมูลพอ</p>'}</div>
      <div class="an-box"><h4>ต้องตามตอนนี้</h4>
        ${worst.length ? `<ol class="an-list">${worst.map(x => `<li><span>${esc(x.i.name)}
          <span class="pc">${esc(x.i.po_no || '—')} · ${esc(x.r.request_id)}</span></span>
          <span class="num over">เลย ${x.lateBy} วัน</span></li>`).join('')}</ol>`
          : '<p class="hint">ไม่มีของที่เลยกำหนด</p>'}</div>
      <div class="an-box"><h4>แยกตามเลข PR/PO</h4>
        <ol class="an-list">${byPo.slice(0, 6).map(g => `<li><span>${esc(g.label)}
          <span class="pc">${fmt(g.items.length)} รายการ · ${money(g.items.reduce((a, x) => a + x.amount, 0))}</span></span>
          <span class="num">${g.s.open ? `รอ ${g.s.open}` : 'ครบแล้ว'}</span></li>`).join('')}</ol></div>
    </div>`;

  const scope = $('#dlvScope').value;
  const list = all.filter(x => scope === 'all' || (scope === 'open' && x.open) || (scope === 'late' && x.open && x.late) || (scope === 'done' && !x.open))
    .sort((a, b) => (b.late ? 1 : 0) - (a.late ? 1 : 0) || String(a.i.eta_date).localeCompare(String(b.i.eta_date)));
  $('#dlvListTitle').textContent = `รายการ (${fmt(list.length)})`;
  $('#tblDelivery').innerHTML = !list.length ? '<tbody><tr><td class="empty">ไม่มีรายการตามเงื่อนไข</td></tr></tbody>' : `
    <thead><tr><th>คำขอ</th><th>สินค้า</th><th>PC</th><th class="n">จำนวน</th><th class="n">รวม</th>
      <th>PR</th><th>PO</th><th>ออก PO</th><th>กำหนดส่ง</th><th>รับจริง</th><th class="n">ระยะ</th><th>สถานะ</th></tr></thead>
    <tbody>${list.map(x => `<tr class="${x.open && x.late ? 'is-late' : ''}">
      <td class="sku"><button class="link-btn" data-open-req="${esc(x.r.request_id)}">${esc(x.r.request_id)}</button></td>
      <td class="name">${esc(x.i.name)}</td><td class="pc">${esc(x.i.pc_code)}</td>
      <td class="n">${fmt(x.i.qty)}</td><td class="n">${money(x.amount)}</td>
      <td class="small">${esc(x.i.pr_no || '–')}</td><td class="small">${esc(x.i.po_no || '–')}</td>
      <td class="small">${esc(thaiDay(x.i.po_at)) || '–'}</td><td class="small">${esc(thaiDay(x.i.eta_date)) || '–'}</td>
      <td class="small">${esc(thaiDay(x.i.received_at)) || '–'}</td>
      <td class="n">${x.lead != null ? x.lead + ' วัน' : '–'}</td>
      <td>${x.open ? (x.late ? `<span class="tag st-rejected">เลย ${x.lateBy} วัน</span>` : '<span class="tag st-pending">รอของ</span>')
        : x.vsEta > 0 ? `<span class="tag st-edit">ช้า ${x.vsEta} วัน</span>` : '<span class="tag st-approved">ตรงเวลา</span>'}</td>
    </tr>`).join('')}</tbody>`;
}

$('#dlvScope').addEventListener('change', renderDelivery);
$('#btnDlvExport').addEventListener('click', () => {
  const all = deliveryItems();
  if (!all.length) return toast('ยังไม่มีของที่ออก PR/PO');
  const rows = all.map(x => ({ 'คำขอ': x.r.request_id, 'สินค้า': x.i.name, 'รหัส': x.i.sku, 'PC': x.i.pc_code,
    'จำนวน': x.i.qty, 'หน่วย': x.i.unit, 'มูลค่า': x.amount, 'PR': x.i.pr_no, 'PO': x.i.po_no,
    'วันออก PO': x.i.po_at, 'กำหนดส่ง': x.i.eta_date, 'รับจริง': x.i.received_at,
    'ระยะการส่ง (วัน)': x.lead == null ? '' : x.lead, 'ช้ากว่ากำหนด (วัน)': x.vsEta == null ? '' : x.vsEta,
    'สถานะ': x.open ? (x.late ? 'เลยกำหนด' : 'รอของ') : (x.vsEta > 0 ? 'ช้า' : 'ตรงเวลา') }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'การส่งของ');
  XLSX.writeFile(wb, `ติดตามการส่งของ_${stamp()}.xlsx`);
});

// ------------------------------------------------------------ boot

// Apps Script: never show a stale per-browser copy of the shared slots — wait for the server's
if (SERVER) for (const t of SHARED_SLOTS) state.files[t] = [];
recompute();
if (SERVER) {
  if (token) {
    applySession({ status: 'loading' });
    gasCall('getSessionInfo').then(applySession, () => { setToken(''); applySession(null); });
  } else {
    applySession(null); // shows the sign-in form
  }
}
})();
