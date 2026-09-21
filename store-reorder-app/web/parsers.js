/**
 * parsers.js
 * Reads the 4 Store export formats (as SheetJS row arrays) into plain records.
 * Column positions are found from each file's header row, not hard-coded,
 * so a shifted column in a future export still parses.
 *
 *   minmax  : MIN_MAX_Calculated.xlsx          (hand-entered Avg/MIN/MAX)
 *   usage   : รวมการใช้ของ.xlsx / การใช้งานเดือน*.xls (monthly withdrawals)
 *   balance : ยอดคงเหลือสินค้า.xls               (stock on hand)
 *   reorder : รายงานสินค้าถึงจุดสั่งซื้อ.xls       (Store's reorder alert)
 */

const SKU_RE = /^P\d{5,}$/;
const PC_RE = /PC\s*(\d{3,})\s*(.*)$/;

const THAI_MONTHS = [
  ['มกราคม', 'ม.ค.'], ['กุมภาพันธ์', 'ก.พ.'], ['มีนาคม', 'มี.ค.'], ['เมษายน', 'เม.ย.'],
  ['พฤษภาคม', 'พ.ค.'], ['มิถุนายน', 'มิ.ย.'], ['กรกฎาคม', 'ก.ค.'], ['สิงหาคม', 'ส.ค.'],
  ['กันยายน', 'ก.ย.'], ['ตุลาคม', 'ต.ค.'], ['พฤศจิกายน', 'พ.ย.'], ['ธันวาคม', 'ธ.ค.']
];
const MONTH_SHORT = THAI_MONTHS.map(m => m[1]);

function num(v) {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  const s = String(v).replace(/[^\d.\-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

const str = v => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim());

/** Returns month number 1-12 found in text (full or abbreviated Thai name), or 0. */
function findMonth(text) {
  const t = String(text || '').replace(/\s+/g, '');
  // Full names first: "มีนาคม" must not be matched by abbreviation "มี.ค." checks etc.
  for (let i = 0; i < 12; i++) if (t.includes(THAI_MONTHS[i][0])) return i + 1;
  // Abbreviations: longest first so "มี.ค." wins over "ม.ค." (which is a substring of it)
  const order = [...Array(12).keys()].sort((a, b) => THAI_MONTHS[b][1].length - THAI_MONTHS[a][1].length);
  for (const i of order) if (t.includes(THAI_MONTHS[i][1])) return i + 1;
  return 0;
}

function findYear(text) {
  const m = String(text || '').match(/(25\d\d|20\d\d)/);
  return m ? Number(m[1]) : 0;
}

/** Locate the header row: first row containing every keyword. */
function findHeader(rows, keywords, limit = 60) {
  for (let r = 0; r < Math.min(rows.length, limit); r++) {
    const cells = rows[r].map(str);
    if (keywords.every(k => cells.some(c => c.includes(k)))) return r;
  }
  return -1;
}

/** Map of logical column -> index, using the first header cell that matches. */
function mapColumns(headerRow, spec) {
  const cells = headerRow.map(str);
  const out = {};
  for (const [key, tests] of Object.entries(spec)) {
    out[key] = cells.findIndex(c => tests.some(t => (t instanceof RegExp ? t.test(c) : c === t)));
  }
  return out;
}

function skuIn(row, preferIdx) {
  if (preferIdx >= 0 && SKU_RE.test(str(row[preferIdx]))) return { sku: str(row[preferIdx]), idx: preferIdx };
  for (let c = 0; c < row.length; c++) if (SKU_RE.test(str(row[c]))) return { sku: str(row[c]), idx: c };
  return null;
}

function pcHeader(row) {
  const joined = row.map(str).filter(Boolean).join(' ');
  const m = joined.match(PC_RE);
  return m ? { pc: 'PC' + m[1], pcName: m[2].trim() } : null;
}

/** Single non-empty text cell and no SKU -> a group label row (e.g. "Cookies"). */
function groupLabel(row, skipIdx = -1) {
  const cells = row.map((c, i) => (i === skipIdx ? '' : str(c))).filter(Boolean);
  return cells.length === 1 && !/^[\d.,\s-]+$/.test(cells[0]) ? cells[0] : null;
}

// ---------------------------------------------------------------- detect

function detectType(rows, fileName = '') {
  const head = rows.slice(0, 40).map(r => r.map(str).join('|')).join('\n');
  if (/Media Location/.test(head) && /GL Code/.test(head)) return 'budget';
  if (/ต้องซื้ออย่างน้อย|จุดต่ำสุด|ถึงจุดสั่งซื้อ/.test(head)) return 'reorder';
  if (/Avg\/month/i.test(head) || (/\bMIN\b/.test(head) && /\bMAX\b/.test(head))) return 'minmax';
  if (/จำนวนใช้|รายงานการใช้/.test(head) || /การใช้งานเดือน|รวมการใช้/.test(fileName)) return 'usage';
  if (/ยอดคงเหลือ|มูลค่าคงเหลือ/.test(head)) return 'balance';
  if (/รายการที่เลือกสั่งซื้อ|Merged/.test(head)) return 'merged';
  return null;
}

// ---------------------------------------------------------------- minmax

function parseMinMax(rows) {
  const h = findHeader(rows, ['SKU']);
  if (h < 0) throw new Error('ไม่พบหัวตาราง SKU ในไฟล์ MIN/MAX');
  const col = mapColumns(rows[h], {
    sku: ['SKU'], name: [/ชื่อสินค้า/], avg: [/Avg/i], unit: ['หน่วย'],
    lead: [/lead\s*time/i], safety: [/safety/i], min: ['MIN'], max: ['MAX'], cat: [/ประเภท|หมวด/]
  });
  const out = [];
  let group = '';
  for (let r = h + 1; r < rows.length; r++) {
    const row = rows[r];
    const hit = skuIn(row, col.sku);
    if (!hit) { const g = groupLabel(row); if (g) group = g; continue; }
    out.push({
      sku: hit.sku, name: str(row[col.name]), group,
      avg: num(row[col.avg]), unit: str(row[col.unit]),
      lead: num(row[col.lead]) || 0, safety: num(row[col.safety]),
      min: str(row[col.min]) === '' ? null : num(row[col.min]),
      max: str(row[col.max]) === '' ? null : num(row[col.max]),
      category: col.cat >= 0 ? str(row[col.cat]) : ''
    });
  }
  return out;
}

// ---------------------------------------------------------------- usage

/**
 * Handles both the Power Query combined file (first column = Source.Name
 * "การใช้งานเดือน ก.ค.xls") and a raw single-month export (month taken from
 * the "ประจำเดือน กรกฎาคม 2569" title row or the file name).
 */
function parseUsage(rows, fileName = '') {
  const out = [];
  let group = '', title = '';
  let col = null;
  const fileMonth = findMonth(fileName);

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const cells = row.map(str);
    const line = cells.join(' ');

    if (/ประจำเดือน|รายงานการใช้/.test(line)) { title = line; continue; }
    if (cells.some(c => c === 'รหัสสินค้า') && cells.some(c => c.includes('จำนวนใช้'))) {
      col = mapColumns(row, {
        sku: ['รหัสสินค้า'], name: ['ชื่อสินค้า', 'สินค้า'], qty: [/จำนวนใช้/], unit: ['หน่วย'],
        price: ['ราคา'], total: [/รวมเป็นเงิน/]
      });
      continue;
    }
    const hit = skuIn(row, col ? col.sku : -1);
    // Source.Name column in the combined file is column 0 and holds the file name
    const srcName = /การใช้งานเดือน|\.xls/.test(cells[0]) ? cells[0] : '';
    if (!hit) {
      const g = groupLabel(row, srcName ? 0 : -1);
      if (g && !/ประจำเดือน/.test(g)) group = g;
      continue;
    }
    const i = hit.idx;
    const c = col || { name: i + 1, qty: i + 4, unit: i + 5, price: i + 6, total: i + 7 };
    const month = findMonth(srcName) || findMonth(title) || fileMonth;
    out.push({
      sku: hit.sku, name: str(row[c.name]), group,
      month, year: findYear(title),
      qty: num(row[c.qty]), unit: str(row[c.unit]),
      price: num(row[c.price]), total: num(row[c.total])
    });
  }
  return out;
}

// ---------------------------------------------------------------- balance

function parseBalance(rows) {
  const out = [];
  let pc = '', pcName = '', col = null;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const pch = pcHeader(row);
    if (pch && !skuIn(row, -1)) { pc = pch.pc; pcName = pch.pcName; continue; }
    const cells = row.map(str);
    if (cells.includes('รหัสสินค้า')) {
      col = mapColumns(row, {
        sku: ['รหัสสินค้า'], name: ['สินค้า', 'ชื่อสินค้า'], loc: [/ตำแหน่ง|ต่ำแหน่ง/],
        bal: ['คงเหลือ'], unit: ['หน่วย'], cost: [/ต้นทุน/], value: [/มูลค่า/]
      });
      continue;
    }
    const hit = skuIn(row, col ? col.sku : -1);
    if (!hit) continue;
    const i = hit.idx;
    const c = col || { name: i + 2, loc: i + 5, bal: i + 6, unit: i + 8, cost: i + 9, value: i + 11 };
    out.push({
      sku: hit.sku, name: str(row[c.name]), pc, pcName, location: str(row[c.loc]),
      balance: num(row[c.bal]), unit: str(row[c.unit]),
      cost: num(row[c.cost]), value: num(row[c.value])
    });
  }
  return out;
}

// ---------------------------------------------------------------- reorder

function parseReorder(rows) {
  const out = [];
  let pc = '', pcName = '', col = null;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const pch = pcHeader(row);
    if (pch && !skuIn(row, -1)) { pc = pch.pc; pcName = pch.pcName; continue; }
    const cells = row.map(str);
    if (cells.includes('รหัสสินค้า')) {
      col = mapColumns(row, {
        sku: ['รหัสสินค้า'], name: ['สินค้า', 'ชื่อสินค้า'], bal: [/คงเหลือ/], min: [/จุดต่ำสุด/],
        need: [/ต้องซื้อ/], unit: ['หน่วย'], lot: [/สั่งซื้อแต่ละครั้ง/]
      });
      continue;
    }
    const hit = skuIn(row, col ? col.sku : -1);
    if (!hit || !col) continue;
    out.push({
      sku: hit.sku, name: str(row[col.name]), pc, pcName,
      balance: num(row[col.bal]), min: num(row[col.min]), need: num(row[col.need]),
      unit: str(row[col.unit]), lot: str(row[col.lot])
    });
  }
  return out;
}

// ---------------------------------------------------------------- merged (baseline)

/** Reads the "รวมข้อมูล" sheet of a workbook this app exported earlier. */
function parseMergedExport(rows) {
  const h = findHeader(rows, ['รหัสสินค้า', 'คงเหลือ']);
  if (h < 0) return [];
  const head = rows[h].map(str);
  const idx = k => head.indexOf(k);
  const out = [];
  for (let r = h + 1; r < rows.length; r++) {
    const sku = str(rows[r][idx('รหัสสินค้า')]);
    if (!SKU_RE.test(sku)) continue;
    out.push({
      sku,
      balance: num(rows[r][idx('คงเหลือ')]),
      orderQty: num(rows[r][idx('แนะนำสั่ง (ใช้จริง)')]),
      avgActual: num(rows[r][idx('เฉลี่ยใช้จริง/เดือน')])
    });
  }
  return out;
}

// ---------------------------------------------------------------- budget

/**
 * "Budget STT 2026 - Revise-Budget" export: one row per Calculation x Month.
 * Aggregated here to one row per (Company x Media Location x GL Code) x month
 * so the Apps Script side stores a compact budget_master.
 *   plan = Revise Budget when filled, else Budget (decided per source row)
 * Field names match BUDGET_HEADER in src/Requests.js.
 */
function parseBudget(rows) {
  const h = findHeader(rows, ['Media Location', 'GL Code']);
  if (h < 0) throw new Error('ไม่พบหัวตาราง Media Location / GL Code ในไฟล์ Budget');
  const col = mapColumns(rows[h], {
    calc: ['Calculation'], company: ['Company'], division: ['Division'], mediaType: ['Media Type'], mediaGroup: ['Media Group'],
    location: ['Media Location'], expenseGroup: ['Expense Group'], gl: ['GL Code'],
    glName: [/^ประเภทของค่าใช้จ่าย/], year: ['Year'], month: ['Month Number'],
    budget: ['Budget'], revise: ['Revise Budget'], actual: ['Actual']
  });
  const lines = new Map();
  let skipped = 0;
  for (let r = h + 1; r < rows.length; r++) {
    const row = rows[r];
    const gl = str(row[col.gl]), month = num(row[col.month]);
    const hasData = num(row[col.budget]) || num(row[col.actual]) || num(row[col.revise]);
    // Media Location is blank on ~25% of rows; the Calculation column always carries it:
    // "Electricity / Cookies - Operation Costs" -> "Cookies"
    const calcLoc = (str(row[col.calc]).match(/\/\s*(.+?)\s+-\s+[^-]+$/) || [])[1] || '';
    const location = str(row[col.location]) || calcLoc || str(row[col.mediaGroup]);
    if (!gl || !location || !(month >= 1 && month <= 12)) { if (hasData) skipped++; continue; }
    const company = str(row[col.company]);
    const key = `${company}|${location}|${gl}`;
    const k = key + '|' + month;
    if (!lines.has(k)) {
      lines.set(k, {
        key, company, division: str(row[col.division]), media_type: str(row[col.mediaType]),
        media_group: str(row[col.mediaGroup]), media_location: location,
        expense_group: str(row[col.expenseGroup]), gl_code: gl, gl_name: str(row[col.glName]),
        year: num(row[col.year]), month_number: month, budget: 0, revise_budget: 0, actual: 0, plan: 0
      });
    }
    const l = lines.get(k);
    const budget = num(row[col.budget]);
    const reviseCell = str(row[col.revise]);
    const revise = num(row[col.revise]);
    l.budget += budget;
    l.revise_budget += revise;
    l.actual += num(row[col.actual]);
    l.plan += reviseCell !== '' ? revise : budget;
  }
  const round = n => Math.round(n * 100) / 100;
  const out = [...lines.values()].map(l => ({
    ...l, budget: round(l.budget), revise_budget: round(l.revise_budget), actual: round(l.actual), plan: round(l.plan)
  }));
  out.skipped = skipped; // rows with amounts but no GL Code / month — reported to the admin
  return out;
}
