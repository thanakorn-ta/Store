/**
 * MinMax.js
 * Step 3-4 of the workflow: recompute MIN/MAX from actual usage
 * (raw_usage, last 9 months) instead of the hand-entered "Avg/month"
 * that turned out to be wrong (111 of 113 SKUs over-estimated >1.5x,
 * see analysis 2569-09-18).
 *
 * Classification per SKU:
 *   - "เบิกประจำ"   : used in >= CONFIG.REGULAR_ITEM_MIN_MONTHS of the last 9
 *                     months -> safe to auto-calc MIN/MAX and auto-flag reorder.
 *   - "เบิกตามงาน"  : used in fewer months -> demand is project-driven, do NOT
 *                     set a fixed reorder point; surface history and let the
 *                     PC owner decide each time (see Notify.js).
 *   - "Dead stock"  : balance > 0 but zero usage in 9 months -> flag for
 *                     Store to review (write-off / redistribute / stop buying),
 *                     never auto-included in a reorder email.
 */

function recomputeMinMax() {
  ownerOnly_();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var usage = readSheetAsObjects_(ss, CONFIG.SHEETS.RAW_USAGE);
  var balance = readSheetAsObjects_(ss, CONFIG.SHEETS.RAW_BALANCE);
  var masterItem = readSheetAsObjects_(ss, CONFIG.SHEETS.MASTER_ITEM);

  var usageBySku = groupUsageBySku_(usage);
  var balBySku = {};
  balance.forEach(function (b) { balBySku[b.sku] = b; });
  var itemBySku = {};
  masterItem.forEach(function (m) { itemBySku[m.sku] = m; });

  var out = [];
  var allSkus = uniqueKeys_([usageBySku, balBySku, itemBySku]);

  allSkus.forEach(function (sku) {
    var u = usageBySku[sku] || { total: 0, monthsUsed: 0, byMonth: {} };
    var bal = balBySku[sku] || { balance: 0, unit_cost: 0, value: 0 };
    var item = itemBySku[sku] || {};

    var avgPerMonth = u.total / 9;
    var leadTime = Number(item.lead_time_days) || CONFIG.DEFAULT_LEAD_TIME_DAYS;
    var avgPerDay = avgPerMonth / 30;

    var category, safety = '', min = '', max = '', status = '';

    if (bal.balance > 0 && u.total <= 0) {
      category = 'Dead stock';
      status = 'ตรวจสอบ: ไม่ถูกเบิกใน 9 เดือน';
    } else if (u.total <= 0) {
      return; // no balance, no usage -> not interesting, skip
    } else if (u.monthsUsed >= CONFIG.REGULAR_ITEM_MIN_MONTHS) {
      category = 'เบิกประจำ';
      safety = Math.ceil(CONFIG.MINMAX_FACTORS.SAFETY * avgPerDay * leadTime);
      min = Math.ceil(CONFIG.MINMAX_FACTORS.MIN * avgPerDay * leadTime);
      max = Math.ceil(CONFIG.MINMAX_FACTORS.MAX * avgPerDay * leadTime);
      status = bal.balance <= min ? 'ต้องสั่งซื้อ' : (bal.balance <= max ? 'ปกติ' : 'สต๊อกเกิน MAX');
    } else {
      category = 'เบิกตามงาน';
      status = 'ถามทีมก่อนสั่งทุกครั้ง (ไม่ตั้งจุดสั่งซื้อตายตัว)';
    }

    out.push([
      sku, u.name || bal.name || item.name || '', category,
      leadTime, u.monthsUsed, round1_(avgPerMonth),
      safety, min, max,
      bal.balance || 0, bal.unit_cost || 0, bal.value || 0,
      status
    ]);
  });

  var ssOut = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ssOut.getSheetByName(CONFIG.SHEETS.MINMAX) || ssOut.insertSheet(CONFIG.SHEETS.MINMAX);
  sheet.clearContents();
  var header = ['sku', 'name', 'category', 'lead_time_days', 'months_used_9',
    'avg_per_month', 'safety_stock', 'min', 'max', 'balance', 'unit_cost', 'value', 'status'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (out.length) sheet.getRange(2, 1, out.length, header.length).setValues(out);

  logActivity_('recomputeMinMax', 'ok', out.length + ' skus processed');
  return out;
}

function groupUsageBySku_(usageRows) {
  var bySku = {};
  usageRows.forEach(function (r) {
    var sku = r.sku;
    if (!bySku[sku]) bySku[sku] = { name: r.name, total: 0, monthsUsed: 0, byMonth: {} };
    var qty = Number(r.qty) || 0;
    var m = Number(r.month);
    if (!bySku[sku].byMonth[m]) bySku[sku].byMonth[m] = 0;
    bySku[sku].byMonth[m] += qty;
  });
  Object.keys(bySku).forEach(function (sku) {
    var months = bySku[sku].byMonth;
    var total = 0, monthsUsed = 0;
    Object.keys(months).forEach(function (m) {
      total += months[m];
      if (months[m] > 0) monthsUsed++;
    });
    bySku[sku].total = total;
    bySku[sku].monthsUsed = monthsUsed;
  });
  return bySku;
}

function uniqueKeys_(objs) {
  var seen = {};
  objs.forEach(function (o) { Object.keys(o).forEach(function (k) { seen[k] = true; }); });
  return Object.keys(seen);
}

function round1_(n) { return Math.round(n * 10) / 10; }

function readSheetAsObjects_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var header = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) obj[header[c]] = values[i][c];
    rows.push(obj);
  }
  return rows;
}
