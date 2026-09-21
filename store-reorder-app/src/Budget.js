/**
 * Budget.js
 * Step 8/8b/9: once a PC owner confirms quantities, check spend-to-date
 * against the annual budget before letting Admin issue a PR/PO.
 *
 * BLOCKED on an accounting decision (see analysis 2569-09-18, "เรื่องที่ต้อง
 * ตัดสินใจก่อนเขียนโค้ด"): Budget.xlsx is keyed by Company x Media Location x
 * GL Code x Month, but most Store PCs (สายไฟ, อุปกรณ์ไฟฟ้า, PM, เครื่องมือช่าง,
 * เบิกประจำวัน, ...) don't map 1:1 to a single Media Location. Needs answers on:
 *   1. Which Location/cost-center absorbs each "shared" PC's spend
 *      (user-selected project? pro-rated? one general pool?)
 *   2. Which GL code applies: 530090 (ค่าวัสดุสิ้นเปลืองใช้ไป) vs
 *      530050 (ค่าซ่อมแซมบำรุงรักษาป้ายโฆษณา) vs something PC-specific
 *   3. Check against the monthly figure or the remaining annual balance
 *      (Budget.xlsx only has Revise Budget for Jul-Dec)
 *   4. Where committed-but-not-yet-actual spend is tracked (this app should
 *      write every issued PR/PO into CONFIG.SHEETS.PR_LOG and treat that as
 *      the source of truth for "used so far", since Budget.xlsx has no actuals)
 *
 * Until master_pc has a validated location/gl mapping, checkBudget_() below
 * is a placeholder that always returns "unknown" so nothing gets silently
 * auto-approved against a guessed number.
 */

function checkBudget_(pcCode, amount, monthNumber) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var masterPc = readSheetAsObjects_(ss, CONFIG.SHEETS.MASTER_PC);
  var pc = masterPc.filter(function (p) { return p.pc_code === pcCode; })[0];

  if (!pc || !pc.budget_location || !pc.budget_gl_code) {
    return { status: 'unknown', reason: 'ยังไม่มีการ mapping งบให้ ' + pcCode + ' ใน master_pc' };
  }

  var budget = readSheetAsObjects_(ss, CONFIG.SHEETS.BUDGET);
  var remaining = getBudgetRemaining_(budget, pc.budget_location, pc.budget_gl_code, monthNumber, pcCode);

  if (remaining === null) {
    return { status: 'unknown', reason: 'ไม่พบงบสำหรับ ' + pc.budget_location + ' / ' + pc.budget_gl_code };
  }
  return remaining >= amount
    ? { status: 'ok', remaining: remaining }
    : { status: 'exceeded', remaining: remaining, shortfall: amount - remaining };
}

/**
 * remaining = revise_budget(location, gl, month) - sum(committed PR/PO so far
 * this month/location/gl from CONFIG.SHEETS.PR_LOG).
 * Budget.xlsx column layout (see raw import): Company, Division, Media Type,
 * Media Group, Media Location, Sub Topic, Expense Group, GL Code, ..., Month,
 * Budget, Revise Budget, ...
 */
function getBudgetRemaining_(budgetRows, location, glCode, monthNumber, pcCode) {
  var monthRow = budgetRows.filter(function (b) {
    return b.gl_code == glCode && b.media_location == location && Number(b.month_number) === Number(monthNumber);
  })[0];
  if (!monthRow) return null;

  var budgeted = Number(monthRow.revise_budget) || Number(monthRow.budget) || 0;

  var prLog = readSheetAsObjects_(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID), CONFIG.SHEETS.PR_LOG);
  var committed = prLog
    .filter(function (r) {
      return r.pc_code === pcCode && Number(r.month_number) === Number(monthNumber);
    })
    .reduce(function (sum, r) { return sum + (Number(r.amount) || 0); }, 0);

  return budgeted - committed;
}
