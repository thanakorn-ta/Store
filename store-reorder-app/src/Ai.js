/**
 * Ai.js
 * Optional layer on top of the deterministic steps (Import/MinMax/Reorder/
 * Budget). Claude never decides quantities or budget approval — it only
 * drafts human-facing text from numbers Apps Script already computed:
 *   1. Friendlier per-PC reorder emails (replaces the plain template in
 *      Notify.js once MIN/MAX has been validated for a few cycles)
 *   2. A summary of a PC owner's confirmation/notes for Admin
 *   3. A budget-exceeded approval request draft for the PC's manager
 *
 * Setup:
 *   Project Settings > Script Properties > add ANTHROPIC_API_KEY
 *   (never commit the key; keep it out of this repo entirely)
 */

var CLAUDE_MODEL = 'claude-sonnet-5';
var CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';

function callClaude_(systemPrompt, userPrompt, maxTokens) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ยังไม่ได้ตั้งค่า ANTHROPIC_API_KEY ใน Script Properties');

  var res = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: maxTokens || 600,
      system: systemPrompt,
      messages: [{ role: 'user', content: userPrompt }]
    }),
    muteHttpExceptions: true
  });

  var body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    throw new Error('Claude API error: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  if (body.stop_reason === 'refusal') throw new Error('Claude ปฏิเสธคำขอนี้');
  // content may start with a thinking block (adaptive thinking is on by default) — join text blocks only
  return body.content
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('\n');
}

function draftReorderEmail_(pcName, items) {
  var system = 'คุณคือผู้ช่วยฝ่ายจัดซื้อของ Plan B Media เขียนอีเมลภาษาไทยสั้น สุภาพ ' +
    'แจ้งรายการที่ต้องสั่งซื้อ ระบุจำนวนที่ระบบคำนวณให้ตรงตัวเลขที่ให้มาเท่านั้น ' +
    'ห้ามเปลี่ยนตัวเลข ห้ามเดาข้อมูลเพิ่ม ปิดท้ายด้วยลิงก์ยืนยัน';
  var user = 'ทีม/PC: ' + pcName + '\nรายการ:\n' + items.map(function (i) {
    return '- ' + i.name + ' คงเหลือ ' + i.balance + ' แนะนำสั่งเพิ่ม ' + i.suggested_qty;
  }).join('\n') + '\n\nลิงก์ยืนยัน: {{CONFIRM_LINK}}';
  return callClaude_(system, user, 500);
}

function summarizeConfirmationForAdmin_(pcName, confirmedItems, notes) {
  var system = 'สรุปคำตอบของทีม Store เป็นภาษาไทย กระชับ สำหรับฝ่ายจัดซื้อ ' +
    'ระบุเฉพาะรายการที่ยืนยันซื้อจริง พร้อมจำนวนและมูลค่ารวม ห้ามคำนวณตัวเลขใหม่เอง';
  var user = 'PC: ' + pcName + '\nรายการที่ยืนยัน:\n' + JSON.stringify(confirmedItems) +
    '\nหมายเหตุจากทีม: ' + (notes || '-');
  return callClaude_(system, user, 400);
}

function draftBudgetExceptionRequest_(pcName, items, shortfall) {
  var system = 'ร่างอีเมลขออนุมัติซื้อของเกินงบ ภาษาไทย สุภาพ ส่งถึงหัวหน้า PC ' +
    'ระบุยอดที่เกินงบตามตัวเลขที่ให้มา ห้ามแต่งเหตุผลทางธุรกิจเอง ให้เว้นที่ให้ผู้ส่งเติมเหตุผลเอง';
  var user = 'PC: ' + pcName + '\nยอดที่เกินงบ: ' + shortfall + ' บาท\nรายการ:\n' +
    items.map(function (i) { return '- ' + i.name + ' x' + i.confirmed_qty; }).join('\n');
  return callClaude_(system, user, 400);
}
