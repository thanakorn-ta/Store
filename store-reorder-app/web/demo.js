// Demo / trial mode switch — runs before app.js.
//   ?demo=1 in the address bar = โหมดทดลอง สำหรับหน้านั้นเท่านั้น
// ปิดหน้าหรือเปิดลิงก์ปกติ = กลับมาใช้ระบบจริงทันที (ไม่จำค่าไว้ในเบราว์เซอร์อีกแล้ว —
// ก่อนหน้านี้เคยจำไว้ ทำให้เผลอดูข้อมูลตัวอย่างค้างอยู่โดยไม่รู้ตัว)
// เมื่อเปิดอยู่: demo/mock.js + demo/server.js (โค้ด src/*.js ตัวจริง) + demo/seed.js
// จะรันทั้งระบบในหน้านี้ด้วยข้อมูลตัวอย่าง ไม่แตะ Google Sheet / Drive / อีเมลจริง
// หน้าเว็บที่เสิร์ฟจาก Apps Script เป็นระบบจริงเสมอ
(function () {
  if (typeof google !== 'undefined' && google.script && google.script.run) return;
  var OLD_KEY = 'store-reorder-ai:demo';
  try { localStorage.removeItem(OLD_KEY); } catch (e) { /* private mode */ }
  if (new URLSearchParams(location.search).get('demo') !== '1') return;
  // The tag names are split on purpose: this file is also inlined into the Apps Script page,
  // where a literal opening script tag inside a script block breaks the rest of the page.
  var open = '<scr' + 'ipt src="', close = '"></scr' + 'ipt>';
  document.write(open + 'demo/mock.js' + close + open + 'demo/server.js' + close + open + 'demo/seed.js' + close);
})();
