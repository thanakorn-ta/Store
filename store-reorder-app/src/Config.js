/**
 * Config.js
 * Central place for spreadsheet IDs, sheet names, and tunable constants.
 * Fill SPREADSHEET_ID after creating the Google Sheet that will hold all
 * imported data + master tables (see README "ระยะ 0/1").
 */

var CONFIG = {
  // ID of the Google Sheet used as the database for this whole workflow.
  // https://docs.google.com/spreadsheets/d/1V91agWVoocDbJ54KLxAN27btFaYvEewR10NHsA3TVSY
  SPREADSHEET_ID: '1V91agWVoocDbJ54KLxAN27btFaYvEewR10NHsA3TVSY',

  SHEETS: {
    RAW_BALANCE: 'raw_balance',       // สำเนายอดคงเหลือสินค้า (ล่าสุด)
    RAW_USAGE: 'raw_usage',           // สำเนารวมการใช้ของ (สะสมทุกเดือน)
    RAW_REORDER: 'raw_reorder',       // สำเนารายงานสินค้าถึงจุดสั่งซื้อ (ล่าสุด, ใช้ตรวจทานเท่านั้น)
    MASTER_PC: 'master_pc',           // รหัส PC, ชื่อ, อีเมล user, อีเมลหัวหน้า, budget mapping
    MASTER_ITEM: 'master_item',       // SKU, lead time, ประเภท(ประจำ/ตามงาน), หน่วยสั่งซื้อ
    MINMAX: 'minmax_calculated',      // ผลลัพธ์คำนวณ MIN/MAX ล่าสุด (เขียนทับทุกรอบ)
    REORDER_QUEUE: 'reorder_queue',   // รายการที่ต้องสั่ง รอ/ระหว่างยืนยันกับ user
    BUDGET: 'budget_master',          // นำเข้าจากไฟล์ Budget ประจำปี
    PR_LOG: 'pr_po_log',              // บันทึกทุกครั้งที่ออก PR/PO เพื่อหักงบเอง
    ACTIVITY_LOG: 'activity_log',     // log ทุกขั้นตอนสำหรับ Dashboard
    MEMBERS: 'members',               // สมาชิก: ID, ชื่อ, อีเมล, role (admin/user), status, รหัสผ่าน (hash)
    REQUESTS: 'requests',             // คำขอสั่งซื้อจาก user (หัวเอกสาร: งบ, เดือน, สถานะ)
    REQUEST_ITEMS: 'request_items',   // รายการสินค้าในแต่ละคำขอ
    SETTINGS: 'settings'              // ค่าตั้งต้นอีเมลขออนุมัติ (ผู้อนุมัติ, สำเนา, ฝ่ายจัดซื้อ, ลายเซ็น)
  },

  // อีเมลของ Admin เริ่มต้น (ใช้รับอีเมลแจ้งเตือน) — แก้ได้ภายหลังในแท็บ Admin
  ADMIN_EMAILS: ['thanakorn@planbmedia.co.th'],

  // บัญชี Admin เริ่มต้น — สร้างอัตโนมัติเมื่อยังไม่มี Admin ในชีต members
  // ระบบบังคับให้เปลี่ยนรหัสผ่านหลังเข้าสู่ระบบครั้งแรก
  DEFAULT_ADMIN: { username: 'admin', password: 'admin2026' },

  // สูตร MIN/MAX (ยืนยันจากไฟล์ MIN_MAX_Calculated.xlsx เดิม)
  // Safety Stock = 0.5 x avg_per_day x lead_time_days
  // MIN          = 1.5 x avg_per_day x lead_time_days
  // MAX          = 2.5 x avg_per_day x lead_time_days
  MINMAX_FACTORS: {
    SAFETY: 0.5,
    MIN: 1.5,
    MAX: 2.5
  },

  DEFAULT_LEAD_TIME_DAYS: 30,

  // จำนวนเดือน (จาก 9 เดือนย้อนหลัง) ที่ต้องมีการเบิก ถึงจะถือว่าเป็น "ของเบิกประจำ"
  // น้อยกว่านี้ = "เบิกตามงาน" ห้ามตั้งจุดสั่งซื้อตายตัว ต้องถาม user ทุกครั้ง
  REGULAR_ITEM_MIN_MONTHS: 4,

  // จำนวนวันที่รอคำตอบจาก user ก่อนส่งเตือนซ้ำ
  REMINDER_AFTER_DAYS: 3
};
