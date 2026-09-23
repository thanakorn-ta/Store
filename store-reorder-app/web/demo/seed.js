// Demo / trial mode: sample data (fictional numbers) + the in-page API used by app.js instead of
// Apps Script. Loaded after demo/mock.js and demo/server.js (see web/demo.js).
(function () {
  const M = window.GASMOCK;
  const ACCOUNTS = [
    { username: 'admin', name: 'Admin ทดลอง', email: 'admin@demo.local', role: 'admin' },
    { username: 'user', name: 'User ทดลอง', email: 'user@demo.local', role: 'user' },
    { username: 'user2', name: 'ทีมซ่อมบำรุง (ทดลอง)', email: 'user2@demo.local', role: 'user' },
    { username: 'head', name: 'ผู้บริหาร (ทดลอง)', email: 'head@demo.local', role: 'user' }
  ];
  const PASSWORD = 'demo1234';

  function seed() {
    const ss = db_();
    const now = new Date();
    ACCOUNTS.forEach(a => {
      const salt = newSalt_();
      appendObject_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADER, { ...a, status: 'active', pw_hash: hashPassword_(PASSWORD, salt),
        pw_salt: salt, must_change: false, created_at: now, updated_at: now, updated_by: 'demo' });
    });

    // ---- budget: company x location x GL, 12 months (fictional amounts)
    const lines = [
      ['Plan B Media', 'Metro Poster', 'Metro Poster', '530090', 'ค่าวัสดุสิ้นเปลือง', 'ค่าวัสดุสิ้นเปลืองใช้ไป', 3000, 'ค่าสีสเปรย์พ่น เคลดดิ่ง+โครงเหล็ก'],
      ['Plan B Media', 'Metro Poster', 'Metro Poster', '530120', 'ค่าซ่อมแซมบำรุงรักษา', 'ค่าซ่อมแซมบำรุงรักษาป้ายโฆษณา', 46750, 'กล่องคอนโทรล / Main Controller'],
      ['Ads Cuisine', 'Metro Poster', 'Metro Poster', '530090', 'ค่าวัสดุสิ้นเปลือง', 'ค่าวัสดุสิ้นเปลืองใช้ไป', 24000, 'สายไฟ + อุปกรณ์ติดตั้ง'],
      ['Plan B Media', 'Street Static', 'Cookies', '530090', 'ค่าวัสดุสิ้นเปลือง', 'ค่าวัสดุสิ้นเปลืองใช้ไป', 5000, 'วัสดุติดตั้งป้าย'],
      ['Plan B Media', 'Street Static', 'Bangkok Linear Park', '530010', 'ค่าไฟฟ้าป้ายโฆษณา', 'Electricity', 7000, ''],
      ['Plan B Media', 'Transit', 'BTS Pillar', '530120', 'ค่าซ่อมแซมบำรุงรักษา', 'ค่าซ่อมแซมบำรุงรักษาป้ายโฆษณา', 30000, 'หลอดไฟ LED / บัลลาสต์']
    ];
    const rows = [];
    lines.forEach(([company, mediaType, loc, gl, glName, group, plan, remark], li) => {
      for (let m = 1; m <= 12; m++) {
        const actual = m <= 8 ? Math.round(plan * (0.35 + ((li * 7 + m * 3) % 10) / 20)) : 0;
        rows.push({ key: `${company}|${loc}|${gl}`, company, division: 'Street-Static', media_type: mediaType, media_group: mediaType,
          media_location: loc, expense_group: group, gl_code: gl, gl_name: glName, year: 2026, month_number: m,
          budget: plan, revise_budget: '', actual, plan, remark });
      }
    });
    writeBudgetMaster_(rows);

    // ---- Store data (the 4 files)
    const items = [
      ['ST-1001', 'น้ำยาล้างคราบสิ่งสกปรกเอนกประสงค์', 'PC-MTR', 'Metro Poster', 'ลิตร', 90, 4, 6, 30],
      ['ST-1002', 'กล่องคอนโทรล SPDE120A', 'PC-MTR', 'Metro Poster', 'กล่อง', 8500, 2, 3, 45],
      ['ST-1003', 'สายไฟ FD-0.6/1KV-CV 4x6', 'PC-MTR', 'Metro Poster', 'เมตร', 214, 0, 40, 30],
      ['ST-1004', 'สีสเปรย์ เทา', 'PC-MTR', 'Metro Poster', 'กระป๋อง', 85, 12, 10, 15],
      ['ST-1005', 'Main Controller', 'PC-MTR', 'Metro Poster', 'ตัว', 10000, 2, 1, 60],
      ['ST-2001', 'หลอดไฟ LED T8 18W', 'PC-BTS', 'BTS Pillar', 'หลอด', 120, 25, 30, 20],
      ['ST-2002', 'บัลลาสต์อิเล็กทรอนิกส์', 'PC-BTS', 'BTS Pillar', 'ตัว', 350, 3, 6, 20],
      ['ST-2003', 'เทปพันสายไฟ', 'PC-BTS', 'BTS Pillar', 'ม้วน', 25, 40, 12, 7],
      ['ST-3001', 'สกรูยึดป้าย 2 นิ้ว', 'PC-CK', 'Cookies', 'กล่อง', 180, 1, 3, 14],
      ['ST-3002', 'ซิลิโคนกันน้ำ', 'PC-CK', 'Cookies', 'หลอด', 95, 0, 5, 14],
      ['ST-3003', 'แผ่นอะคริลิคใส 3 มม.', 'PC-CK', 'Cookies', 'แผ่น', 650, 6, 1, 30],
      ['ST-3004', 'ผ้าเช็ดทำความสะอาด', 'PC-CK', 'Cookies', 'แพ็ค', 60, 30, 0, 7]
    ];
    const put = (sheet, fields, recs, file) => {
      const sh = ensureSheet_(ss, sheet, ['source_file'].concat(fields));
      recs.forEach(r => sh.appendRow([file].concat(fields.map(f => r[f] == null ? '' : r[f]))));
      appendObject_(ss, 'data_files', DATA_FILES_HEADER, { slot: sheet.replace('data_', ''), file_name: file, rows: recs.length, uploaded_by: 'admin@demo.local', uploaded_at: now });
    };
    put('data_balance', DATA_SLOTS.balance.fields, items.map(([sku, name, pc, pcName, unit, cost, bal]) =>
      ({ sku, name, pc, pcName, location: '', balance: bal, unit, cost, value: bal * cost })), 'ยอดคงเหลือสินค้า (ตัวอย่าง).xls');
    put('data_minmax', DATA_SLOTS.minmax.fields, items.map(([sku, name, pc, pcName, unit, cost, bal, avg, lead]) =>
      ({ sku, name, group: pcName, avg: avg * 1.5, unit, lead, safety: '', min: Math.ceil(avg * 1.5), max: Math.ceil(avg * 3), category: '' })), 'MIN_MAX (ตัวอย่าง).xlsx');
    const usage = [];
    items.forEach(([sku, name, pc, pcName, unit, cost, bal, avg], i) => {
      for (let m = 1; m <= 8; m++) {
        const qty = avg ? Math.max(0, Math.round(avg * (0.6 + ((i + m) % 5) / 5))) : 0;
        if (qty) usage.push({ sku, name, group: pcName, month: m, year: 2569, qty, unit, price: cost, total: qty * cost });
      }
    });
    put('data_usage', DATA_SLOTS.usage.fields, usage, 'รวมการใช้ของ (ตัวอย่าง).xlsx');
    put('data_reorder', DATA_SLOTS.reorder.fields, items.filter(x => x[6] < x[7]).map(([sku, name, pc, pcName, unit, cost, bal, avg]) =>
      ({ sku, name, pc, pcName, balance: bal, min: avg, need: Math.max(1, avg * 2 - bal), unit, lot: '' })), 'รายงานสินค้าถึงจุดสั่งซื้อ (ตัวอย่าง).xls');

    ensureSheet_(ss, CONFIG.SHEETS.MASTER_PC, ['pc_code', 'pc_name', 'owner_emails', 'manager_email', 'default_budget_key', 'note']);
    [['PC-MTR', 'Metro Poster', 'user@demo.local'], ['PC-BTS', 'BTS Pillar', 'user2@demo.local'], ['PC-CK', 'Cookies', 'user@demo.local']]
      .forEach(p => M.sheets[CONFIG.SHEETS.MASTER_PC].appendRow([p[0], p[1], p[2], 'head@demo.local', '', '']));

    // ---- two sample requests so every screen has something to show
    const as = (u, fn) => { CURRENT_USER_ = toMe_(readMembers_(ss).find(m => m.username === u)); try { return fn(); } finally { CURRENT_USER_ = null; } };
    const A = 'Plan B Media|Metro Poster|530090', B = 'Plan B Media|Metro Poster|530120';
    const quote = 'data:application/pdf;base64,' + btoa('%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj 3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 144]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj 4 0 obj<</Length 60>>stream\nBT /F1 14 Tf 24 90 Td (Quotation - DEMO sample) Tj ET\nendstream endobj 5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF');
    const done = as('user', () => submitOrderRequest({ note: 'ตัวอย่าง: สั่งเดือนก่อน', items: [
      { sku: 'ST-1001', name: 'น้ำยาล้างคราบสิ่งสกปรกเอนกประสงค์', pc: 'PC-MTR', pcName: 'Metro Poster', unit: 'ลิตร', qty: 10, unitCost: 90, note: 'ใช้ล้างแกนมอเตอร์ป้าย Metro', budgetKey: A, month: 9, storeBalance: 4 }
    ] }));
    as('admin', () => { sendApprovalEmail(done.id, { to: 'thanakorn@planbmedia.co.th' }); recordApproval(done.id, true, 'Approved (ตัวอย่าง)'); });
    as('user', () => submitOrderRequest({ note: 'ตัวอย่าง: รอ Admin ตรวจ', items: [
      { sku: 'ST-1002', name: 'กล่องคอนโทรล SPDE120A', pc: 'PC-MTR', pcName: 'Metro Poster', unit: 'กล่อง', qty: 2, unitCost: 8500, note: 'Spare Part ซ่อมบำรุง', budgetKey: B, month: 10, storeBalance: 2 },
      { sku: 'ST-1004', name: 'สีสเปรย์ เทา', pc: 'PC-MTR', pcName: 'Metro Poster', unit: 'กระป๋อง', qty: 6, unitCost: 85, note: 'พ่นโครงเหล็ก', budgetKey: A, month: 10, storeBalance: 12 }
    ], files: [{ name: 'ใบเสนอราคา-ตัวอย่าง.pdf', type: 'application/pdf', data: quote }] }));
    M.mail.length = 0; // start with an empty outbox
  }

  if (!M.load()) { seed(); M.save(); }

  // app.js calls this instead of Apps Script. JSON round trip = what the real transport returns.
  window.DEMO_API = (token, fn, args) => {
    try {
      const out = api(token, fn, args || []);
      return out === undefined ? null : JSON.parse(JSON.stringify(out));
    } finally {
      M.save();
    }
  };
  window.DEMO = {
    accounts: ACCOUNTS.map(a => ({ ...a, password: PASSWORD })),
    outbox: () => M.mail.slice().reverse(),
    reset: () => { M.reset(); }
  };
})();
