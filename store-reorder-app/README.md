# Store Reorder AI (Plan B Media — Static Media / PG44)

## หน้าเว็บเทียบยอดสั่งซื้อ — รันได้ 2 ที่จากโค้ดชุดเดียว (`web/`)

| | GitHub Pages | Apps Script Web App |
|---|---|---|
| ลิงก์ | `https://<user>.github.io/store-reorder-app/` | `https://script.google.com/.../exec` |
| ใครเข้าได้ | สมาชิกที่เข้าสู่ระบบ (ID/รหัสผ่าน) | สมาชิกที่เข้าสู่ระบบ (ID/รหัสผ่าน) |
| ผู้ช่วย AI | ผู้ใช้ใส่ API key เอง (เก็บในเบราว์เซอร์) | ใช้ `ANTHROPIC_API_KEY` ใน Script Properties — ผู้ใช้ไม่เห็น key |
| ข้อมูล 4 ช่อง | **ข้อมูลกลางใน Google Sheet** (เรียก Apps Script ผ่าน `doPost`) | **ข้อมูลกลางใน Google Sheet** |
| สมาชิก / อีเมล / workflow | ✓ (backend = Apps Script) | ✓ |
| deploy | `.github/workflows/pages.yml` | `.github/workflows/appsscript.yml` (clasp push) |

แก้โค้ดหน้าเว็บที่ `web/` เท่านั้น แล้วรัน `npm run build` (หรือ
`powershell -ExecutionPolicy Bypass -File tools/build-appsscript.ps1`) เพื่อสร้าง
`src/Index.html` + `src/ui_*.html` สำหรับ Apps Script — ไฟล์พวกนี้ generate ห้ามแก้เอง
(GitHub Action build ใหม่ให้ทุกครั้งก่อน push อยู่แล้ว)

เปิดในเบราว์เซอร์ นำไฟล์ Store ทั้ง 4 ไฟล์เข้า → ระบบรวมเป็นตารางเดียวต่อรหัสสินค้า →
เทียบยอดแนะนำสั่ง 3 วิธีข้างกัน → ติ๊กเลือกรายการ/แก้จำนวน → ส่งออกใบขอซื้อ .xlsx

| ช่อง | ไฟล์ | ใช้ทำอะไร |
|---|---|---|
| 1 | `MIN_MAX_Calculated.xlsx` | Avg/MIN/MAX ที่กรอกเอง → ยอดสั่ง "MIN/MAX เดิม" |
| 2 | `รวมการใช้ของ.xlsx` หรือ `การใช้งานเดือน*.xls` หลายไฟล์ | ยอดเบิกจริงรายเดือน → คำนวณ MIN/MAX ใหม่ → ยอดสั่ง "ใช้จริง" |
| 3 | `ยอดคงเหลือสินค้า.xls` | คงเหลือ, ต้นทุน, มูลค่า, PC |
| 4 | `รายงานสินค้าถึงจุดสั่งซื้อ.xls` | ยอดสั่ง "Store แจ้ง" |
| ± | ไฟล์รวมที่ส่งออกจากแอปนี้รอบก่อน (ไม่บังคับ) | เทียบคงเหลือรอบนี้กับรอบก่อน |

- ลากไฟล์มาวางที่ไหนก็ได้ ระบบดูหัวตารางแล้วจัดช่องเอง · กด × เพื่อนำไฟล์ออก · เปลี่ยนไฟล์ได้ทีละช่อง
- **ไฟล์ไม่ถูกอัปโหลดไปไหน** — อ่านและคำนวณในเบราว์เซอร์ทั้งหมด (SheetJS) ข้อมูลที่นำเข้าและรายการที่เลือก
  จำไว้ในเบราว์เซอร์เครื่องนั้น (localStorage) จนกด "ล้างข้อมูลทั้งหมด"
- สูตรเดียวกับ `src/MinMax.js`: MIN = 1.5 × (ใช้จริงเฉลี่ย/เดือน ÷ 30) × lead time, MAX = 2.5 × …,
  "เบิกประจำ" = มีการเบิก ≥ 4 เดือน (ปรับได้ในหน้าเว็บ) · รายการ "เบิกตามงาน" ไม่คำนวณยอดสั่งอัตโนมัติ
- แท็บ "ผู้ช่วย AI" (ไม่บังคับ) เรียก Claude (`claude-opus-5`) ด้วย API key ของผู้ใช้เอง
  ส่งเฉพาะสรุปรายการที่มีข้อสังเกต/ที่เลือกสั่ง ไม่ส่งไฟล์ทั้งไฟล์ · key เก็บใน localStorage ของเครื่องนั้น

### เปิดใช้ GitHub Pages

1. push repo นี้ขึ้น GitHub
2. Settings > Pages > Source = **GitHub Actions** — workflow `.github/workflows/pages.yml` จะ deploy `web/` ทุกครั้งที่ push เข้า `main`
3. ได้ลิงก์ `https://<user-or-org>.github.io/store-reorder-app/`

> GitHub Pages ของ repo **private** ต้องใช้แพลน GitHub Pro/Team/Enterprise (Enterprise Cloud ตั้งให้เห็นเฉพาะคนในองค์กรได้)
> ถ้าใช้แพลนฟรีต้องเป็น repo public — ในโค้ดไม่มีข้อมูลสินค้าจริง (ไฟล์ .xls/.xlsx ถูก .gitignore ไว้)
> แต่ควรย้าย `docs/budget_mapping_draft.md` ออกก่อนถ้าจะเปิด public

### ข้อมูลกลาง (ทั้งลิงก์ Apps Script และ GitHub/Vercel)

ข้อมูล 4 ช่อง (MIN/MAX, การใช้ของ, คงเหลือ, จุดสั่งซื้อ) เก็บใน Google Sheet ฐานข้อมูล — ทุกคนเปิดมาเห็นชุดเดียวกัน

- Admin: ลากไฟล์มาวาง / "เปลี่ยนไฟล์" / × นำออก ได้ตลอด (ช่องการใช้ของเก็บได้หลายไฟล์ เช่นทีละเดือน)
- แก้รายแถวได้โดยตรงในชีต `data_minmax`, `data_usage`, `data_balance`, `data_reorder`
  (คอลัมน์ A = ไฟล์ที่มา) แล้วกด "โหลดข้อมูลใหม่" ในแอป · ประวัติการอัปโหลดอยู่ในชีต `data_files`
- User: ดูอย่างเดียว · ช่อง ± (ไฟล์รอบก่อน) ยังเป็นของแต่ละเครื่อง

### ระบบสมาชิก (ID + รหัสผ่าน) + คำขอสั่งซื้อ

เข้าสู่ระบบด้วย ID/รหัสผ่านของระบบเอง — ใช้ได้ทั้งลิงก์ Apps Script และหน้า GitHub/Vercel
(หน้า GitHub/Vercel เรียก Apps Script เป็น backend ผ่าน `doPost` ใน `src/Api.js`, ตั้ง URL ที่
`APPS_SCRIPT_URL` ใน `web/app.js`) — **Deploy: Execute as: Me · Who has access: Anyone**

- **Admin เริ่มต้น: ID `admin` / รหัส `admin2026`** (`CONFIG.DEFAULT_ADMIN`) สร้างอัตโนมัติเมื่อยังไม่มี Admin —
  เข้าครั้งแรกระบบบังคับตั้งรหัสใหม่ · ⚠ รหัสเริ่มต้นนี้อยู่ในโค้ด/GitHub จึงต้องเปลี่ยนทันที
- รหัสผ่านเก็บแบบ salt + hash ในชีต `members` (ไม่เก็บตัวจริง) · ผิด 5 ครั้งล็อก 10 นาที · เข้าระบบค้างไว้ 6 ชม. (ต่ออายุเมื่อใช้งาน)
- Admin เพิ่มสมาชิก (ตั้งรหัสเริ่มต้น → ผู้ใช้ต้องเปลี่ยนตอนเข้าครั้งแรก) / ตั้งรหัสใหม่ / เปลี่ยนสิทธิ์ / ระงับ
- **ดูรหัสผ่านของแต่ละ ID**: Admin → สมาชิก & ตั้งค่า → "รหัสผ่านเข้าใช้งาน — Admin และ User ทุกคน" ·
  รหัสเดิมเปิดดูไม่ได้ (เก็บเป็น hash) แต่กด "สร้างรหัสใหม่" (ทีละคนหรือทุกคน) แล้วระบบจะ**แสดงรหัสนั้นบนหน้าจอ**
  พร้อมปุ่มคัดลอก ไปแจ้งผู้ใช้ได้ทันที · ติ๊ก "บังคับให้เปลี่ยนรหัสเองตอนเข้าครั้งแรก" ได้ถ้าต้องการให้ใช้ได้ครั้งเดียว ·
  รหัสที่แสดงอยู่บนหน้าจอเท่านั้น ไม่ถูกบันทึกเป็นข้อความทั้งในชีตและใน `activity_log` · ทุกช่องรหัสผ่านมีปุ่ม 👁 แสดง/ซ่อน
- ทุกฟังก์ชันที่หน้าเว็บเรียกต้องผ่าน `api()` ที่ตรวจ token · ฟังก์ชันชุดเก่า (import/weekly pipeline/trigger)
  รันได้จาก editor/trigger เท่านั้น (`ownerOnly_`)
- ใส่ **อีเมล** ให้สมาชิกทุกคน — ใช้ส่งแจ้งเตือน และจับคู่กับผู้รับผิดชอบ/หัวหน้าใน Master PC

| บทบาท | ทำอะไรได้ |
|---|---|
| ผู้ใช้ใหม่ | "สมัครสมาชิก" (ID, รหัสผ่าน, ชื่อ, อีเมล) → รอ Admin อนุมัติ (Admin ได้อีเมล) |
| User | เทียบยอด/เลือกรายการ+จำนวน → **เลือก Budget + เดือนที่จะใช้ของ** → ส่งคำขอให้ Admin · ดู/ยกเลิกคำขอของตัวเอง |
| Admin | ทุกอย่างของ User + อนุมัติ/ไม่อนุมัติคำขอ (แจ้งผู้ขอทางอีเมล) · จัดการสมาชิก (อนุมัติ/เปลี่ยนสิทธิ์/ระงับ) · นำเข้าไฟล์ Budget · Master PC |

#### แยกเมนู Admin / User

| | User (ผู้ใช้งาน) | Admin (ผู้ดูแลระบบ) |
|---|---|---|
| เมนู | หน้าแรก · **สั่งซื้อตาม Budget** · คำขอของฉัน · งบประมาณ · รายการแนะนำจาก Store | ภาพรวม · **ตรวจคำขอสั่งซื้อ** · สั่งซื้อตาม Budget · งบประมาณ · รายการแนะนำ · ข้อมูล Store · ตารางรวม · สมาชิก & ตั้งค่า · ผู้ช่วย AI |
| ทำอะไรได้ | เลือก Budget + เดือน → ใส่รายการ → แนบใบเสนอราคา → ส่ง · เรียกคำขอกลับมาแก้ไขแล้วส่งใหม่ · ขอ Over Budget | ทุกอย่างของ User + ตรวจ/แก้ไข/ส่งกลับ/ลบคำขอ · เปิดดูใบเสนอราคา · ส่งเมล · บันทึกผลอนุมัติ · PR/PO · นำเข้าข้อมูล Store/Budget · จัดการสมาชิก |

สีและป้ายบนหัวเรื่องบอกบทบาท (Admin = ม่วง, User = น้ำเงิน) เมนูของอีกฝั่งจะไม่แสดงเลย

#### โหมดทดลอง (ทดลองได้ทั้ง Admin และ User ก่อนใช้จริง)

เปิดหน้าเว็บด้วย `?demo=1` (มีลิงก์ "ทดลองใช้งาน" ที่หน้าเข้าสู่ระบบด้วย) — โค้ดฝั่งเซิร์ฟเวอร์ชุดเดียวกัน (`src/*.js`)
จะรันในเบราว์เซอร์พร้อมข้อมูลตัวอย่าง: **ไม่แตะ Google Sheet จริง ไม่เขียน Drive จริง และไม่ส่งอีเมลจริง**

- บัญชีทดลอง: `admin` / `user` / `user2` / `head` — รหัสผ่าน `demo1234` (หรือกดปุ่ม "เข้าเป็น Admin / User")
- แถบสีเหลืองด้านบน: สลับบทบาท · **อีเมลที่ส่ง** (ดูอีเมลทุกฉบับพร้อมไฟล์แนบที่ระบบจะส่งจริง) · เริ่มข้อมูลใหม่ · ออกจากโหมดทดลอง
- ข้อมูลทดลองเก็บใน localStorage ของเบราว์เซอร์นั้น แยกจากข้อมูลจริงคนละชุด
- ไฟล์: `web/demo.js` (สวิตช์), `web/demo/mock.js` (จำลองบริการ Google), `web/demo/server.js` (สร้างจาก `src/*.js` ตอน build), `web/demo/seed.js` (ข้อมูลตัวอย่าง)
- บนหน้า Apps Script จริงจะไม่มีโหมดทดลอง (เป็นระบบจริงเสมอ)

#### ขั้นตอนขออนุมัติสั่งซื้อ (ตามอีเมล "ขออนุมัติสั่งซื้ออุปกรณ์ (ทีม Static)")

| ขั้น | ใคร | ในแอป | สถานะ |
|---|---|---|---|
| 1 | User | **ตรวจงบประมาณ → เลือก Budget + เดือน → ใส่รายการสั่งซื้อในงบนั้น** (ค้นหาจาก Store, กดรายการแนะนำ หรือเพิ่มของที่ไม่มีใน Store) · ใส่ขอบเขตการใช้งาน · **แนบใบเสนอราคาได้** · งบไม่พอให้ติ๊ก **ขอ Over Budget** พร้อมเหตุผล แล้วกดส่งให้ Admin | รอ Admin ตรวจ |
| 1b | User | **เรียกคำขอกลับมาแก้ไข** ได้ (ก่อน Admin ส่งเมล) หรือเมื่อ Admin ส่งกลับ/ไม่อนุมัติ → แก้ในหน้าสั่งซื้อ แล้ว "บันทึกและส่งใหม่" (เลขคำขอเดิม) | รอผู้ขอแก้ไข → รอ Admin ตรวจ |
| 2 | Admin | ตรวจรายการ + **เปิดดูใบเสนอราคา** · แก้ไขข้อมูลได้ทุกช่อง (จำนวน ราคา Budget เดือน) · ลบรายการ/ลบคำขอ · หรือส่งกลับให้ผู้ขอแก้ | รอ Admin ตรวจ / รอผู้ขอแก้ไข |
| 3 | Admin | **ส่งเมลรายการสั่งซื้อ + Budget** ถึง `thanakorn@planbmedia.co.th` (ค่าตั้งต้น แก้ได้) — เนื้อหาตามรูปแบบเดิม (Budget เดือน…, งบตาม BG vs ค่าใช้จ่ายโดยประมาณ, "สื่อ … หมวด…", ยอดคงเหลือ STORE, ตาราง 14 คอลัมน์, Over Budget + เหตุผล) **พร้อมไฟล์แนบ**: ไฟล์ Excel รายการสั่งซื้อ + Budget ที่ระบบสร้างให้ และใบเสนอราคาของผู้ขอ | รอผู้บริหารอนุมัติ |
| 4 | ผู้บริหาร / Admin | ตอบ "Approved" ทางอีเมล → Admin กด **บันทึก: ได้รับอนุมัติแล้ว** (ผู้บริหารที่เป็นสมาชิกกดอนุมัติในแอปเองได้) | อนุมัติ / ไม่อนุมัติ |
| 5 | Admin | **ส่งแจ้งฝ่ายจัดซื้อ** ("ได้รับการอนุมัติสั่งซื้ออุปกรณ์ …" แนบไฟล์เหมือนกัน) → ใส่เลข **PR/PO** (ตัดงบลง `pr_po_log` ตาม PC × Budget × เดือน) | ออก PR/PO แล้ว |
| 6 | Admin | **รับของเข้าแล้ว** | รับของแล้ว |

ทางเลือก (Admin): ส่งรายการให้ผู้รับผิดชอบแต่ละ PC ยืนยันเองก่อน (**ส่งให้ PC ยืนยันทางอีเมล** → [รอ PC ยืนยัน] →
PC แก้จำนวน เลือก Budget + เดือน และขอ Over Budget ได้ → เข้าขั้น 2) · ไม่ตอบ 3 วัน → อีเมลเตือน

ทุกขั้นแจ้งอีเมลผู้เกี่ยวข้อง และบันทึกใน `activity_log` · แท็บคำขอมีตัวกรอง "รอฉันดำเนินการ", ค้นหา และสรุปจำนวน/มูลค่าตามสถานะ ·
ค่าตั้งต้นของอีเมล (ผู้รับ, สำเนา, ฝ่ายจัดซื้อ, ชื่อทีม, ลายเซ็น) ตั้งที่ Admin → **ตั้งค่าอีเมลส่งรายการสั่งซื้อ + Budget** (ชีต `settings`)

- `CONFIG.ADMIN_EMAILS[0]` = อีเมลของ Admin เริ่มต้น (รับแจ้งเตือน)
- Budget: Admin นำเข้าไฟล์ "Budget STT 2026 - Revise-Budget" ที่ สมาชิก & ตั้งค่า (หรือลากไฟล์มาวาง) —
  รวมเป็น Budget ต่อ บริษัท × Media Location × GL Code × เดือน (พร้อม Remark ตาม Budget) · ถ้าช่อง Media Location ว่าง
  ใช้ชื่อจากคอลัมน์ Calculation · แถวที่ไม่มี GL Code/เลขเดือนจะถูกข้ามและแจ้งจำนวน
- คงเหลือ = งบเดือนนั้น (Revise Budget ถ้ามี ไม่งั้น Budget) − Actual − รายการในคำขอที่รอตรวจ/รออนุมัติ/อนุมัติแล้ว
  (นับตาม Budget + เดือนของแต่ละรายการ · คำขอที่ "รอผู้ขอแก้ไข" ไม่จองงบ)
- ใบเสนอราคา: อัปโหลดจากหน้าสั่งซื้อ (PDF/รูป/Excel/Word ไฟล์ละ ≤ 5 MB รวม ≤ 10 MB ไม่เกิน 5 ไฟล์) เก็บในโฟลเดอร์ Drive
  ของเจ้าของสคริปต์ (Script Property `QUOTE_FOLDER_ID`) และลงทะเบียนในชีต `request_files` — เปิดดูผ่านระบบ (ไม่ต้องแชร์ Drive)
- ชีตในฐานข้อมูล: `members`, `requests`, `request_items`, `request_files`, `budget_master`, `settings`; คำขอที่ออก PR/PO แล้วบันทึกลง
  `pr_po_log` แยกตาม PC × Budget × เดือน (ใช้ต่อใน `Budget.js`) · คำขอเก่าสถานะ "รออนุมัติ" แสดงเป็น "รอ Admin ตรวจ"

### เปิดใช้บน Apps Script

ฐานข้อมูล: [Google Sheet](https://docs.google.com/spreadsheets/d/1V91agWVoocDbJ54KLxAN27btFaYvEewR10NHsA3TVSY/edit)
(ตั้งไว้แล้วใน `src/Config.js`) — ระบบสร้างชีตของตัวเอง (`raw_*`, `reorder_queue`, `activity_log` …) ไม่ทับชีตเดิม

**แบบ A — คัดลอกวาง 1 ไฟล์ (ไม่ต้องติดตั้งอะไร)** — ใช้ `build/apps-script/Code.gs` (รวมโค้ดทุกไฟล์ + หน้าเว็บไว้แล้ว):

1. Sheet > ส่วนขยาย > Apps Script → **ลบไฟล์อื่นทั้งหมด** (.gs / .html เก่า — ไฟล์เก่าที่ชื่อฟังก์ชันซ้ำจะทับของใหม่)
2. วาง `Code.gs` ทั้งไฟล์ → 💾
3. การตั้งค่าโปรเจกต์ > ติ๊ก "แสดงไฟล์ appsscript.json" → วาง `build/apps-script/appsscript.json` ลงไฟล์นั้น (ห้ามวางเป็น .gs)
4. เลือกฟังก์ชัน **`setup`** → เรียกใช้ → อนุญาตสิทธิ์ → ดูบันทึก: ทุกบรรทัดต้องเป็น ✓
5. Deploy (Execute as: Me · Who has access: Anyone) → เปิดลิงก์ → เข้า `admin` / `admin2026`

ฟังก์ชันอื่น (`login`, `importBudget` ฯลฯ) กด Run ใน editor ไม่ได้ — ต้องเรียกจากหน้าเว็บหลังเข้าสู่ระบบ (จะขึ้น "กรุณาเข้าสู่ระบบ" ซึ่งเป็นปกติ)

**แบบ B — clasp** (ต้องมี Node.js):

1. ทำ Setup ข้อ 1–3 ด้านล่าง (clasp login, สร้างโปรเจกต์, กรอก `SPREADSHEET_ID`)
2. Script Properties: เพิ่ม `ANTHROPIC_API_KEY` ถ้าจะใช้แท็บผู้ช่วย AI
3. `npm run push` (build + clasp push) แล้ว Deploy > New deployment > Web app
   (Execute as: Me, Who has access: Anyone) — ได้ลิงก์ `/exec`
4. ลิงก์ `/exec` = หน้าเว็บเต็ม (ลิงก์ในอีเมล `?req=...` เปิดคำขอนั้นให้เลย)

ให้ GitHub push ขึ้น Apps Script อัตโนมัติ: ตั้ง repo secrets `CLASPRC_JSON` (เนื้อหาไฟล์
`~/.clasprc.json` หลัง `clasp login`), `SCRIPT_ID` (ถ้าไม่ได้ commit `.clasp.json`) และ
`DEPLOYMENT_ID` (ถ้าอยากให้ลิงก์ `/exec` เดิมอัปเดตตาม) — ไม่ตั้ง `CLASPRC_JSON` ก็แค่ build ไม่ push

> `CLASPRC_JSON` คือ token ของบัญชี Google ที่ login — ใช้บัญชีที่เป็นเจ้าของโปรเจกต์นี้เท่านั้น
> และใส่ใน Secrets เท่านั้น ห้าม commit

### ทดสอบในเครื่อง

ใช้ web server อะไรก็ได้ที่เสิร์ฟโฟลเดอร์ `web/` เช่น VS Code extension "Live Server" หรือ `npx serve web`

---

## Apps Script pipeline (`src/`) — ระยะถัดไป (อีเมล/ยืนยัน/งบ)

Apps Script project implementing the workflow in `../store-ai-workflow.html`:
Store export files → recompute MIN/MAX from **actual usage** → email each PC
owner to confirm quantities → check annual budget → draft PR/PO.

## สถานะตอนนี้ (2569-09-18)

- ✅ วิเคราะห์ข้อมูลจริง 9 เดือน พบว่า MIN/MAX เดิม (`MIN_MAX_Calculated.xlsx`) กรอก
  Avg/เดือนสูงเกินจริงเกือบทุกรายการ — ดูผลคำนวณใหม่ใน `../MIN_MAX_Recalculated_2569.xlsx`
- ✅ โครงโค้ด Apps Script ทั้ง pipeline อยู่ใน `src/` (import → minmax → reorder → email → confirm web app)
- ⛔ `Budget.js` ยัง **ใช้งานจริงไม่ได้** — รอ mapping PC ↔ Budget location/GL จากฝ่ายบัญชี
  ดู `docs/budget_mapping_draft.md`
- ⛔ ยังไม่ได้ผูกกับ Google Sheet จริง / Drive folder จริง / อีเมลจริง — เป็น skeleton ที่รันได้เมื่อกรอกค่า config

## โครงสร้าง

```
src/
  Config.js       ค่าคงที่ + ชื่อชีตทั้งหมด (แก้ SPREADSHEET_ID ที่นี่)
  Import.js       ดึงไฟล์ .xls จาก Drive แปลงเป็น Google Sheet แล้วล้าง/รวมข้อมูล
  MinMax.js       คำนวณ MIN/MAX จากยอดเบิกจริง 9 เดือน + แยกประเภทเบิกประจำ/ตามงาน/dead stock
  Reorder.js      สร้างคิวรายการที่ต้องสั่ง แยกตาม PC
  Notify.js       ส่งอีเมลแจ้ง PC owner + เตือนซ้ำถ้าไม่ตอบใน 3 วัน
  WebApp.js       doGet → หน้าเว็บ (Ui.js) ทุกลิงก์; submitConfirmation ชุดเก่ารันได้จาก editor เท่านั้น
  confirm.html    หน้า UI ของ WebApp.js
  Ui.js           เสิร์ฟหน้าเทียบยอดสั่งซื้อ + askClaudeFromUi
  DataFiles.js    ข้อมูลกลาง 4 ช่อง (ชีต data_*) อ่าน/เพิ่ม/เปลี่ยน/นำออก
  Members.js      สมาชิก Admin/User แบบ ID + รหัสผ่าน (hash), session token
  Api.js          ทางเข้าเดียวของหน้าเว็บ: api(token, fn, args) + doPost สำหรับหน้า GitHub/Vercel
  SheetUtil.js    ตัวช่วยอ่าน/เขียนชีต + ownerOnly_
  Setup.js        setup(): รันใน editor เพื่อขอสิทธิ์และตรวจการติดตั้ง (✓/✗)
  Requests.js     workflow คำขอ: เลือก Budget รายรายการ/Over Budget → Admin ตรวจ/แก้ไข/ส่งกลับ/ลบ → PR/PO → รับของ
  Quotes.js       ใบเสนอราคาที่ผู้ขอแนบ (เก็บใน Drive + ชีต request_files, เปิดดูผ่านระบบ)
  Approval.js     อีเมลขออนุมัติผู้บริหาร (รูปแบบทีม Static), บันทึกผลอนุมัติ, แจ้งฝ่ายจัดซื้อ, ตั้งค่าอีเมล
  MasterPc.js     PC → อีเมลผู้รับผิดชอบ / หัวหน้า / Budget ตั้งต้น
  Index.html, ui_*.html   GENERATED จาก web/ โดย tools/build-appsscript.ps1
  Budget.js       เช็คงบก่อนออก PR/PO (บล็อกจนกว่าจะ mapping เสร็จ)
  Ai.js           เรียก Claude API มาช่วยร่างข้อความ (ไม่ยุ่งกับตัวเลข/อนุมัติ)
  Main.js         runWeeklyPipeline() + ตั้ง trigger รายสัปดาห์ + activity log
docs/
  budget_mapping_draft.md   ร่างจับคู่ PC กับ Budget + คำถามที่ต้องถามฝ่ายบัญชี
```

## Setup

### 1. เครื่องมือที่ต้องติดตั้งก่อน (เครื่องนี้ยังไม่มี)

```bash
winget install OpenJS.NodeJS.LTS
```

ปิด/เปิด terminal ใหม่ แล้วติดตั้ง clasp:

```bash
npm install -g @google/clasp
clasp login
```

เปิด Apps Script API ที่ https://script.google.com/home/usersettings ก่อน `clasp login`
ครั้งแรก มิฉะนั้น push จะ error

### 2. สร้าง Google Sheet ฐานข้อมูล + ผูก Apps Script

สร้าง Google Sheet เปล่า 1 ไฟล์ (นี่คือฐานข้อมูลของระบบ ไม่ใช่ไฟล์ export จาก Store)
แล้ว Extensions > Apps Script เพื่อได้ scriptId จาก URL หรือรันคำสั่งนี้ในโฟลเดอร์นี้:

```bash
clasp create --type sheets --title "Store Reorder AI" --rootDir src
```

จะได้ไฟล์ `.clasp.json` (อย่า commit ไฟล์นี้ถ้ามี token ส่วนตัวปน — ปกติมีแค่ scriptId/rootDir ก็คอมมิตได้)

### 3. กรอก config

- `src/Config.js`: `SPREADSHEET_ID`
- `src/Import.js`: `IMPORT_FOLDER_ID` (โฟลเดอร์ Drive ที่ Store อัปโหลดไฟล์ export)
- Project Settings ในหน้า Apps Script > Script Properties: เพิ่ม `ANTHROPIC_API_KEY` (ถ้าจะใช้ `Ai.js`)

### 4. Push และรันครั้งแรกด้วยมือก่อน (อย่าเพิ่ง auto-trigger)

```bash
clasp push
```

ในหน้า Apps Script editor รันทีละฟังก์ชันแล้วดูผลใน Sheet ก่อนตั้ง trigger:

1. `importAllFiles`
2. `recomputeMinMax` — เทียบผลกับ `../MIN_MAX_Recalculated_2569.xlsx` ว่าตรงกัน
3. `buildReorderQueue`
4. `sendReorderEmails` — ทดสอบกับอีเมลตัวเองก่อน อย่าเพิ่งส่งจริงทุก PC

เมื่อมั่นใจแล้วค่อยรัน `installWeeklyTrigger` ครั้งเดียว

## GitHub

Repo นี้ตั้งใจให้เป็น **private** (มีข้อมูล lead time/ราคาสินค้าที่เป็นความลับทางธุรกิจ
แม้จะไม่มีไฟล์ Excel จริงอยู่ในนี้ก็ตาม — ดู `.gitignore`)

```bash
gh repo create <org-or-user>/store-reorder-app --private --source=. --remote=origin
git add -A
git commit -m "Initial Apps Script skeleton for Store reorder workflow"
git push -u origin main
```

ค่อยเพิ่ม GitHub Actions ให้ `clasp push` อัตโนมัติตอน merge เข้า `main` ทีหลัง — ต้องเก็บ
`clasp login` credentials เป็น GitHub Secret (`CLASPRC_JSON`) ก่อน ยังไม่ได้ตั้งในรอบนี้

## เรื่องที่ต้องตัดสินใจก่อนไปต่อ

ดู `docs/budget_mapping_draft.md` — เป็นตัวบล็อกหลักก่อนเปิดใช้ `Budget.js` และก่อนส่งอีเมล
ขออนุมัติเกินงบ (ขั้นตอน 8b ใน workflow diagram)
