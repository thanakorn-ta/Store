# Store Reorder AI (Plan B Media — Static Media / PG44)

## หน้าเว็บเทียบยอดสั่งซื้อ — รันได้ 2 ที่จากโค้ดชุดเดียว (`web/`)

| | GitHub Pages | Apps Script Web App |
|---|---|---|
| ลิงก์ | `https://<user>.github.io/store-reorder-app/` | `https://script.google.com/.../exec` |
| ใครเข้าได้ | ตามสิทธิ์ repo/Pages | คนในโดเมน planbmedia (ตาม `appsscript.json`) |
| ผู้ช่วย AI | ผู้ใช้ใส่ API key เอง (เก็บในเบราว์เซอร์) | ใช้ `ANTHROPIC_API_KEY` ใน Script Properties — ผู้ใช้ไม่เห็น key |
| ข้อมูล 4 ช่อง | เก็บในเบราว์เซอร์ของแต่ละคน | **ข้อมูลกลางใน Google Sheet** — ทุกคนเห็นชุดเดียวกัน Admin แก้ได้ |
| สั่งซื้อ | ส่งออก .xlsx | ส่งออก .xlsx **+ ส่งคำขอให้ Admin** (เลือก Budget + เดือน, ระบบสมาชิก) |
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

### ข้อมูลกลาง (เฉพาะเวอร์ชัน Apps Script)

ข้อมูล 4 ช่อง (MIN/MAX, การใช้ของ, คงเหลือ, จุดสั่งซื้อ) เก็บใน Google Sheet ฐานข้อมูล — ทุกคนเปิดมาเห็นชุดเดียวกัน

- Admin: ลากไฟล์มาวาง / "เปลี่ยนไฟล์" / × นำออก ได้ตลอด (ช่องการใช้ของเก็บได้หลายไฟล์ เช่นทีละเดือน)
- แก้รายแถวได้โดยตรงในชีต `data_minmax`, `data_usage`, `data_balance`, `data_reorder`
  (คอลัมน์ A = ไฟล์ที่มา) แล้วกด "โหลดข้อมูลใหม่" ในแอป · ประวัติการอัปโหลดอยู่ในชีต `data_files`
- User: ดูอย่างเดียว · ช่อง ± (ไฟล์รอบก่อน) ยังเป็นของแต่ละเครื่อง
- เวอร์ชัน GitHub Pages/Vercel ไม่มี server จึงยังเก็บข้อมูลในเบราว์เซอร์ของแต่ละคนเหมือนเดิม
  (ไม่ได้ฝังข้อมูลสต๊อก/ราคาไว้ในเว็บสาธารณะ)

### ระบบสมาชิก + คำขอสั่งซื้อ (เฉพาะเวอร์ชัน Apps Script)

ยืนยันตัวตนด้วยบัญชี Google ของบริษัท (ไม่มีรหัสผ่านแยก) — **ต้อง Deploy แบบ
"Execute as: Me" + "Who has access: Anyone within planbmedia.co.th"** ถ้าตั้งเป็น "Anyone"
ระบบจะอ่านอีเมลผู้ใช้ไม่ได้และเข้าใช้งานไม่ได้

| บทบาท | ทำอะไรได้ |
|---|---|
| ผู้ใช้ใหม่ | เปิดลิงก์ → กรอกชื่อ "ขอเข้าใช้งาน" → Admin ได้อีเมล |
| User | เทียบยอด/เลือกรายการ+จำนวน → **เลือก Budget + เดือนที่จะใช้ของ** → ส่งคำขอให้ Admin · ดู/ยกเลิกคำขอของตัวเอง |
| Admin | ทุกอย่างของ User + อนุมัติ/ไม่อนุมัติคำขอ (แจ้งผู้ขอทางอีเมล) · จัดการสมาชิก (อนุมัติ/เปลี่ยนสิทธิ์/ระงับ) · นำเข้าไฟล์ Budget · Master PC |

#### ขั้นตอนตาม workflow (`../store-ai-workflow.html`)

| ขั้น | ใคร | ในแอป | สถานะ |
|---|---|---|---|
| 1–4 | Admin | นำเข้าไฟล์ Store → รวม/คำนวณ MIN/MAX → เลือกรายการที่ต้องสั่ง | |
| 5 | Admin | แท็บรายการที่เลือกสั่ง → **ส่งให้ PC ยืนยันทางอีเมล** (แยกตาม PC, ส่งถึงผู้รับผิดชอบใน Master PC, AI ร่างอีเมลถ้ามี API key) | รอ PC ยืนยัน |
| 6 | ผู้รับผิดชอบ PC | เปิดลิงก์ในอีเมล → แก้จำนวน/ใส่ 0 → เลือก Budget + เดือน → ยืนยัน (หรือ "ไม่สั่งรอบนี้") · ไม่ตอบ 3 วัน → อีเมลเตือน (Admin เปิดที่แท็บ Admin) | รออนุมัติ / PC ปฏิเสธ |
| 7–8 | Admin | ตรวจงบ: พอ → อนุมัติ · ไม่พอ → **ขออนุมัติเกินงบ** ส่งหัวหน้า PC | อนุมัติ / รอหัวหน้าอนุมัติเกินงบ |
| 8b | หัวหน้า PC | เปิดลิงก์ → อนุมัติ/ไม่อนุมัติเกินงบ → กลับไปให้ Admin | รออนุมัติ / ไม่อนุมัติ |
| 9 | Admin | ใส่เลข **PR/PO** → ตัดงบ (บันทึก `pr_po_log`) | ออก PR/PO แล้ว |
| 10 | Admin | **รับของเข้าแล้ว** | รับของแล้ว |

ทุกขั้นแจ้งอีเมลผู้เกี่ยวข้อง และบันทึกใน `activity_log` · แท็บคำขอมีตัวกรอง "รอฉันดำเนินการ" และสรุปจำนวน/มูลค่าตามสถานะ ·
User ยังส่งคำขอเองจากตะกร้าได้ (เข้าที่ขั้น 7 ทันที)

- Admin ตั้งต้น: `CONFIG.ADMIN_EMAILS` ใน `src/Config.js` (เข้าครั้งแรกได้สิทธิ์ Admin อัตโนมัติ)
- Budget: Admin นำเข้าไฟล์ "Budget STT 2026 - Revise-Budget" ที่แท็บ Admin (หรือลากไฟล์มาวาง) —
  รวมเป็น Budget ต่อ บริษัท × Media Location × GL Code × เดือน · ถ้าช่อง Media Location ว่าง
  ใช้ชื่อจากคอลัมน์ Calculation · แถวที่ไม่มี GL Code/เลขเดือนจะถูกข้ามและแจ้งจำนวน
- คงเหลือ = งบเดือนนั้น (Revise Budget ถ้ามี ไม่งั้น Budget) − Actual − คำขอที่รออนุมัติ/อนุมัติแล้ว ·
  ส่งเกินงบได้แต่จะติดป้าย "เกินงบ" ให้ Admin ตัดสินใจ
- ชีตในฐานข้อมูล: `members`, `requests`, `request_items`, `budget_master`; คำขอที่อนุมัติแล้วบันทึกลง
  `pr_po_log` แยกตาม PC (ใช้ต่อใน `Budget.js`)

### เปิดใช้บน Apps Script

ฐานข้อมูล: [Google Sheet](https://docs.google.com/spreadsheets/d/1V91agWVoocDbJ54KLxAN27btFaYvEewR10NHsA3TVSY/edit)
(ตั้งไว้แล้วใน `src/Config.js`) — ระบบสร้างชีตของตัวเอง (`raw_*`, `reorder_queue`, `activity_log` …) ไม่ทับชีตเดิม

**แบบ A — คัดลอกวาง (ไม่ต้องติดตั้งอะไร)**: รัน `tools/build-appsscript.ps1` แล้วเอา 4 ไฟล์ใน
`build/apps-script/` ไปวางใน Sheet > Extensions > Apps Script
(`Code.gs`, `Index.html`, `confirm.html`, และ `appsscript.json` — เปิดให้เห็นก่อนที่ Project Settings >
"Show appsscript.json manifest file in editor")

**แบบ B — clasp** (ต้องมี Node.js):

1. ทำ Setup ข้อ 1–3 ด้านล่าง (clasp login, สร้างโปรเจกต์, กรอก `SPREADSHEET_ID`)
2. Script Properties: เพิ่ม `ANTHROPIC_API_KEY` ถ้าจะใช้แท็บผู้ช่วย AI
3. `npm run push` (build + clasp push) แล้ว Deploy > New deployment > Web app
   (Execute as: Me, Who has access: Anyone within planbmedia) — ได้ลิงก์ `/exec`
4. ลิงก์ `/exec` เปล่าๆ = หน้าเทียบยอดสั่งซื้อ · ลิงก์ที่มี `?pc=...` (จากอีเมล) = หน้ายืนยันเดิม

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
  WebApp.js       doGet: ?pc=... → หน้ายืนยัน/แก้จำนวน, ไม่มี → หน้าเทียบยอด (Ui.js)
  confirm.html    หน้า UI ของ WebApp.js
  Ui.js           เสิร์ฟหน้าเทียบยอดสั่งซื้อ + askClaudeFromUi
  DataFiles.js    ข้อมูลกลาง 4 ช่อง (ชีต data_*) อ่าน/เพิ่ม/เปลี่ยน/นำออก
  Members.js      สมาชิก Admin/User
  Requests.js     workflow คำขอ: ส่ง PC → ยืนยัน → ตรวจงบ/เกินงบ → PR/PO → รับของ + Budget + เตือนซ้ำ
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
