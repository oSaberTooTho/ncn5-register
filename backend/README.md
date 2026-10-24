# อัปเดต 22 ต.ค. 2569 — งานครั้งที่ 5 + หน้าตาใหม่

- **งาน:** การประชุมวิชาการระดับชาติ ครั้งที่ 5 สาขาพยาบาลศาสตร์ · "พลังพยาบาล พลังนวัตกรรม ร่วมสร้างอนาคตสุขภาพที่ยั่งยืน" · **27–28 พฤษภาคม 2570** · Hybrid (คณะพยาบาลศาสตร์ มรภ.นครปฐม + Zoom)
- วันที่ในหน้าเว็บ อีเมล ชื่อคอลัมน์ชีท (อาหาร_27พค / อาหาร_28พค / รับ_27พค / รับ_28พค) เปลี่ยนครบ · REG ID ยังขึ้นต้น `INHI26` (รอตัดสินใจ)
- **หน้าเว็บใหม่:** banner artwork เป็น header (ฝังในไฟล์แล้ว ไม่ต้องอัปโหลด) · โทนงาช้าง/ครีม + ทอง + เขียวตามโลโก้ · ฟอนต์ Prompt/Sarabun · รองรับมือถือ
- อีเมลยืนยันเปลี่ยนเป็นโทนครีม-ทอง-เขียว
- อีเมลยืนยัน/อนุมัติแนบ PDF A4 (ใบยืนยัน + คูปองเมื่อชำระแล้ว/ฟรี) สร้างตอนรันคิว; สำเนาเก็บใน Drive โฟลเดอร์ระบบ `PDF/<REG_ID>.pdf`
- ปุ่ม "Google Sheet" ในหน้า staff ลิงก์ไป Sheet ของระบบนั้นเองอัตโนมัติ (เดิมฝังลิงก์ Sheet ปีก่อน)
- ทดสอบ: backend mock 218/218 ผ่าน · ดูหน้าเว็บด้วย mock ทั้ง desktop/มือถือ (ลงทะเบียน/สถานะ/staff)
- เปลี่ยนรูป banner ภายหลัง: แทนไฟล์ใน `assets/` แล้วรัน `python3 -I work/build.py work/index.src.html fixed/index.html`

---

# INHI — ติดตั้งระบบใหม่บนบัญชี it.nurse (Sheet ใหม่ ไม่มีข้อมูลปีก่อน)

criteria ทั้งหมด (รายการแข่งขัน ค่าธรรมเนียม 1/2 วัน จำนวนคนต่อทีม วันรับอาหาร ประเภทกรรมการ) อยู่ใน `COMP` / `TYPE_LBL` ในโค้ด → ติดมาเอง ข้อมูลผู้ลงทะเบียนปีก่อนไม่ถูกดึงมา

## ขั้นตอน (login ด้วย it.nurse)
1. สร้าง **Google Sheet เปล่า** (ตั้งชื่ออะไรก็ได้)
2. ในชีท: **ส่วนขยาย → Apps Script**
3. ไฟล์ `Code.gs` → ลบของเดิม วางเนื้อหา `Code.gs` ทั้งหมด
   - **วางไฟล์ `Submissions.gs` เป็นไฟล์ที่ 2** (+ → Script → ตั้งชื่อ `Submissions` → วางเนื้อหาทั้งหมด) — ฟีเจอร์ "ส่งผลงานเข้าประกวด" ในหน้าเว็บ (ขั้นตอน 03) แทน Google Form เดิม · ต้องวางก่อนรัน `setupNewInstance` (จะสร้างชีท `SUBMISSIONS` ให้) · ถ้าติดตั้งไปแล้ว ให้รัน `setupSheets` อีกครั้งหลังวาง
   - **วางไฟล์ `Print.gs` เป็นไฟล์ที่ 3** (ชื่อ `Print`) — หน้าพิมพ์ A4 (ใบยืนยัน + คูปอง) ที่ปุ่มในอีเมลลิงก์ไป · ต้องวางก่อนรัน `setupNewInstance` (สร้าง `PRINT_SECRET` ให้)
   - ไฟล์ผลงานเก็บใน Drive `…/SUBMISSIONS/<ชีทรายการ>/` (ไม่แชร์สาธารณะ ทีมงานเปิดจากบัญชีเจ้าของ) · ส่งซ้ำ = เวอร์ชันใหม่แทนที่เวอร์ชันก่อน (v1 ถูกทำเครื่องหมาย "แทนที่แล้ว")
4. กด **+ → HTML** ตั้งชื่อ **`Index`** (ตัว I ใหญ่) → วางเนื้อหา `index.html` ทั้งหมด → บันทึก
5. เลือกฟังก์ชัน **`setupNewInstance`** → ▶ Run → อนุญาตสิทธิ์ (Sheets/Drive/Gmail)
   - ครั้งแรกจะขึ้น "Google hasn't verified this app" → กด **Advanced → Go to … (unsafe)** → Allow (เป็นปกติของ Apps Script ที่ใช้ Gmail)
   - `setupNewInstance` ตั้ง timezone ของ Sheet เป็น **Asia/Bangkok** ให้เอง (ถ้าอยากเช็ค: ไฟล์ → การตั้งค่า → เขตเวลา)
   - ระบบจะสร้างชีทครบ 18 ชีท + โฟลเดอร์ Drive "… — ไฟล์ระบบ" (ย้าย Sheet เข้าไป) + โฟลเดอร์ `SLIPS` และบันทึก ID ให้เอง
6. เปิด Google Sheet → แท็บ **"ตั้งค่า"** (แท็บแรก สีทอง) กรอกคอลัมน์ B:
   - `STAFF_PASSWORD` = รหัส staff ใหม่ (**ห้ามใช้รหัสปีก่อน** — หลุดแล้ว) — ว่าง = staff เข้าระบบไม่ได้
   - `PROMPTPAY_ID` = เลข PromptPay สำหรับ QR รับชำระ
   - `SUBMIT_DEADLINE` = วันหมดเขตส่งผลงาน `YYYY-MM-DD` เช่น `2027-04-30` (ว่าง = เปิดรับตลอด)
   - `WEB_APP_URL` = URL `/exec` จากหน้า Deploy (แนะนำให้กรอก) — บัญชีองค์กรจะได้ลิงก์แบบ `/a/webmail.npru.ac.th/` ซึ่งคนนอกเปิดไม่ได้ ระบบตัดให้เองแต่กรอกไว้ชัวร์กว่า
   - มีผลภายใน 1 นาที ไม่ต้อง Deploy ใหม่ · เปลี่ยนรหัส staff ก็แก้ตรงนี้ · (Script Properties ยังใช้เป็นค่าสำรองได้ถ้าเคยตั้งไว้)
7. **Deploy → New deployment → Web app** · Execute as: **Me** · Who has access: **Anyone** → Deploy
   - ถ้าไม่มีตัวเลือก Anyone (บัญชี Workspace ที่ admin ปิดไว้) → ติดต่อ admin โดเมน; ผู้ร่วมงานนอกโดเมนจะเจอหน้า login ของ Google
   - URL ของหน้าเว็บใส่ให้อัตโนมัติ ไม่ต้องแก้ใน index.html แล้ว
8. **ชื่อไฟล์ HTML ต้องเป็น `Index` เป๊ะ** · เปิด URL `/exec` ใน incognito ลองลงทะเบียน 1 รายการ + แนบสลิป 1 ครั้ง แล้วเปิดลิงก์สลิปจากหน้า staff (ถ้า Drive บล็อก "Anyone" ระบบจะ fallback เป็นแชร์ในโดเมน) · อีเมล: Gmail ทั่วไปส่งได้ ~100 ฉบับ/วัน, Workspace ~1,500 → ถ้าเต็ม ลงทะเบียนสำเร็จแต่ไม่มีอีเมล ใช้ `resendApprovedEmails` / `resendOneEmail` ส่งซ้ำวันถัดไป · dialog ของเมนู admin จะโผล่ในแท็บ Google Sheet ไม่ใช่แท็บ Apps Script · ทดสอบ: ลงทะเบียน 1 รายการ → ดูว่าเข้าชีท REGISTRATIONS + ชีทรายการ แล้วลบแถวทดสอบ (ใช้ `deleteRowsByRegId`)
9. แก้โค้ดครั้งต่อไป: Deploy → Manage deployments → ✏️ → Version: New version (URL เดิม)

ทดสอบด้วย mock (ติดตั้งใหม่ → ลงทะเบียน 4 แบบ → สลิป → อนุมัติ → เช็คอินอาหาร → รายงาน): **218/218 ผ่าน** (`work/harness/`, `VARIANT=orig node run.js`)

---

# INHI 2026 — รายงานแก้บั๊ก (20 ต.ค. 2569)

ไฟล์ส่งมอบ: `Code.gs` (+ `.md`) และ `index.html` (+ `.md`) — แทนที่ทั้งไฟล์ใน Apps Script
ผลตรวจ: `node --check` ผ่านทั้ง 2 ไฟล์ · ไม่มีฟังก์ชันซ้ำ · fn ที่หน้าเว็บเรียกมีใน doPost ครบ · **ยังไม่ได้รันกับ Sheet จริง**

## ⚠️ ต้องทำตอน deploy (ต้องทำคู่กัน)
1. วาง **Code.gs และ index.html พร้อมกัน** — ระบบ login staff เปลี่ยน contract (`getStaffPassword` ถูกลบ → `staffLogin` + ส่ง `staffPw` ทุกครั้ง) ถ้าวางไฟล์เดียว staff จะเข้าไม่ได้
2. **เปลี่ยนรหัส staff** (รหัสเดิมหลุดสู่สาธารณะแล้ว ใครก็ดึงได้จาก endpoint เก่า): Project Settings → Script Properties → เพิ่ม `STAFF_PASSWORD` = รหัสใหม่ แล้วลบค่าเดิมในโค้ดออก
3. Deploy → Manage deployments → Edit → New version (URL เดิม)
4. แก้ `ADMIN_EMAIL`, `PROMPTPAY_ID` ที่ยังเป็นค่า placeholder

## ❓ รอเจ้าของตัดสินใจ (ยังไม่แก้)
- **C2 สมัครฟรีได้โดยเลือกประเภท "กรรมการ"** — server เชื่อค่าจากหน้าเว็บ ใครก็เลือก committee แล้วไม่ต้องจ่าย + ได้อาหาร 2 วัน → เสนอ: ให้สถานะเป็น "รอตรวจสอบสิทธิ์" แล้ว staff อนุมัติ หรือเช็คจากรายชื่อกรรมการ
- `updateSheets` / `setupSheets`: ตอนนี้ข้ามชีทแข่งขันที่มีข้อมูลแล้ว (กัน title แถว 1 โดนทับ)
- `approveSpecificRegIds` แบบ force ไม่สร้างแถว PAYMENTS (เรื่องข้อมูลการเงิน)
- ชีทแข่งขันเดิม: number format คอลัมน์วันที่/ค่าธรรมเนียมเดิมผิดคอลัมน์ ต้องแก้ format ครั้งเดียว
- ไม่มี rate-limit การเดารหัส staff (ถ้าตั้งรหัสยาวพอถือว่าพอ)

## สรุปบั๊กที่แก้ (critical/high)
- รหัส staff ส่งไปให้ใครก็ได้ + endpoint staff (อนุมัติ/ปฏิเสธ/ดูสลิป/สรุปอาหารที่มีเลขบัตร ปชช./เช็คอิน) ไม่มีการตรวจสิทธิ์ฝั่ง server → ตรวจทุกครั้งแล้ว
- ค้นสถานะด้วยชื่อบางส่วน (1 ตัวอักษร) ได้ข้อมูลทุกคน+สลิป → ต้อง ≥3 ตัว, หลายคนแสดงแค่ ชื่อ/สังกัด/สถานะ
- ส่งสลิปทับ REG ID ที่ชำระแล้ว / REG ID ปลอม → ตรวจก่อน + จำกัดชนิด/ขนาดไฟล์
- XSS ในหน้า staff/สถานะ + HTML injection ในอีเมล + formula injection ในชีท → escape ทั้งหมด
- REG ID ซ้ำเมื่อลบแถว (นับจำนวนแถว) → ใช้เลขสูงสุด+1
- อนุมัติซ้ำ/กดเบิ้ล = ส่งอีเมลซ้ำ; ปฏิเสธทำให้คนที่จ่ายแล้วกลับเป็นรอ → lock + idempotent
- setupSheets รันซ้ำแล้ว header ชีทแข่งขันเลื่อน → อนุมัติไปเขียนทับ "ใบประกอบ 3"
- resetAndSetupSheets ลบทุกชีทด้วยกด OK ครั้งเดียว → ต้องพิมพ์ RESET + backup อัตโนมัติ
- หน้าเช็คอินอาหารวนเรียก server ไม่หยุดเมื่อไม่มีข้อมูล; timeout 25 วิ ทำให้ลงทะเบียนซ้ำ
- เลขโทรศัพท์/ใบประกอบ 0 นำหน้าหาย; ผู้ลงทะเบียน Online หน้าเว็บโชว์ว่าได้อาหารแต่จริงไม่ได้

## รายละเอียดทั้งหมดต่อ finding

# CHANGES part1.gs

- C1 doPost / _isStaff: removed `getStaffPassword`; added `staffLogin` (returns {success:true} or {success:false,auth:false,...}); STAFF_FNS (getPendingSlips, approvePayment, rejectPayment, getMealSummary, mealCheckIn, getMealCheckIn) gated by `_isStaff(p)` (p.staffPw vs Script Property STAFF_PASSWORD, fallback constant; constant kept + comment). reject now also receives p.approvedBy.
- H1 _getStatusData: name substring only if query >= 3 chars; multi-match returns only {rid,name,org,payStatus}, cap 20 (+`total`); single match keeps full object + payment. Also fixes M5 (payment read only once, for the single match).
- H2/L3 _submitSlip: rejects unknown REG_ID and already-paid ('ชำระแล้ว'); mime whitelist (jpeg/png/webp/pdf via p.fileType), ~6MB decoded cap, extension derived from mime (not client filename), regId sanitized in filename; append + status update under script lock with re-check.
- H3/M8 `_safe()` added (strings starting = + - @ tab/CR, or ^0\d+$ -> apostrophe prefix); applied in _registerUser (REGISTRATIONS, COMPETITIONS), _writeMealRows, _writeCompSheets, _submitSlip, _mealCheckIn text fields.
- H4 `_esc()` added; used for name/org/type/regId/r.nm/c.nm in _sendConfirmEmail and _sendPaymentConfirmEmail HTML (plain-text bodies unchanged).
- M2 _registerUser `_sheet` fallback header: inserted 'เลขใบประกอบวิชาชีพ' (now 21 cols, matches row + setupSheets).
- M3 _rejectPayment: writes G/H/I = ['ปฏิเสธ', by, date].
- M4 _approvePayment / _rejectPayment: script lock (tryLock 15s); require pending slip ('รอตรวจสอบ'), second approve returns success:false 'อนุมัติไปแล้ว' and sends no email; reject refuses if registration already 'ชำระแล้ว'; emails sent after lock release; reject now returns emailSent and logs email errors.
- M7 _writeCompSheets/_compSheet: teamNo = max(lastRow-1,1); date format moved to col 36 (was 15), '#,##0' to fee col 10 (was 9); ประชุม_Online new sheets get own 8-col header (COMP_ONLINE_HDR via optional 5th arg of _compSheet; backward compatible).
- M9 _mealCheckIn: script lock; day must be 17/18; count integer >= 1 and <= registration quota (d17/d18 from REGISTRATIONS); requires registration exists and quota >= 1.
- L1 _genRegId: max numeric suffix over all REG_IDs + 1 (global running number, as the original count-based scheme was; MMdd is cosmetic).
- L2 `_clearMealCache()` after register / approve / reject / submitSlip (check-in already had it).
- M6 _registerUser: COMPETITIONS and MEAL_SUMMARY written with one setValues each (`_appendRows`, formulas as "=" strings, MEAL_SUMMARY column order unchanged, 13 cols); confirmation email + _scheduleAutoExports now run after lock release.
- New helpers: _isStaff, _safe, _esc, _appendRows, _getRegInfo.

## Skipped / notes
- C2 untouched (as instructed).
- H1 part 2 (drop `rid` from _checkDuplicate / dupWarning) skipped: frontend may depend on it; recommend removing later.
- M1 (MEAL_SUMMARY header) is in part2 (setupSheets). M7 existing comp sheets keep old number formats; one-off manual fix needed. Existing ประชุม_Online sheet keeps its old 36-col header (not touched).
- M8 only via apostrophe prefix (no sheet-wide text format).
- L4 (Gmail quota, placeholder ADMIN_EMAIL/PROMPTPAY_ID) not changed.
- Payment-status requirement for meal check-in not added (only quota/existence), to avoid blocking staff on-site.
- part2 `_approvePayment(rid,'staff-retro-approve')` call (line ~1989) is only made when a pending slip exists, so still compatible.
- Frontend must send `staffPw` in every staff call and use `staffLogin`.

# CHANGES part2.gs
- F1 setupSheets: competition sheets use COMP_SHEET_HDR (36 col); makeSheet(...,{onlyIfEmpty:true}) skips any existing non-empty sheet (real sheets have title row 1 + header row 2, so row-1 rewrite would clobber the title).
- F2/M1 setupSheets+resetAndSetupSheets: MEAL_SUMMARY header 13 cols (การแพ้อาหาร L, วันที่บันทึก M), widths 13. setupSheets on existing data: if M header empty and L2 is not a Date, rewrite only L1:M1; else skip and log.
- F3 setupSheets+reset: PAYMENTS col E = 'วันที่โอน' (matches updateSheets/_submitSlip), width 140.
- F4 resetAndSetupSheets: ui.prompt requiring RESET, shows per-sheet row counts + total; DriveApp makeCopy backup 'BACKUP before reset <ts>' before deletion.
- F5 fixMealSummaryFormulas: batch setFormulas for I/J/K (COUNTIFS formulas restored). _updateMealSummaryCheckin: does not overwrite formula cells.
- F6 _scheduleAutoExports: LockService, trigger created before flag, flag = timestamp (stale >3 min rescheduled), stores trigger uid. _runAutoExportsNow: clears flag at start; finally deletes handler triggers except freshly queued one.
- F7 approveSpecificRegIds: IDs not in REGISTRATIONS skipped + logged (no PAYMENTS row added - left as is).
- F8 exportFullRegistrationSummary: committee rows (non-rejected) count as paid, status shown 'ฟรี'. Column-mapping point not changed.
- F10 resendApprovedEmails: MailApp quota check/confirm, stops at 4.5 min or quota, reports where stopped.
- F11 fixOnlineSheetData: refuses 8-col sheet; correct date/saved columns.
- F12 fixPaymentRegIds2: normalized full-name (prefix+first+last / fullname) map; ambiguous skipped and logged; card match only when no name candidates.
- F13 trimEmptyFormattedRows: detects header row (rows 1-3), deletes only rows with no value in any column. (debugCompSheetsStatus lastRow-2 left: sheets have title+header rows.)
- F14 exportFullRegistrationSummary: c00 adds cfg.d1 to feeTotal/feeCollected.
- F16 fixDoubleHeader: row 2 treated as old header only if a cell === 'REG_ID'.
- Duplicate testUpdatePayStatus (first copy) deleted; surviving copy validates REG_ID exists + ui confirm. testPaymentConfirmEmail sends to Session.getActiveUser().getEmail().
- Skipped: F9 (covered above), F15 (fixed 1000/2000-row format windows; low value, risk), F8 judge column remap, F7 PAYMENTS row insertion, F10 sent-tracking (per instruction), updateSheets (not in findings) still rewrites row 1 of comp sheets with HDR34 (same title-row clobber risk) - needs owner decision.

# CHANGES part3.gs (review-gs-admin2)
- ADM2-01 generateSafeRegId_: uniqueness checked vs REGISTRATIONS+PAYMENTS+MEAL_SUMMARY+MEAL_CHECKIN, 500 tries, throws instead of returning a colliding ID. (_genRegId itself is part1.)
- ADM2-02 moveTeamToCompetition: missing dest sheet created via _compSheet (inside lock, before any data write).
- ADM2-03 moveTeamToCompetition: script lock; all validation first (dest header, dup, COMPETITIONS/REGISTRATIONS headers); source row re-located by REG_ID right before use; order append dest -> delete src -> COMPETITIONS -> MEAL -> recalc; try/catch alerts completed steps. Auto-export scheduled after lock release.
- ADM2-04 moveTeamToCompetition: abort if REG_ID already in dest; ลำดับทีม = destSh.getLastRow() (same as _writeCompSheets). Combined/split member header mismatch (4c) NOT handled (noted).
- ADM2-05 _recalcRegTotals: returns (logs) if no COMPETITIONS rows; throws on missing REGISTRATIONS headers; logs unknown cid.
- ADM2-06 fixDuplicateRegId: MEAL_SUMMARY and MEAL_CHECKIN matched by REG_ID + full name (card only if card cell non-empty).
- ADM2-07 fixDuplicateRegId: COMPETITIONS match restricted to cids matched by card in comp sheets, full name (not first name), ambiguous (same cid >1 row) skipped, header guard.
- ADM2-08 lock + re-verify REG_ID before write/delete in fixDuplicateRegId, deleteRowsByRegId, compactCompSheets (compact also aborts a sheet if row count changed since preview).
- ADM2-09 deleteRowsByRegId: removed rows copied to DELETED_ROWS_LOG; sync exports replaced with _scheduleAutoExports. (9c other derived sheets not refreshed - not changed.)
- ADM2-10 compactCompSheets: formulas preserved (getFormulasR1C1/setFormulaR1C1); sheets with rows having data but blank REG_ID skipped.
- ADM2-11 exportJudgesByRoom: breakApart before clear; single setValues/setBackgrounds; merges only on title rows.
- ADM2-12 exportJudgesByRoom: registrant/judge normalized name must be >=2 chars; contains-fallback needs both >=4.
- ADM2-13 exportReceiptsByAddress: empty later PAYMENTS row no longer overwrites earlier addr/receipt; shading batched via setBackgrounds.
- ADM2-15 a/b: diagnoseRegistrationsSheet header-only guard; _colLetter_ helper replaces fromCharCode(65+i). (15c showDuplicateRegDetails perf: not changed.)
- ADM2-17: exportLicensedPersonsList cols C:D set to text format. Dedupe NOT applied (changes counts; business decision).
- ADM2-18: runAllSummaryReports per-export try/catch with ok/fail report.
- Skipped: ADM2-14 (override key normalization; touches multiple functions, risk of merging wrong people), ADM2-16 (part1/2).

# CHANGES-html (fixed/index.html)
- Auth: getStaffPassword removed; staffLogin({staffPw}); staffPw kept in module var + sessionStorage (try/catch); staffCall() wrapper adds staffPw to getPendingSlips/approvePayment/rejectPayment/getMealSummary/mealCheckIn/getMealCheckIn; auth===false -> staffLogout(message). Logout button added; session restored on opening Staff tab. Meal check-in UI lives inside staff-dash, so it was already behind login (now server-gated too).
- F02: esc()/safeUrl() added; all server/user data in innerHTML + value attrs escaped; inline onclick string interpolation replaced by data-* + this.dataset; slipUrl must start with https://.
- F03: ciLoaded flag breaks ciSearch<->ciLoadAll loop (also set in catch); refreshed on entering check-in tab.
- F04: registerUser/submitPaymentSlip timeout 90s; timeout/cancel message tells user to check status before resubmitting.
- F05/F06: overlay cancel -> cancelOv() aborts in-flight calls (+opSeq ignores late FileReader result); slip submit disabled while busy, re-enabled on all outcomes; slipFile reset on rejected pick; type (jpg/png/webp/pdf) + 5MB check; reader.onerror; non-JSON response -> clear Thai message.
- F07: c00 meal 'none' (no meals, badge "ไม่มีอาหาร (Online)").
- F08: backend _mealCheckIn reads no giver field; field made optional, sent as `giver` (ignored by backend until a column is added).
- F10/F11: Bangkok date/time (Intl) for slip default; "today" KPI parses dd/MM/yyyy; KPI number has .kpi-v, approve/reject decrement only it.
- F12: checkDuplicate failure no longer blocks. F13: GAS const uses https:\/\/. F15: loadStaffData shows error on success:false; submit-btn re-enabled after reset; leaving step 3 hides conf-page and restores s3-btn-row.
- Multi-match list renders rid/name/org/payStatus only; click re-queries by rid.
- Skipped: dead members block in status page (left, escaped); #addr shared id (left); Google Sheet link still in staff-dash HTML; pt-step tracker not reset after a free registration (cosmetic).

- อีเมลยืนยันลงทะเบียน/อนุมัติ/ปฏิเสธส่งโดยคิวเบื้องหลัง (ชีทซ่อน EMAIL_QUEUE) ภายใน ~1 นาที; ถ้าอีเมลไม่มา ใช้เมนู "✉️ ส่งอีเมลค้างในคิวตอนนี้"
- **พิมพ์คูปองรวม (staff):** Staff Dashboard → การ์ด "พิมพ์คูปองรวม (ตาม session)" → `getBatchPrintLink` → `Print.gs` `mode=batch` (ลิงก์หมดอายุ 2 ชม.; ต้อง deploy Code.gs + Print.gs + Index พร้อมกัน)

## โฮสต์หน้าเว็บบน GitHub Pages

เหตุผล: ผู้ใช้ที่ล็อกอินหลายบัญชี Google เปิด script.google.com แล้ว Google เลือกบัญชีผิด (/u/1/) และขึ้นหน้า Drive error — `fetch()` ไปที่ `/exec` ไม่มีคุกกี้ จึงไม่โดน ดังนั้นเสิร์ฟหน้าเว็บจาก GitHub Pages แล้วให้ Apps Script เป็น backend อย่างเดียว

1. Deploy `Code.gs` + `Print.gs` ใหม่ (New version) — deployment ยังเป็น **Execute as: Me / Who has access: Anyone** เหมือนเดิม; URL `/exec` ใส่ไว้ใน `work/gas-url.txt`
2. Build: `python3 -I work/build.py work/index.src.html site/index.html --gas auto` → ได้ `site/index.html`, `site/print.html`, `site/.nojekyll` (ใช้ `--gas <url>` เพื่อระบุ URL เอง)
3. Push โฟลเดอร์ `site/` ไป repo `ncn5-register` (branch main, Settings → Pages → root) → https://osabertootho.github.io/ncn5-register/
4. ในชีท "ตั้งค่า" ใส่ `SITE_URL` = `https://osabertootho.github.io/ncn5-register` (ไม่มี / ท้าย) — ลิงก์พิมพ์ในอีเมลจะเป็น `SITE_URL/print.html?rid=..&k=..` และปุ่ม "ตรวจสอบสถานะ" เป็น `SITE_URL/#st`; ลิงก์พิมพ์รวมของ staff ก็ชี้ไปหน้านี้ (เว้นว่าง = ใช้ลิงก์ `/exec?p=print` แบบเดิม)
5. Backend ใหม่: `getPrintHtml` {rid,k,em} · `getBatchHtml` {day,cid,paid,exp,k} (ตรวจสิทธิ์เหมือน `?p=print`) · `registerUser`/`getStatusData`(ค้นด้วย REG ID เต็ม) ส่ง `printToken` · `getBatchPrintLink` ส่ง `params`
6. `fixed/index.html` (build เดิมไม่มี `--gas`) ยังใช้กับ doGet ได้ตามปกติ
