// ================================================================
//  Code.gs (v3)  —  Full Backend v3
//  NEW: QR Code links · PDF receipt trigger · PromptPay QR
//       Meal logic (27/28 May, quota-based) · Bilingual email
//       Duplicate check · Staff dashboard + meal summary
//  ★ แก้ค่าใน SECTION 1 ก่อน Deploy ★
// ================================================================

// ── SECTION 1: CONFIG ──────────────────────────────────────────
// ค่าที่ผูกกับบัญชี/ไฟล์ เก็บใน Script Properties (Project Settings → Script Properties)
//   SS_ID, DRIVE_FOLDER_ID  ← setupNewInstance() ตั้งให้เอง
//   STAFF_PASSWORD          ← ต้องตั้งเอง (ไม่มีค่า = staff เข้าระบบไม่ได้)
//   PROMPTPAY_ID            ← ตั้งเอง (ADMIN_EMAIL ยังไม่มีโค้ดส่วนไหนใช้)
var _PROPS          = PropertiesService.getScriptProperties();
// SS_ID: ใช้ Sheet ที่สคริปต์นี้ผูกอยู่ (เปิดจาก ส่วนขยาย → Apps Script) เป็นหลักเสมอ — Script Properties เป็นแค่ตัวสำรอง
// กันกรณีวางโค้ดเก่า/Properties ค้างค่า แล้วระบบไปเขียน Sheet ปีก่อน
function _boundSheetId(){ try { var a = SpreadsheetApp.getActiveSpreadsheet(); return a ? a.getId() : ''; } catch (e) { return ''; } }
var SS_ID           = _boundSheetId() || _PROPS.getProperty('SS_ID') || '';
var DRIVE_FOLDER_ID = _PROPS.getProperty('DRIVE_FOLDER_ID') || '';
var ADMIN_EMAIL     = _PROPS.getProperty('ADMIN_EMAIL') || '';
var STAFF_PASSWORD  = '';                        // ← ห้ามใส่รหัสในโค้ด — ใช้ Script Property 'STAFF_PASSWORD'

// ⚠️  Gmail permission: ครั้งแรกต้อง Run ฟังก์ชัน testEmail() ด้านล่างก่อน
//     เพื่อ authorize scope gmail.send — ไม่งั้น sendEmail จะ throw ไม่ส่ง
// PROMPTPAY_ID / STAFF_PASSWORD / SUBMIT_DEADLINE → กรอกในชีท "ตั้งค่า" (ดู _cfg)

var CONF_NAME_TH = 'การประชุมวิชาการระดับชาติ ครั้งที่ 5 สาขาพยาบาลศาสตร์ “พลังพยาบาล พลังนวัตกรรม ร่วมสร้างอนาคตสุขภาพที่ยั่งยืน”';
var CONF_NAME_EN = 'Nursing and Innovation: Co-creating a Sustainable Health Future';
var CONF_DATE_TH = '27–28 พฤษภาคม พ.ศ. 2570';
var CONF_DATE_EN = '27–28 May 2027 (B.E. 2570)';
var CONF_VENUE_TH= 'แบบ Hybrid ณ คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏนครปฐม และออนไลน์ผ่าน Zoom Application';
var CONF_VENUE_EN= 'Hybrid: Faculty of Nursing, Nakhon Pathom Rajabhat University and online via Zoom';
var REG_PREFIX   = 'INHI26';
var MAIL_SENDER_NAME = 'ประชุมวิชาการพยาบาล ครั้งที่ 5 · มรภ.นครปฐม'; // ชื่อผู้ส่งที่โชว์ในกล่องจดหมาย (ส่งจากบัญชีที่ Deploy)

// Sheet names — หลัก
var SH_REG     = 'REGISTRATIONS';
var SH_COMP    = 'COMPETITIONS';
var SH_PAY     = 'PAYMENTS';
var SH_MEAL    = 'MEAL_SUMMARY';
var SH_CHECKIN = 'MEAL_CHECKIN';   // ← sheet บันทึกการรับอาหาร

// Sheet names — แยกรายการแข่งขัน (สร้างอัตโนมัติเมื่อมีผู้ลงทะเบียน)
var COMP_SHEET = {
  c00:  'ประชุม_Online',
  c00s: 'ประชุม_Onsite',
  c02:  'SIM_MAN',
  c03:  'SIM_MOM',
  c04:  'SIM_PED',
  c05:  'แนวปฏิบัติ_สถาบัน',
  c06:  'แนวปฏิบัติ_สถานบริการ',
  c07a: 'บทความวิจัย_บุคลากร',
  c07b: 'บทความวิจัย_นศ',
  c08:  'ตอบคำถาม_พยาบาล',
  c09:  'ตอบคำถาม_กายวิภาค',
  c10:  'นวัตกรรม_นศ',
  c11:  'นวัตกรรม_อาจารย์'
};

// Header สำหรับชีทรายการแข่งขัน
var COMP_SHEET_HDR = [
  'ลำดับทีม','REG_ID','วันที่ลงทะเบียน',
  'ชื่อทีม',
  'ชื่อหัวหน้าทีม','สังกัด','อีเมล','โทรศัพท์',
  'จำนวนวัน','ค่าธรรมเนียม (บาท)',
  'คำนำหน้า 1','ชื่อ 1','นามสกุล 1','เลขบัตร 1','ใบประกอบ 1','แพ้อาหาร 1',
  'คำนำหน้า 2','ชื่อ 2','นามสกุล 2','เลขบัตร 2','ใบประกอบ 2','แพ้อาหาร 2',
  'คำนำหน้า 3','ชื่อ 3','นามสกุล 3','เลขบัตร 3','ใบประกอบ 3','แพ้อาหาร 3',
  'คำนำหน้า 4','ชื่อ 4','นามสกุล 4','เลขบัตร 4','ใบประกอบ 4','แพ้อาหาร 4',
  'สถานะชำระ','วันที่บันทึก'
];

// ── COMP CONFIG ──────────────────────────────────────────────
//   mealDay: 'both'|'d17'|'d18'
//   d1 = 1-day fee, d2 = 2-day fee (0 = no 2-day option)
var COMP = {
  c00:  { nm:'ลงทะเบียนประชุมวิชาการ Online',              nm_en:'Conference Registration (Online)',        d1:700,  d2:0,    mealDay:'none', team:1, date:'27–28 พ.ค. 2570', time:'', venue:'ออนไลน์ผ่าน Zoom Application' },
  c00s: { nm:'ลงทะเบียนประชุมวิชาการ Onsite',             nm_en:'Conference Registration (Onsite)',       d1:1200, d2:0,    mealDay:'both', team:1, date:'27–28 พ.ค. 2570', time:'', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
  c02:  { nm:'Skill Contest: SIM MAN',                      nm_en:'Skill Contest: SIM MAN',                d1:1200, d2:2000, mealDay:'d18',  team:4, date:'28 พ.ค. 2570', time:'08.30–15.30 น.', venue:'คณะพยาบาลศาสตร์ ม.ราชภัฏนครปฐม' },
  c03:  { nm:'Simulation: SIM MOM',                         nm_en:'Simulation: SIM MOM',                   d1:1200, d2:2000, mealDay:'d18',  team:4, date:'28 พ.ค. 2570', time:'08.30–12.00 น.', venue:'คณะพยาบาลศาสตร์ ม.ราชภัฏนครปฐม' },
  c04:  { nm:'Simulation: SIM PED',                         nm_en:'Simulation: SIM PED',                   d1:1200, d2:2000, mealDay:'d18',  team:4, date:'28 พ.ค. 2570', time:'08.30–12.00 น.', venue:'คณะพยาบาลศาสตร์ ม.ราชภัฏนครปฐม' },
  c05:  { nm:'แนวปฏิบัติที่ดี (สถาบันการศึกษา)',          nm_en:'Best Practice — Academic (Hybrid)',      d1:1200, d2:2000, mealDay:'d17',  team:2, date:'27 พ.ค. 2570', time:'13.00–15.00 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม (Hybrid)' },
  c06:  { nm:'แนวปฏิบัติที่ดี (สถานบริการ)',              nm_en:'Best Practice — Hospital (Hybrid)',      d1:1200, d2:2000, mealDay:'d17',  team:2, date:'27 พ.ค. 2570', time:'13.00–15.00 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม (Hybrid)' },
  c07a: { nm:'นำเสนอบทความวิจัย (บุคลากรสาธารณสุข)',      nm_en:'Research Paper — Health Personnel',      d1:1200, d2:2000, mealDay:'d17',  team:2, date:'27 พ.ค. 2570', time:'13.00–17.00 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
  c07b: { nm:'นำเสนอบทความวิจัย (นักศึกษา)',                 nm_en:'Research Paper — Student',                d1:1200, d2:2000, mealDay:'d17',  team:2, date:'27 พ.ค. 2570', time:'13.00–17.00 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
 c08: { nm:'การแข่งขันตอบคำถามทางการพยาบาล',   nm_en:'Nursing Quiz Competition', d1:1200, d2:2000, mealDay:'d18', team:3, date:'28 พ.ค. 2570', time:'08.30–10.30 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
c09: { nm:'การแข่งขันตอบคำถามกายวิภาคศาสตร์', nm_en:'Anatomy Quiz Competition', d1:1200, d2:2000, mealDay:'d18', team:3, date:'28 พ.ค. 2570', time:'08.30–10.30 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
 c10: { nm:'นวัตกรรม (ห้องนักศึกษา)',                   nm_en:'Innovation — Student Track',       d1:1200, d2:2000, mealDay:'d18', team:2, date:'28 พ.ค. 2570', time:'08.30–10.30 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
c11: { nm:'นวัตกรรม (ห้องอาจารย์/บุคลากรสาธารณสุข)', nm_en:'Innovation — Faculty/Staff Track', d1:1200, d2:2000, mealDay:'d18', team:2, date:'28 พ.ค. 2570', time:'08.30–10.30 น.', venue:'อาคารสิริวรปัญญา ม.ราชภัฏนครปฐม' },
};

var TYPE_LBL    = { committee_dir:'กรรมการอำนวยการ', committee_judge:'กรรมการตัดสินการแข่งขัน', other:'ผู้เข้าร่วมทั่วไป' };
var TYPE_LBL_EN = { committee_dir:'Program Committee', committee_judge:'Competition Judge', other:'General Participant' };

// ── SECTION 2: WEB APP ENTRY POINTS ──────────────────────────

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.p === 'print') return _printServe(e);   // Print.gs — หน้าพิมพ์ A4
    // ไม่ใช้ template (กัน gotcha template ตัด https://) — แทน __GAS_URL__ ด้วย URL ของ deployment นี้
    var url  = _webAppUrl() || '';
    var html = HtmlService.createHtmlOutputFromFile('Index').getContent()
                 .replace('__GAS_URL__', url.replace(/\//g, '\\/'));
    var out = HtmlService.createHtmlOutput(html);
    out.setTitle(CONF_NAME_EN)
       .addMetaTag('viewport','width=device-width,initial-scale=1')
       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return out;
  } catch(err) {
    return HtmlService.createHtmlOutput('<h2 style="color:red">Error: '+err.message+'</h2>');
  }
}

function doPost(e) {
  var out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try {
    // ต้องผ่าน setupNewInstance กับ Sheet นี้แล้วเท่านั้น (setup บันทึก SS_ID ของ Sheet ที่ติดตั้งไว้ใน Properties)
    if (!SS_ID || _PROPS.getProperty('SS_ID') !== SS_ID) throw new Error('ระบบยังไม่ได้ตั้งค่า — ผู้ดูแลต้องรัน setupNewInstance() กับ Sheet นี้ก่อน');
    var p  = JSON.parse(e.postData.contents);
    var fn = (e.parameter && e.parameter.fn) || '';
    var result;
    var STAFF_FNS=['getPendingSlips','approvePayment','rejectPayment','getMealSummary','mealCheckIn','getMealCheckIn','getBatchPrintLink'];
    if(STAFF_FNS.indexOf(fn)>=0 && !_isStaff(p)){
      out.setContent(JSON.stringify({success:false,auth:false,message:'ต้องเข้าสู่ระบบเจ้าหน้าที่'}));
      return out;
    }
    switch(fn) {
      case 'registerUser':      result = _registerUser(p);           break;
      case 'checkDuplicate':    result = _checkDuplicate(p);         break;
      case 'getStatusData':     result = _getStatusData(p.rid||''); break;
      case 'submitPaymentSlip': result = _submitSlip(p);             break;
      case 'staffLogin':        result = _isStaff(p) ? {success:true, sheetUrl:'https://docs.google.com/spreadsheets/d/'+SS_ID+'/edit'} : {success:false,auth:false,message:'รหัสผ่านไม่ถูกต้อง'}; break;
      case 'getPendingSlips':   result = _getPendingSlips();          break;
      case 'approvePayment':    result = _approvePayment(p.regId||'', p.approvedBy||''); break;
      case 'rejectPayment':     result = _rejectPayment(p.regId||'', p.approvedBy||''); break;
      case 'getMealSummary':    result = _getMealSummary();           break;
      case 'getPromptPayInfo':  result = {success:true, ppId:_cfg('PROMPTPAY_ID'), accountName:'มหาวิทยาลัยราชภัฏนครปฐม บริการทางวิชาการ', bank:'ธนาคารกรุงไทย', account:'986-5-34173-5'}; break;
      case 'mealCheckIn':       result = _mealCheckIn(p);             break;
      case 'getBatchPrintLink': result = _getBatchPrintLink(p);       break;
      case 'getPrintHtml':      result = _getPrintHtml(p);            break;   // Print.gs — หน้าพิมพ์สำหรับโฮสต์แยก (GitHub Pages)
      case 'getBatchHtml':      result = _getBatchHtml(p);            break;
      case 'getMealCheckIn':    result = _getMealCheckIn(p);          break;
      default:
        // ฟังก์ชันจากไฟล์อื่น (เช่น Submissions.gs) ลงทะเบียนผ่าน _extraHandlers
        var ext = (typeof _extraHandlers === 'function') ? _extraHandlers() : {};
        result = ext[fn] ? ext[fn](p) : {success:false, message:'Unknown fn: '+fn};
    }
    out.setContent(JSON.stringify(result || {success:false, message:'no result'}));
  } catch(err) {
    Logger.log('doPost err: '+err.message);
    out.setContent(JSON.stringify({success:false, message:err.message}));
  }
  return out;
}

// ── HELPERS: auth / sanitize ─────────────────────────────────
// ── CONFIG จากชีท "ตั้งค่า" ─────────────────────────────────
//  เจ้าของระบบกรอกค่าในชีทได้เอง ไม่ต้องเข้า Script Properties
//  ลำดับ: ชีท "ตั้งค่า" (คอลัมน์ A=คีย์, B=ค่า) → Script Properties (สำรอง); cache 60 วินาที
var SH_CFG = 'ตั้งค่า';
var CFG_ROWS = [
  ['STAFF_PASSWORD',  '', 'รหัสผ่านเจ้าหน้าที่ (หน้า Staff) — ตั้งใหม่ อย่าใช้รหัสปีก่อน'],
  ['PROMPTPAY_ID',    '', 'เลข PromptPay สำหรับ QR รับชำระ (เบอร์โทร/เลขผู้เสียภาษี 13 หลัก)'],
  ['SUBMIT_DEADLINE', '', 'วันหมดเขตส่งผลงาน รูปแบบ YYYY-MM-DD เช่น 2027-04-30 (เว้นว่าง = เปิดรับตลอด)'],
  ['WEB_APP_URL',     '', 'URL ของเว็บแอป (ลงท้าย /exec จากหน้า Deploy) ใช้ในลิงก์ในอีเมล — เว้นว่างได้ ระบบหาให้เอง'],
  ['SITE_URL',        '', 'URL หน้าเว็บที่โฮสต์แยก เช่น https://osabertootho.github.io/ncn5-register — ถ้าตั้ง ลิงก์ในอีเมลจะชี้ไปที่นี่']
];
function _cfg(key){
  var cache = null, ck = 'cfg_' + key;
  try { cache = CacheService.getScriptCache(); var c = cache.get(ck); if (c !== null && c !== '') return c; } catch (e) {}
  var val = '';
  try {
    var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SH_CFG);
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === key) { val = String(rows[i][1] == null ? '' : rows[i][1]).trim(); break; }
      }
    }
  } catch (e) { Logger.log('_cfg sheet: ' + e.message); }
  if (!val) { try { val = String(PropertiesService.getScriptProperties().getProperty(key) || '').trim(); } catch (e) {} }
  if (val && cache) { try { cache.put(ck, val, 60); } catch (e) {} }
  return val;
}
function _ensureCfgSheet(ss, log){
  var sh = ss.getSheetByName(SH_CFG);
  if (sh) { if (log) log.push('⏭ [มีอยู่แล้ว] ' + SH_CFG); return sh; }
  sh = ss.insertSheet(SH_CFG);
  sh.getRange(1, 1, 1, 3).setValues([['คีย์ (ห้ามแก้)', 'ค่า — กรอกตรงนี้', 'คำอธิบาย']])
    .setBackground('#0b5844').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, CFG_ROWS.length, 3).setValues(CFG_ROWS);
  sh.getRange(2, 2, CFG_ROWS.length, 1).setBackground('#fff8e1').setNumberFormat('@');
  sh.getRange(2, 1, CFG_ROWS.length, 1).setFontFamily('Courier New');
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 520);
  sh.setFrozenRows(1); sh.setTabColor('#d4a64a');
  if (log) log.push('✅ [สร้างใหม่] ' + SH_CFG + ' — กรอกค่าในคอลัมน์ B');
  return sh;
}

// URL สาธารณะของ Web app สำหรับใส่ในอีเมล/หน้าเว็บ
//  บัญชีองค์กร (Workspace) ScriptApp.getService().getUrl() คืนรูปแบบ script.google.com/a/<โดเมน>/macros/... ซึ่งคนนอกโดเมนเปิดไม่ได้
//  → ตัด /a/<โดเมน>/ ออก หรือใช้ค่า WEB_APP_URL ที่กรอกในชีท "ตั้งค่า" (URL /exec จาก Deploy) ถ้ามี
function _webAppUrl(){
  var u = '';
  try { u = _cfg('WEB_APP_URL'); } catch (e) {}
  if (!/^https:\/\/script\.google\.com\//.test(u)) { try { u = ScriptApp.getService().getUrl() || ''; } catch (e) { u = ''; } }
  return String(u).replace(/^https:\/\/script\.google\.com\/a\/[^\/]+\/macros\//, 'https://script.google.com/macros/').replace(/\/dev$/, '/exec');
}
// staff: ลิงก์พิมพ์คูปองรวมตาม session (Print.gs) — หมดอายุ 2 ชม.
function _getBatchPrintLink(p){
  var day=String(p.day||''), cid=String(p.cid||'all'), paid=p.onlyPaid===false||p.onlyPaid==='false'||p.onlyPaid===0||p.onlyPaid==='0'?'0':'1';
  if(!_batchParamsOk(day,cid,paid)) return {success:false,message:'เงื่อนไขไม่ถูกต้อง'};
  var exp=String(Math.floor(Date.now()/1000)+7200);
  var count=_batchBuild(day,cid,paid==='1').items.length;
  var k=_batchSig(day,cid,paid,exp), site=_siteUrl();
  var q='mode=batch&day='+day+'&cid='+encodeURIComponent(cid)+'&paid='+paid+'&exp='+exp+'&k='+k;
  return {success:true,count:count,params:{day:day,cid:cid,paid:paid,exp:exp,k:k},
          url:site?site+'/print.html?'+q:(_webAppUrl()||'')+'?p=print&'+q};
}
// URL หน้าเว็บที่โฮสต์แยก (GitHub Pages) — ว่าง = ใช้ Web app เดิม; ตัด / ท้ายออก
function _siteUrl(){
  var u=''; try{ u=String(_cfg('SITE_URL')||'').trim(); }catch(e){}
  return /^https:\/\//.test(u) ? u.replace(/\/+$/,'') : '';
}
function _isStaff(p){
  var pw=_cfg('STAFF_PASSWORD');
  return !!pw && !!p && String(p.staffPw||'')===String(pw);
}
// กัน formula injection (= + - @) และเลขนำหน้า 0 หาย (เบอร์โทร/ใบประกอบ) — เฉพาะ string
function _safe(v){
  if(typeof v!=='string') return v;
  if(/^[=+\-@\t\r]/.test(v) || /^0\d+$/.test(v)) return "'"+v;
  return v;
}
function _esc(s){
  return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
// เขียนหลายแถวรวดเดียว (ขยายจำนวนแถวชีทถ้าไม่พอ) — คืนเลขแถวเริ่มต้น
function _appendRows(sh,rows){
  if(!rows.length) return sh.getLastRow()+1;
  var start=sh.getLastRow()+1, need=start+rows.length-1;
  if(need>sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need-sh.getMaxRows());
  sh.getRange(start,1,rows.length,rows[0].length).setValues(rows);
  return start;
}
// ข้อมูลสถานะของ REG_ID (null ถ้าไม่พบ)
function _getRegInfo(ss,regId){
  var sh=ss.getSheetByName(SH_REG);
  if(!sh||sh.getLastRow()<2) return null;
  var lc=sh.getLastColumn(), hdr=sh.getRange(1,1,1,lc).getValues()[0];
  var iS=hdr.indexOf('สถานะชำระ'); if(iS<0) iS=17;
  var i17=hdr.indexOf('อาหาร_27พค'); if(i17<0) i17=18;
  var i18=hdr.indexOf('อาหาร_28พค'); if(i18<0) i18=19;
  var iE=hdr.indexOf('อีเมล'), iN=hdr.indexOf('ชื่อเต็ม'), iPre=hdr.indexOf('คำนำหน้า'), iFn=hdr.indexOf('ชื่อ'), iLn=hdr.indexOf('นามสกุล');
  var iOrg=hdr.indexOf('สังกัด/หน่วยงาน'), iFee=hdr.indexOf('ยอดรวม'), iTy=hdr.indexOf('ประเภทผู้ลงทะเบียน');
  var data=sh.getRange(2,1,sh.getLastRow()-1,lc).getValues();
  var id=String(regId||'').trim();
  if(!id) return null;
  var g=function(row,ix){ return ix<0?'':String(row[ix]==null?'':row[ix]); };
  for(var i=0;i<data.length;i++){
    if(String(data[i][0]).trim()===id){
      var r=data[i], nm=g(r,iN).trim();
      if(!nm) nm=(g(r,iPre)+' '+g(r,iFn)+' '+g(r,iLn)).trim();
      // ฟิลด์เสริมใช้ส่งอีเมล (อ่านรอบเดียวกัน ไม่ต้องอ่านชีทซ้ำ)
      return{status:String(data[i][iS]||'').trim(),d17:Number(data[i][i17]||0),d18:Number(data[i][i18]||0),
             email:g(r,iE).trim(),name:nm,org:g(r,iOrg),fee:iFee<0?0:Number(r[iFee]||0),typeLbl:g(r,iTy)};
    }
  }
  return null;
}

// ── SECTION 3: DUPLICATE CHECK ───────────────────────────────

function _checkDuplicate(d) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SH_REG);
    if(!sh||sh.getLastRow()<2) return {success:true,duplicate:false};
    var data=sh.getDataRange().getValues(), hdr=data[0];
    var iE=hdr.indexOf('อีเมล'), iC=hdr.indexOf('เลขบัตรประชาชน');
    var em=(d.email||'').trim().toLowerCase(), ic=(d.idcard||'').trim();
    for(var i=1;i<data.length;i++){
      if(em && String(data[i][iE]||'').trim().toLowerCase()===em)
        return {success:true,duplicate:true,field:'email',rid:data[i][0]};
      if(ic && String(data[i][iC]||'').trim()===ic)
        return {success:true,duplicate:true,field:'idcard',rid:data[i][0]};
    }
    return {success:true,duplicate:false};
  } catch(e){ return {success:false,message:e.message}; }
}

// ── SECTION 4: REGISTER ──────────────────────────────────────

// กำหนดการ/สถานที่ต่อรายการ — "28 พ.ค. 2570 · 08.30–15.30 น. · คณะพยาบาลศาสตร์"
function _schedLine(cid){
  var c=COMP[cid]; if(!c) return '';
  return [c.date,c.time,c.venue].filter(function(x){return x;}).join(' · ');
}
// กำหนดการของผู้ลงทะเบียน (จากชีท COMPETITIONS) — ใช้ในอีเมลยืนยันชำระเงิน
function _schedHtmlFromSheets(regId){
  try{
    var sh=SpreadsheetApp.openById(SS_ID).getSheetByName(SH_COMP);
    if(!sh||sh.getLastRow()<2) return '';
    var v=sh.getDataRange().getValues(), rows='', id=String(regId||'').trim();
    for(var i=1;i<v.length;i++){
      if(String(v[i][0]).trim()!==id) continue;
      var cid=String(v[i][2]||'').trim(), c=COMP[cid]; if(!c) continue;
      rows+='<tr><td style="padding:6px 10px;border-top:1px solid #eee;font-size:12px">'+_esc(c.nm)+
        '<div style="font-size:11px;color:#6b7280;margin-top:2px">'+_esc(_schedLine(cid))+'</div></td></tr>';
    }
    if(!rows) return '';
    return '<p style="font-size:10px;color:#888;letter-spacing:.08em;text-transform:uppercase;margin:14px 0 6px">กำหนดการของท่าน / Your Schedule</p>'+
      '<table style="width:100%;border-collapse:collapse;margin-bottom:14px">'+rows+'</table>';
  }catch(e){ Logger.log('SCHED ERROR ['+regId+']: '+e.message); return ''; }
}

// สถาบันร่วมจัด (Co-host): โควต้ากรรมการฟรีต่อสถาบัน
var COHOST_LIMIT = { committee_dir:1, committee_judge:2 };
function _normOrg(s){ return String(s||'').replace(/\s+/g,' ').trim().toLowerCase(); }
// คืนข้อความ error ถ้าโควต้าเต็ม/ไม่ระบุสถาบัน, '' ถ้าผ่าน — เรียกภายใน lock ก่อนเขียนชีท
function _checkCohostQuota(ss,type,org){
  var lim=COHOST_LIMIT[type]; if(!lim) return '';
  var key=_normOrg(org);
  if(!key) return 'กรุณาระบุสถาบัน/หน่วยงาน';
  var sh=ss.getSheetByName(SH_REG), cnt=0;
  if(sh&&sh.getLastRow()>1){
    var v=sh.getDataRange().getValues(), h=v[0];
    var iT=h.indexOf('ประเภทผู้ลงทะเบียน'), iO=h.indexOf('สังกัด/หน่วยงาน'), iS=h.indexOf('สถานะชำระ');
    for(var i=1;i<v.length;i++){
      if(String(v[i][iT]||'')!==TYPE_LBL[type]) continue;
      if(_normOrg(v[i][iO])!==key) continue;
      var st=String(v[i][iS]||'');
      if(st==='ปฏิเสธ'||st==='ยกเลิก') continue;
      cnt++;
    }
  }
  if(cnt>=lim) return 'สถาบัน "'+String(org).trim()+'" ใช้สิทธิ์'+TYPE_LBL[type]+'ครบ '+lim+' ท่านแล้ว (ตามเงื่อนไขสถาบันร่วมจัด) — หากเป็นผู้เข้าร่วมทั่วไปกรุณาเลือกประเภท "ผู้เข้าร่วมทั่วไป"';
  return '';
}

// กันส่งซ้ำ: หน้าเว็บส่ง submitId (สุ่มครั้งเดียวต่อการกดส่ง) — ถ้าเคยบันทึกแล้ว คืนผลเดิม ไม่สร้างซ้ำ
// (กรณีเน็ตช้า/หมดเวลาแล้วระบบลองส่งซ้ำให้อัตโนมัติ หรือผู้ใช้กดซ้ำ)
function _submitCacheKey(d){ var id=String(d&&d.submitId||'').trim(); return /^[A-Za-z0-9_-]{8,64}$/.test(id) ? 'reg_submit_'+id : ''; }
function _registerUser(d) {
  var ck = _submitCacheKey(d), cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (ck && cache) { try { var prev = cache.get(ck); if (prev) { var pr = JSON.parse(prev); pr.replayed = true; return pr; } } catch (e) {} }

  // Online (c00) กับ Onsite (c00s) เลือกได้อย่างใดอย่างหนึ่ง
  if (d.type !== 'committee_dir' && d.type !== 'committee_judge') {
    var _cs = d.comps || [];
    if (_cs.indexOf('c00') >= 0 && _cs.indexOf('c00s') >= 0) {
      return { success: false, message: 'เลือกได้อย่างใดอย่างหนึ่งระหว่าง Online และ Onsite' };
    }
  }

  // ตรวจซ้ำแบบ warn-only (ไม่ block — user อาจลงหลายรายการได้) — ทำนอก lock เพื่อให้ถือ lock สั้นที่สุด
  var dup = _checkDuplicate(d);
  var dupWarning = dup.duplicate ? { field: dup.field, rid: dup.rid } : null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (le) {
    // busy:true → หน้าเว็บจะลองส่งซ้ำให้เองอัตโนมัติ (ยังไม่มีอะไรถูกบันทึก)
    return { success: false, busy: true, message: 'ระบบมีคนลงทะเบียนพร้อมกันเยอะ กำลังลองใหม่ให้อัตโนมัติ…' };
  }
  var res, mail = null;
  try {
    if (ck && cache) { try { var prev2 = cache.get(ck); if (prev2) { var pr2 = JSON.parse(prev2); pr2.replayed = true; return pr2; } } catch (e) {} }
    var ss       = SpreadsheetApp.openById(SS_ID);
    if(COHOST_LIMIT[d.type]){
      var qMsg=_checkCohostQuota(ss,d.type,d.dirOrg||d.judgeOrg||d.org||'');
      if(qMsg) return { success:false, quotaFull:/ครบ/.test(qMsg), message:qMsg };
    }
    var regId    = _genRegId(ss);
    var isFree   = (d.type==='committee_dir'||d.type==='committee_judge');
    var fee      = isFree ? 0 : _calcFee(d);
    var typeLbl  = TYPE_LBL[d.type]    || d.type || 'ไม่ระบุ';
    var typeLblEn= TYPE_LBL_EN[d.type] || d.type;
    var org      = d.dirOrg||d.judgeOrg||d.org||'';
    var fullName = ((d.prefix||'')+' '+(d.fname||'')+' '+(d.lname||'')).trim();
    var judgeComp= d.judgeComp?(COMP[d.judgeComp]?COMP[d.judgeComp].nm:d.judgeComp):'';
    var meal     = _calcMeal(d, isFree, fullName, org, typeLbl, regId);

    // REGISTRATIONS
    _sheet(ss,SH_REG,[
      'REG_ID','วันที่ลงทะเบียน','คำนำหน้า','ชื่อ','นามสกุล','ชื่อเต็ม',
      'อีเมล','โทรศัพท์','เลขบัตรประชาชน',
      'ประเภทผู้ลงทะเบียน','ฟรี/ชำระ','สังกัด/หน่วยงาน','ตัดสินประเภท',
      'การแพ้อาหาร','ที่อยู่ใบเสร็จ','เลขใบประกอบวิชาชีพ','ยอดรวม','สถานะชำระ',
      'อาหาร_27พค','อาหาร_28พค','คนได้ข้าว_รวม'
    ]).appendRow([
      regId,new Date(),_safe(d.prefix||''),_safe(d.fname||''),_safe(d.lname||''),_safe(fullName),
      _safe(d.email||''),_safe(d.phone||''),_safe(d.idcard||''),
      _safe(typeLbl),isFree?'ฟรี':'ชำระเงิน',_safe(org),_safe(judgeComp),
      _safe(d.allergy||''),_safe(d.addr||''),_safe(d.license||''),fee,'รอชำระเงิน',
      meal.d17Count,meal.d18Count,meal.totalMealPersons
    ]);

    // COMPETITIONS
    if(!isFree&&(d.comps||[]).length){
      var csh=_sheet(ss,SH_COMP,[
        'REG_ID','ลำดับ','รหัสรายการ','ชื่อรายการ',
        'จำนวนวัน','โควต้าทีม','รายชื่อสมาชิก','วันที่บันทึก'
      ]);
      var crows=(d.comps||[]).map(function(cid,i){
        var days=(d.compDays||{})[cid]||1;
        var team=(d.members||{})[cid]||[];
        var names=team.map(function(m){return((m.prefix||'')+' '+(m.fname||m.name||'')+' '+(m.lname||'')).trim();}).filter(Boolean).join(', ');
        return [regId,i+1,_safe(String(cid)),_safe(COMP[cid]?COMP[cid].nm:String(cid)),days,(COMP[cid]||{}).team||1,_safe(names),new Date()];
      });
      _appendRows(csh,crows);
    }

    // MEAL_SUMMARY rows
    _writeMealRows(ss,regId,meal,fullName,org,typeLbl,d.allergy||'');

    // COMPETITION SHEETS — sheet แยกรายการ
    if(!isFree && (d.comps||[]).length){
      _writeCompSheets(ss, regId, d, fullName, org, fee);
    }
    SpreadsheetApp.flush();
    _clearMealCache();

    if(d.email) mail = {fullName:fullName,org:org,typeLbl:typeLbl,typeLblEn:typeLblEn,email:d.email,regId:regId,fee:fee,isFree:isFree,d:d,meal:meal};

    res = {
      success:    true,
      rid:        regId,
      fee:        fee,
      meal:       {d17:meal.d17Count, d18:meal.d18Count, total:meal.totalMealPersons},
      mealRows:   meal.rows.map(function(r){return{nm:r.src,count:r.count||1,d17:r.d17,d18:r.d18};}),
      emailStatus: 'not_sent',
      emailError:  '',
      printToken:  _printToken(regId),   // ใช้สร้างลิงก์ print.html?rid=..&k=.. บนเว็บที่โฮสต์แยก
      dupWarning:  dupWarning   // แจ้งกลับว่าเคยมีข้อมูลซ้ำ แต่ยังสมัครได้
    };
    if (ck && cache) { try { cache.put(ck, JSON.stringify(res), 21600); } catch (e) {} } // จำไว้ 6 ชม.
  } catch(err){
    Logger.log('_registerUser: '+err.message);
    res = {success:false, message:err.message};
  } finally {
    lock.releaseLock();
  }

  // ── หลังปล่อย lock: ส่งอีเมล + จองคิว auto-export (ไม่บล็อกผู้ลงทะเบียนรายอื่น) ──
  if(res.success){
    if(mail){
      // ส่งผ่านคิวเบื้องหลัง (ภายใน ~1 นาที) — ไม่รอ GmailApp ในคำขอนี้
      try{
        _queueEmail('register',mail.regId,{fullName:mail.fullName,org:mail.org,typeLbl:mail.typeLbl,typeLblEn:mail.typeLblEn,email:mail.email,fee:mail.fee,isFree:mail.isFree,d:mail.d});
        res.emailStatus = 'queued';
      } catch(emailErr){
        res.emailStatus = 'error';
        res.emailError  = emailErr.message;
        Logger.log('EMAIL QUEUE ERROR for '+mail.regId+': '+emailErr.message);
      }
    }
    // Auto-export: จองคิวให้รันเบื้องหลัง ไม่บล็อกการตอบกลับผู้ใช้
    try{ _scheduleAutoExports(); }catch(autoErr){ Logger.log('schedule auto export err: '+autoErr.message); }
  }
  return res;
}

function _calcFee(d){
  return(d.comps||[]).reduce(function(s,cid){
    var c=COMP[cid];if(!c)return s;
    var days=(d.compDays||{})[cid]||1;
    return s+(days===2&&c.d2?c.d2:c.d1);
  },0);
}

// ── SECTION 5: MEAL CALCULATION ──────────────────────────────
// Rules:
//  ประชุมวิชาการ (c00/c00s)  → BOTH days always (1 person)
//  d18 events (c02-c04)       → 1 day=d18 only · 2 days=d17+d18
//  d17 events (c05-c11)       → 1 day=d17 only · 2 days=d17+d18
//  Committee (isFree)         → BOTH days (1 person)
// Person count = team QUOTA, not actual members entered

function _calcMeal(d, isFree, fullName, org, typeLbl, regId) {
  var d17Set={}, d18Set={};

  function add(set, key, name, idcard, typeStr, orgStr, src) {
    if(!set[key]) set[key]={name:name,idcard:idcard||'',type:typeStr,org:orgStr,sources:[]};
    if(src&&set[key].sources.indexOf(src)<0) set[key].sources.push(src);
  }

  if(isFree){
    var k=(d.idcard&&d.idcard.length===13)?d.idcard:fullName.toLowerCase().replace(/\s/g,'');
    add(d17Set,k,fullName,d.idcard||'',typeLbl,org,'กรรมการ');
    add(d18Set,k,fullName,d.idcard||'',typeLbl,org,'กรรมการ');
  } else {
    (d.comps||[]).forEach(function(cid){
      var cfg=COMP[cid];if(!cfg)return;
      var days=(d.compDays||{})[cid]||1;
      var quota=cfg.team||1;
      var teamArr=(d.members||{})[cid]||[];

      for(var i=0;i<quota;i++){
        var m=teamArr[i];
        var name,idcard;
        if(m&&((m.fname||m.name||'').trim())){
          name=((m.prefix||'')+' '+(m.fname||m.name||'')+' '+(m.lname||'')).trim();
          idcard=m.idcard||'';
        } else if(i===0){
          name=fullName; idcard=d.idcard||'';
        } else {
          name='สมาชิก '+(i+1)+' ('+cfg.nm+')'; idcard='';
        }
        var key=(idcard&&idcard.length===13)?idcard:(name+'_'+cid+'_'+i).toLowerCase().replace(/\s/g,'');

        if(cfg.mealDay==='none'){ /* ออนไลน์ ไม่มีอาหาร */ } else if(cfg.mealDay==='both'){
          add(d17Set,key,name,idcard,typeLbl,org,cfg.nm);
          add(d18Set,key,name,idcard,typeLbl,org,cfg.nm);
        } else if(cfg.mealDay==='d18'){
          if(days===2) add(d17Set,key,name,idcard,typeLbl,org,cfg.nm+' (2 วัน)');
          add(d18Set,key,name,idcard,typeLbl,org,cfg.nm);
        } else { // d17
          add(d17Set,key,name,idcard,typeLbl,org,cfg.nm);
          if(days===2) add(d18Set,key,name,idcard,typeLbl,org,cfg.nm+' (2 วัน)');
        }
      }
    });
  }

  // นับตามโควต้าต่อรายการ (ไม่ deduplicate) — ใช้ตัวเดียวกับคูปอง (_mealCoupons)
  var mc=_mealCoupons(_couponListFromPayload(d,isFree),isFree,fullName);
  var d17Count=mc.d17.length, d18Count=mc.d18.length;

  // Build rows for MEAL_SUMMARY sheet (ยังคง dedup เพื่อรายชื่อ)
  var allKeys={};
  Object.keys(d17Set).forEach(function(k){allKeys[k]=true;});
  Object.keys(d18Set).forEach(function(k){allKeys[k]=true;});
  var rows=Object.keys(allKeys).map(function(k){
    var base=d17Set[k]||d18Set[k];
    return{regId:regId,name:base.name,idcard:base.idcard,type:base.type,
           org:base.org,src:base.sources.join(', '),d17:!!d17Set[k],d18:!!d18Set[k]};
  });

  // Build summary rows per competition
  var summaryRows=[];
  if(!isFree){
    (d.comps||[]).forEach(function(cid){
      var cfg=COMP[cid];if(!cfg)return;
      var days=(d.compDays||{})[cid]||1;
      var getsD17=cfg.mealDay==='both'||cfg.mealDay==='d17'||(cfg.mealDay==='d18'&&days===2);
      var getsD18=cfg.mealDay==='both'||cfg.mealDay==='d18'||(cfg.mealDay==='d17'&&days===2);
      summaryRows.push({nm:cfg.nm,nm_en:cfg.nm_en,count:cfg.team,d17:getsD17,d18:getsD18,src:cfg.nm});
    });
  } else {
    summaryRows.push({nm:'กรรมการ',nm_en:'Committee',count:1,d17:true,d18:true,src:'กรรมการ'});
  }

  return{
    d17Count:d17Count,
    d18Count:d18Count,
    totalMealPersons:d17Count+d18Count,
    rows:rows,
    summaryRows:summaryRows
  };
}

// ── MEAL E-COUPONS ───────────────────────────────────────────
var MEAL_NAME='อาหารกลางวัน', MD17='27 พ.ค.', MD18='28 พ.ค.';
// compList=[{cid,days,members:[{name,allergy}]}] → {d17:[{n,total,name,comp,allergy}],d18:[..],rules:[{comp,seats,days}]}
function _mealCoupons(compList,isFree,fullName){
  var r={d17:[],d18:[],rules:[]};
  function put(arr,name,comp,al){arr.push({n:0,total:0,name:name,comp:comp,allergy:al||''});}
  if(isFree){
    var m0=(((compList||[])[0]||{}).members||[])[0]||{};
    put(r.d17,fullName,'',m0.allergy); put(r.d18,fullName,'',m0.allergy);
    r.rules.push({comp:'กรรมการ',seats:1,days:MD17+'–'+MD18});
  } else {
    (compList||[]).forEach(function(c){
      var cfg=COMP[c.cid]; if(!cfg) return;
      var q=cfg.team||1, days=c.days||1;
      var g17=cfg.mealDay==='both'||cfg.mealDay==='d17'||(cfg.mealDay==='d18'&&days===2);
      var g18=cfg.mealDay==='both'||cfg.mealDay==='d18'||(cfg.mealDay==='d17'&&days===2);
      for(var i=0;i<q&&(g17||g18);i++){
        var m=(c.members||[])[i]||{};
        var nm=m.name||(i===0?fullName:'สมาชิก '+(i+1)+' ('+cfg.nm+')');
        if(g17) put(r.d17,nm,cfg.nm,m.allergy);
        if(g18) put(r.d18,nm,cfg.nm,m.allergy);
      }
      r.rules.push({comp:cfg.nm,seats:(g17||g18)?q:0,days:g17&&g18?MD17+'–'+MD18:g17?MD17:g18?MD18:'ไม่มีอาหาร'});
    });
  }
  [r.d17,r.d18].forEach(function(a){a.forEach(function(x,i){x.n=i+1;x.total=a.length;});});
  return r;
}
function _fullNm(m){return ((m.prefix||'')+' '+(m.fname||m.name||'')+' '+(m.lname||'')).trim();}
// compList จาก payload ตอนลงทะเบียน (ตรงกับชื่อที่ _calcMeal ใช้)
function _couponListFromPayload(d,isFree){
  if(isFree) return [{cid:'',days:1,members:[{name:'',allergy:d.allergy||''}]}];
  return (d.comps||[]).map(function(cid){
    var arr=(d.members||{})[cid]||[], mem=[];
    for(var i=0;i<4;i++){
      var m=arr[i]; var ok=m&&((m.fname||m.name||'').trim());
      mem.push({name:ok?_fullNm(m):'',allergy:ok?(m.allergy||''):(i===0?(d.allergy||''):'')});
    }
    return {cid:cid,days:(d.compDays||{})[cid]||1,members:mem};
  });
}
// compList จากชีทรายการ ตอนอนุมัติ → {isFree,fullName,compList}
function _couponListFromSheets(ss,regId,fullName){
  var id=String(regId||'').trim(), isFree=false, al0='';
  var rsh=ss.getSheetByName(SH_REG);
  if(rsh&&rsh.getLastRow()>1){
    var rd=rsh.getDataRange().getValues(), h=rd[0], iF=h.indexOf('ฟรี/ชำระ'), iT=h.indexOf('ประเภทผู้ลงทะเบียน'), iA=h.indexOf('การแพ้อาหาร');
    for(var j=1;j<rd.length;j++) if(String(rd[j][0]).trim()===id){
      isFree=String(rd[j][iF]||'')==='ฟรี'||[TYPE_LBL.committee_dir,TYPE_LBL.committee_judge].indexOf(String(rd[j][iT]||''))>=0;
      al0=String(rd[j][iA]||''); break;
    }
  }
  var list=[];
  if(isFree) return {isFree:true,compList:[{cid:'',days:1,members:[{name:'',allergy:al0}]}]};
  var myCids=_regCids(ss,regId); // เปิดเฉพาะชีทของผู้ลงทะเบียน (ไม่มีแถว → วนทั้งหมดเหมือนเดิม)
  Object.keys(COMP_SHEET).forEach(function(cid){
    if(cid==='c00') return;
    if(myCids&&myCids.indexOf(cid)<0) return;
    var sh=ss.getSheetByName(COMP_SHEET[cid]); if(!sh||sh.getLastRow()<2) return;
    var top=sh.getRange(1,1,Math.min(3,sh.getLastRow()),sh.getLastColumn()).getValues(), hr=-1;
    for(var t=0;t<top.length&&hr<0;t++) if(top[t].indexOf('REG_ID')>=0) hr=t;
    if(hr<0||sh.getLastRow()<=hr+1) return;
    var v=sh.getRange(hr+1,1,sh.getLastRow()-hr,sh.getLastColumn()).getValues(), H=v[0], iR=H.indexOf('REG_ID'), iD=H.indexOf('จำนวนวัน');
    for(var k=1;k<v.length;k++) if(String(v[k][iR]).trim()===id){
      var mem=[];
      for(var n=1;n<=4;n++){
        var iFn=H.indexOf('ชื่อ '+n); if(iFn<0) iFn=H.indexOf('ชื่อ สมาชิก '+n);
        var iP=H.indexOf('คำนำหน้า '+n), iL=H.indexOf('นามสกุล '+n), iAl=H.indexOf('แพ้อาหาร '+n);
        var g=function(x){return x<0?'':String(v[k][x]||'').trim();};
        mem.push({name:g(iFn)?(g(iP)+' '+g(iFn)+' '+g(iL)).trim():'',allergy:g(iAl)});
      }
      list.push({cid:cid,days:Number(v[k][iD])===2?2:1,members:mem}); break;
    }
  });
  return {isFree:false,compList:list};
}
// ── HTML คูปอง (ตาราง + inline style เท่านั้น) ──
var CP_G='#b8862b', CP_D='#0b5844', CP_C='#fbf6ec', CP_C2='#f6efe1', CP_R='#f3e6c4';
var CP_F="font-family:'Sarabun','Prompt',Tahoma,Arial,sans-serif;";
var CP_M="font-family:'Courier New',Courier,monospace;";
var CP_BOX='border:2px solid '+CP_G+';background:'+CP_C+';border-radius:12px;border-collapse:separate;';
var CP_TD='padding:8px 10px;border-top:1px solid #e3d3a8;font-size:13px;word-break:break-word;';
// บาร์โค้ดตกแต่ง (ไม่ใช่โค้ดจริง) — กำหนดจาก REG ID + ลำดับคูปอง
function _cpBarcode(seed){
  var s=0,i; for(i=0;i<seed.length;i++) s=(s*31+seed.charCodeAt(i))>>>0;
  var h='<table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>';
  for(i=0;i<40;i++){
    s=(s*1103515245+12345)>>>0; var w=1+((s>>>16)%3);
    h+=(i%2===0)?'<td style="width:'+w+'px;height:16px;background:#1f2937;font-size:0;line-height:0">&nbsp;</td>':'<td style="width:'+w+'px;font-size:0;line-height:0"></td>';
  }
  return h+'</tr></table>';
}
function _couponSummaryHtml(c){
  function cell(day,n){
    var on=n>0, p=day.split(' ');
    return '<td width="50%" align="center" valign="top" style="padding:0;background:#fff;border:1px solid #e3d3a8;border-radius:8px">'+
      '<div style="background:'+(on?CP_D:'#9ca3af')+';color:#fff;font-size:12px;font-weight:700;padding:5px 4px;border-radius:7px 7px 0 0;letter-spacing:.03em">'+day+' 2570</div>'+
      (on?'<div style="padding:8px 4px 10px;color:'+CP_D+';font-weight:700"><span style="font-size:36px;line-height:42px">'+n+'</span> <span style="font-size:14px">ที่</span></div>'
         :'<div style="padding:12px 4px 14px;font-size:16px;line-height:36px;font-weight:700;color:#9ca3af">ไม่มีสิทธิ์</div>')+'</td>';
  }
  var th='padding:7px 10px;font-size:12px;color:'+CP_D+';background:'+CP_R+';border-bottom:2px solid '+CP_G+';';
  var rows=c.rules.map(function(r){
    return '<tr><td style="'+CP_TD+'">'+_esc(r.comp)+'</td><td align="center" style="'+CP_TD+'font-weight:700">'+r.seats+'</td><td align="center" style="'+CP_TD+'">'+_esc(r.days)+'</td></tr>';
  }).join('');
  return '<table width="100%" cellpadding="0" cellspacing="0" style="'+CP_BOX+'margin:14px 0;'+CP_F+'"><tr><td style="padding:14px">'+
    '<div style="font-size:16px;font-weight:700;color:'+CP_D+';text-align:center;margin:0 0 10px">สิทธิ์อาหารของท่าน</div>'+
    '<table width="100%" cellpadding="0" cellspacing="6" style="border-collapse:separate"><tr>'+cell(MD17,c.d17.length)+cell(MD18,c.d18.length)+'</tr></table>'+
    '<div style="font-size:14px;font-weight:700;color:#7a1f1f;text-align:center;margin:10px 0">1 คูปอง = อาหาร 1 ที่ · กรุณาแสดงคูปองนี้ที่จุดรับอาหาร (บนมือถือหรือพิมพ์)</div>'+
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:#fff"><tr><th align="left" style="'+th+'">รายการ</th><th style="'+th+'">จำนวนที่</th><th style="'+th+'">วัน</th></tr>'+rows+'</table>'+
    '</td></tr></table>';
}
function _couponTicket(x,day,regId,compact,gap){
  var dp=day.split(' '), al=x.allergy?'<div style="margin-top:6px"><span style="display:inline-block;background:#fdecec;color:#b91c1c;font-size:12px;font-weight:700;padding:3px 8px;border-radius:6px">⚠ แพ้อาหาร: '+_esc(x.allergy)+'</span></div>':'';
  var sw=compact?64:104;
  var stub='<td width="'+sw+'" align="center" valign="middle" style="width:'+sw+'px;background:'+CP_D+';color:#fff;padding:'+(compact?'8px 4px':'14px 6px')+';border-radius:10px 0 0 0;'+CP_F+'">'+
    (compact?'':'<div style="font-size:10px;letter-spacing:.2em;color:#e8cc7a;font-weight:700">LUNCH</div><div style="font-size:11px;color:#e8cc7a">'+MEAL_NAME+'</div>')+
    '<div style="font-size:'+(compact?22:34)+'px;line-height:'+(compact?26:40)+'px;font-weight:700;margin:'+(compact?0:'4px 0 2px')+'">'+_esc(dp[0])+'</div>'+
    '<div style="font-size:11px;color:#f3e6c4">'+_esc(dp[1]||'')+' 2570</div></td>';
  var body='<td valign="middle" style="padding:'+(compact?'8px 10px':'12px 14px')+';border-left:2px dashed '+CP_G+';'+CP_F+'">'+
    '<table width="100%" cellpadding="0" cellspacing="0"><tr><td valign="middle"><span style="display:inline-block;background:'+CP_R+';border:1px solid '+CP_G+';color:#7a5a12;font-size:11px;font-weight:700;padding:1px 8px;border-radius:10px">คูปอง '+x.n+'/'+x.total+'</span></td>'+
    (x.comp?'<td align="right" valign="middle" style="font-size:11px;color:#6b7280;padding-left:6px">'+_esc(x.comp)+'</td>':'')+'</tr></table>'+
    '<div style="font-size:'+(compact?15:19)+'px;font-weight:700;color:#1f2937;line-height:1.35;margin:'+(compact?'4px 0':'8px 0 10px')+'">'+_esc(x.name)+'</div>'+
    '<div style="background:'+CP_C2+';border:1px solid '+CP_G+';border-radius:6px;padding:'+(compact?'4px 6px':'6px 10px')+';'+CP_M+'font-size:'+(compact?13:16)+'px;font-weight:700;color:'+CP_D+';letter-spacing:1px;word-break:break-all">'+_esc(regId)+'</div>'+al+'</td>';
  var strip=compact?'':'<tr><td colspan="2" style="background:'+CP_C2+';border-top:1px solid #e3d3a8;padding:6px 12px 5px;border-radius:0 0 10px 10px;'+CP_F+'">'+_cpBarcode(regId+'#'+day+x.n)+
    '<div style="font-size:10px;color:#6b7280;margin-top:2px">สำหรับเจ้าหน้าที่: ค้นหาด้วย REG ID</div></td></tr>';
  return '<table class="cp-ticket" width="100%" cellpadding="0" cellspacing="0" style="'+CP_BOX+'"><tr>'+stub+body+'</tr>'+strip+'</table>'+
    (gap?'<div style="text-align:center;color:#c9b887;font-size:11px;letter-spacing:2px;margin:6px 0;line-height:14px;overflow:hidden">✂ - - - - - - - - - - - - - - - - - -</div>':'');
}
function _couponDayHtml(arr,day,regId,compact){
  if(!arr.length) return '';
  var h='<table class="cp-day" width="100%" cellpadding="0" cellspacing="0" style="margin:18px 0 10px;border-collapse:separate;background:'+CP_R+';border-bottom:2px solid '+CP_G+';border-radius:8px 8px 0 0;'+CP_F+'"><tr>'+
    '<td style="padding:8px 12px;font-size:15px;font-weight:700;color:'+CP_D+'">คูปองอาหารกลางวัน</td>'+
    '<td align="right" style="padding:8px 12px"><span style="display:inline-block;background:'+CP_D+';color:#fff;font-size:12px;font-weight:700;padding:2px 10px;border-radius:12px;white-space:nowrap">'+day+' 2570 · '+arr.length+' ที่</span></td></tr></table>';
  if(!compact) return h+arr.map(function(x,i){return '<div style="margin:0 0 4px">'+_couponTicket(x,day,regId,false,i<arr.length-1)+'</div>';}).join('');
  h+='<table width="100%" cellpadding="0" cellspacing="4" style="border-collapse:separate">';
  for(var i=0;i<arr.length;i+=2)
    h+='<tr><td width="50%" valign="top">'+_couponTicket(arr[i],day,regId,true,false)+'</td><td width="50%" valign="top">'+(arr[i+1]?_couponTicket(arr[i+1],day,regId,true,false):'')+'</td></tr>';
  return h+'</table>';
}
// full=true → สรุป+คูปองทั้งหมด · false → สรุป+กติกาเท่านั้น (ยังไม่ชำระ)
function _couponHtml(regId,c,full){
  var h='<style>@media print{ .cp-day{page-break-before:always} .cp-ticket{page-break-inside:avoid} }</style>'+_couponSummaryHtml(c);
  if(full===false) return h+'<div style="font-size:12px;color:#6b5a3a;background:#fff8e1;border-left:4px solid '+CP_G+';padding:8px 12px;margin:0 0 12px">คูปองอาหารจะส่งให้ทางอีเมลหลังเจ้าหน้าที่ยืนยันการชำระเงิน</div>';
  var note='<div style="font-size:12px;color:#7a1f1f;font-weight:700;margin:10px 0">คูปองใช้ได้เฉพาะวันที่ระบุ · ไม่สามารถเพิ่มจำนวนหน้างาน</div>';
  var out=h+_couponDayHtml(c.d17,MD17,regId,false)+_couponDayHtml(c.d18,MD18,regId,false)+note;
  if(out.length>90000) out=h+_couponDayHtml(c.d17,MD17,regId,true)+_couponDayHtml(c.d18,MD18,regId,true)+note;
  return out;
}
// ชุดคูปองจากชีท (ตอนอนุมัติ) — คืน html ว่างถ้าไม่มีสิทธิ์/พลาด
function _couponHtmlFromSheets(regId,fullName,d17,d18){
  try{
    var b=_couponListFromSheets(SpreadsheetApp.openById(SS_ID),regId,fullName);
    var c=_mealCoupons(b.compList,b.isFree,fullName);
    if(c.d17.length!==Number(d17||0)||c.d18.length!==Number(d18||0))
      Logger.log('COUPON MISMATCH ['+regId+'] sheet d17/d18='+c.d17.length+'/'+c.d18.length+' REGISTRATIONS='+d17+'/'+d18);
    if(!c.d17.length&&!c.d18.length) return '';
    return _couponHtml(regId,c,true);
  }catch(e){ Logger.log('COUPON ERROR ['+regId+']: '+e.message); return ''; }
}

function _writeMealRows(ss,regId,mealResult,fullName,org,typeLbl,allergy){
  var sh=_sheet(ss,SH_MEAL,[
    'REG_ID','ชื่อ-นามสกุล','เลขบัตรประชาชน','ประเภท','สังกัด','รายการ/ที่มา',
    'อาหาร_27พค','อาหาร_28พค','สถานะชำระเงิน','รับ_27พค','รับ_28พค','การแพ้อาหาร','วันที่บันทึก'
  ]);
  var rows=(mealResult.summaryRows||[]).filter(function(r){
    return (r.d17||0)>0||(r.d18||0)>0; // กรอง Online (ไม่มีสิทธิ์อาหาร)
  });
  if(!rows.length) return;
  // batch: เขียนทุกแถวรวดเดียว (สูตรเป็น string ขึ้นต้น "=" → setValues ตั้งเป็นสูตรให้) — คอลัมน์เรียงเดิม 13 คอลัมน์
  var start=sh.getLastRow()+1;
  var vals=rows.map(function(r,i){
    var rn=start+i;
    return [regId,_safe(fullName),'',_safe(typeLbl),_safe(org),_safe(r.nm||r.src||''),
      r.d17?(r.count||1):0, r.d18?(r.count||1):0,
      '=IFERROR(VLOOKUP(A'+rn+',PAYMENTS!$A:$G,7,0),"รอชำระเงิน")',
      '=IFERROR(IF(COUNTIFS(MEAL_CHECKIN!$A:$A,A'+rn+',MEAL_CHECKIN!$D:$D,"*27*")>0,"✓ รับแล้ว"),"")',
      '=IFERROR(IF(COUNTIFS(MEAL_CHECKIN!$A:$A,A'+rn+',MEAL_CHECKIN!$D:$D,"*28*")>0,"✓ รับแล้ว"),"")',
      _safe(allergy||''),new Date()];
  });
  _appendRows(sh,vals);
}
function _getStatusData(query){
  if(!query) return{success:false,message:'กรุณากรอกข้อมูลค้นหา'};
  var ss=SpreadsheetApp.openById(SS_ID);
  var sh=ss.getSheetByName(SH_REG);
  if(!sh||sh.getLastRow()<2) return{success:false,message:'ไม่พบข้อมูล'};
  var data=sh.getDataRange().getValues(),hdr=data[0];
  var q=String(query).trim().toLowerCase();
  var iRid=0;
  var iName=hdr.indexOf('ชื่อเต็ม');
  var iPre=hdr.indexOf('คำนำหน้า'),iFn=hdr.indexOf('ชื่อ'),iLn=hdr.indexOf('นามสกุล');
  var iCard=hdr.indexOf('เลขบัตรประชาชน');
  var matches=[];
  for(var i=1;i<data.length;i++){
    var rid=String(data[i][iRid]).trim();
    if(!rid) continue;
    var fullName=String(data[i][iName]||'').trim();
    if(!fullName) fullName=((String(data[i][iPre]||'')+' '+String(data[i][iFn]||'')+' '+String(data[i][iLn]||'')).trim().replace(/\s+/g,' '));
    var nameLow=fullName.toLowerCase();
    var card=String(data[i][iCard]||'').trim();
    // RID / เลขบัตร = ตรงทั้งหมดเท่านั้น; ชื่อ = substring เฉพาะ query ยาว >= 3 ตัวอักษร
    if(rid.toLowerCase()===q || (/^\d{4}-\d{4}$/.test(q) && rid.toLowerCase().slice(-9)===q) || (q && card===q) || (q.length>=3 && nameLow.indexOf(q)>=0)){
      var obj={};hdr.forEach(function(h,j){obj[h]=data[i][j];});
      matches.push({
        success:true,rid:rid,
        name:fullName,org:String(obj['สังกัด/หน่วยงาน']||''),
        regType:String(obj['ประเภทผู้ลงทะเบียน']||''),
        payStatus:String(obj['สถานะชำระ']||''),fee:Number(obj['ยอดรวม']||0),
        mealD17:Number(obj['อาหาร_27พค']||0),mealD18:Number(obj['อาหาร_28พค']||0)
      });
      if(rid.toLowerCase()===q) matches[matches.length-1].printToken=_printToken(rid); // เฉพาะค้นด้วย REG ID เต็ม (ไม่ใช่ชื่อ/เลขบัตร)
    }
  }
  if(matches.length===0) return{success:false,message:'ไม่พบข้อมูล: '+query};
  if(matches.length===1){ return matches[0]; } // ไม่คืน payment/slipUrl (public endpoint)
  // พบหลายคน — return list แบบย่อ (ไม่คืน payment/fee/meals) จำกัด 20 รายการ
  var list=matches.slice(0,20).map(function(m){return{rid:m.rid,name:m.name,org:m.org,payStatus:m.payStatus};});
  return{success:true,multiple:true,matches:list,total:matches.length};
}

function _getLatestPayment(ss,regId){
  var sh=ss.getSheetByName(SH_PAY);
  if(!sh||sh.getLastRow()<2) return null;
  var rows=sh.getDataRange().getValues().slice(1).filter(function(r){return String(r[0])===String(regId);});
  if(!rows.length) return null;
  var r=rows[rows.length-1];
  return{amount:r[2],method:r[3],ref:r[4],slipUrl:r[5],status:r[6]};
}

// ── SECTION 7: PAYMENT SLIP ──────────────────────────────────

function _submitSlip(p){
  var lock=null;
  try{
    var regId=String(p.regId||'').trim();
    var ss=SpreadsheetApp.openById(SS_ID);
    var info=_getRegInfo(ss,regId);
    if(!info) return{success:false,message:'ไม่พบ REG ID'};
    if(info.status==='ชำระแล้ว') return{success:false,message:'REG ID นี้ชำระเงินและได้รับการยืนยันแล้ว'};

    var slipUrl='';
    if(p.slipBase64||p.fileB64){
      var b64=p.slipBase64||p.fileB64;
      var mime=String(p.fileType||'image/jpeg').toLowerCase();
      var EXT={'image/jpeg':'jpg','image/png':'png','image/webp':'webp','application/pdf':'pdf'};
      if(!EXT[mime]) return{success:false,message:'รองรับเฉพาะไฟล์ JPG, PNG, WEBP หรือ PDF'};
      if(String(b64).length*0.75>6*1024*1024) return{success:false,message:'ไฟล์ใหญ่เกิน 6MB'};
      var folder=DriveApp.getFolderById(DRIVE_FOLDER_ID);
      var it=folder.getFoldersByName('SLIPS');
      var sf=it.hasNext()?it.next():folder.createFolder('SLIPS');
      var ts=Utilities.formatDate(new Date(),'Asia/Bangkok','yyyyMMdd_HHmmss');
      var blob=Utilities.newBlob(Utilities.base64Decode(b64),mime,regId.replace(/[^\w\-]/g,'')+'_'+ts+'.'+EXT[mime]);
      var file=sf.createFile(blob);
      try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); }
      catch(se){ try{ file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK,DriveApp.Permission.VIEW); }catch(e2){ Logger.log('slip share failed: '+se.message); } }
      slipUrl=file.getUrl();
    }
    lock=LockService.getScriptLock();
    lock.waitLock(15000);
    info=_getRegInfo(ss,regId); // ตรวจซ้ำใน lock
    if(!info) return{success:false,message:'ไม่พบ REG ID'};
    if(info.status==='ชำระแล้ว') return{success:false,message:'REG ID นี้ชำระเงินและได้รับการยืนยันแล้ว'};
    _sheet(ss,SH_PAY,[
      'REG_ID','วันที่ส่ง','จำนวนเงิน','วิธีชำระ','วันที่โอน','URL_สลิป','สถานะ','ผู้อนุมัติ','วันอนุมัติ','ที่อยู่ใบเสร็จ'
    ]).appendRow([regId,new Date(),_safe(p.amount||0),_safe(p.method||'โอนเงิน'),_safe(p.transferDate||p.ref||''),slipUrl,'รอตรวจสอบ','','',_safe(p.addr||'')]);
    _updatePayStatus(ss,regId,'รอตรวจสอบ');
    _clearMealCache();
    return{success:true,message:'โอนเงินเรียบร้อยแล้ว ทีมงานจะตรวจสอบและยืนยันภายใน 1–2 วันทำการ',slipUrl:slipUrl};
  }catch(err){return{success:false,message:err.message};}
  finally{ if(lock){ try{lock.releaseLock();}catch(e){} } }
}

function _updatePayStatus(ss,regId,status){
  var sh=ss.getSheetByName(SH_REG);
  if(!sh||sh.getLastRow()<2) return;

  // หา col สถานะชำระ — ลอง indexOf ก่อน ถ้าไม่เจอใช้ hardcode R=18
  var hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var col=hdr.indexOf('สถานะชำระ');
  if(col<0){
    // scan หา column ที่ขึ้นต้นด้วย สถานะ
    for(var h=0;h<hdr.length;h++){
      if(String(hdr[h]).indexOf('สถานะชำระ')>=0){col=h;break;}
    }
  }
  if(col<0) col=17; // hardcode col R (0-based)

  // หา REG_ID แล้วเขียน
  var ridCol=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<ridCol.length;i++){
    if(String(ridCol[i][0]).trim()===String(regId).trim()){
      sh.getRange(i+2,col+1).setValue(status);
      SpreadsheetApp.flush();
      break;
    }
  }
  // อัปเดต competition sheets ด้วย
  try{ _updateCompSheetsPayStatus(ss,regId,status); }catch(e){ Logger.log('compSheet err:'+e.message); }
}

// ── SECTION 8: STAFF ─────────────────────────────────────────

function _getPendingSlips(){
  try{
    var ss=SpreadsheetApp.openById(SS_ID);
    var psh=ss.getSheetByName(SH_PAY),rsh=ss.getSheetByName(SH_REG);
    if(!psh||psh.getLastRow()<2) return{success:true,slips:[]};
    var regMap={};
    if(rsh&&rsh.getLastRow()>1){
      var rd=rsh.getDataRange().getValues(),rh=rd[0];
      var iN=rh.indexOf('ชื่อเต็ม'),iPre=rh.indexOf('คำนำหน้า'),iFn=rh.indexOf('ชื่อ'),iLn=rh.indexOf('นามสกุล'),iO=rh.indexOf('สังกัด/หน่วยงาน'),iF=rh.indexOf('ยอดรวม');
      var i17=rh.indexOf('อาหาร_27พค'),i18=rh.indexOf('อาหาร_28พค');
      rd.slice(1).forEach(function(r){
        var fullNm=String(r[iN]||'').trim();if(!fullNm)fullNm=((String(r[iPre]||'')+' '+String(r[iFn]||'')+' '+String(r[iLn]||'')).trim().replace(/\s+/g,' '));
        regMap[String(r[0])]={name:fullNm,org:String(r[iO]||''),fee:Number(r[iF]||0),d17:Number(r[i17]||0),d18:Number(r[i18]||0)};
      });
    }
    var pd=psh.getDataRange().getValues();
    var slips=pd.slice(1).filter(function(r){return String(r[6]||'')==='รอตรวจสอบ';}).map(function(r){
      var rid=String(r[0]||''),info=regMap[rid]||{};
      var ts=r[1]?Utilities.formatDate(new Date(r[1]),'Asia/Bangkok','dd/MM/yyyy HH:mm'):'';
      return{regId:rid,name:info.name||'',org:info.org||'',amount:Number(r[2]||0),
             ref:String(r[4]||''),slipUrl:String(r[5]||''),paidAt:ts,
             fee:info.fee||0,d17:info.d17||0,d18:info.d18||0};
    });
    return{success:true,slips:slips};
  }catch(e){return{success:false,message:e.message};}
}

// ================================================================
//  PATCH: _approvePayment + _sendPaymentConfirmEmail
//  วิธีใช้: เปิด Apps Script > แทนที่ฟังก์ชัน _approvePayment เดิม
//           และเพิ่มฟังก์ชัน _sendPaymentConfirmEmail ใหม่ด้านล่าง
// ================================================================


// ── แทนที่ฟังก์ชัน _approvePayment เดิมทั้งหมด ─────────────────

function _approvePayment(regId, approvedBy) {
  regId = String(regId || '').trim();
  if (!regId) return { success: false, message: 'ไม่ระบุ REG ID' };
  var by  = approvedBy || 'staff';
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var lock = LockService.getScriptLock();
  try { lock.tryLock(15000); } catch (le) {}
  if (!lock.hasLock()) return { success: false, message: 'ระบบไม่ว่าง กรุณาลองใหม่อีกครั้ง' };
  var info = null;
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    info = _getRegInfo(ss, regId);
    if (!info) return { success: false, message: 'ไม่พบ REG ID' };

    // ── 1. อัป PAYMENTS sheet (เฉพาะสลิปที่รอตรวจสอบ) ──────────
    var hit = false;
    var psh = ss.getSheetByName(SH_PAY);
    if (psh && psh.getLastRow() > 1) {
      var pd = psh.getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        if (String(pd[i][0]).trim() === regId && String(pd[i][6]) === 'รอตรวจสอบ') {
          psh.getRange(i + 1, 7, 1, 3).setValues([['ชำระแล้ว', by, now]]);
          hit = true;
        }
      }
      SpreadsheetApp.flush();
    }
    if (!hit) {
      return { success: false, message: info.status === 'ชำระแล้ว' ? 'REG ID นี้อนุมัติไปแล้ว' : 'ไม่พบสลิปที่รอตรวจสอบของ REG ID นี้' };
    }

    // ── 2. อัป REGISTRATIONS + competition sheets ──────────────
    _updatePayStatus(ss, regId, 'ชำระแล้ว');
    _clearMealCache();
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }

  // ── 3. จองคิวอีเมล (ใช้ข้อมูลจาก _getRegInfo รอบเดียว ไม่อ่านชีทซ้ำ) ──
  var emailQueued = false, emailErr = '';
  try {
    if (info.email) {
      _queueEmail('approve', regId, { name: info.name, org: info.org, typeLbl: info.typeLbl, email: info.email, fee: info.fee, d17: info.d17, d18: info.d18, approvedAt: now });
      emailQueued = true;
    } else {
      Logger.log('APPROVE EMAIL SKIP [' + regId + ']: email is empty');
    }
  } catch (me) {
    emailErr = me.message;
    Logger.log('APPROVE EMAIL QUEUE ERROR [' + regId + ']: ' + me.message);
  }

  // ── Auto-export: อัปเดต 4 ชีทสรุป/รายชื่ออาหาร ให้ตรงหลังอนุมัติ ──
  try{ _scheduleAutoExports(); }catch(autoErr){ Logger.log('schedule auto export err: '+autoErr.message); }

  return { success: true, emailQueued: emailQueued, emailErr: emailErr };
}


// ── เพิ่มฟังก์ชันใหม่นี้ (วางไว้ใต้ _approvePayment) ──────────

function _sendPaymentConfirmEmail(fullName, org, typeLbl, email, regId, fee, d17, d18, approvedAt) {
  var pdfAtt  = _printPdfBlob(regId); // PDF แนบ (null = ไม่แนบ) — เรียกจากคิว/ส่งซ้ำเท่านั้น
  var url     = (_siteUrl()?_siteUrl()+'/#st':_webAppUrl());
  var total   = Number(fee).toLocaleString();
  var dateStr = approvedAt || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var eName = _esc(fullName), eOrg = _esc(org), eType = _esc(typeLbl), eRid = _esc(regId); // กัน HTML injection (plain-text body ใช้ค่าเดิม)

  // Meal badge (เดิม) — ใช้เมื่อสร้างคูปองไม่ได้
  var couponHtml = _couponHtmlFromSheets(regId, fullName, d17, d18);
  var mealHtml = '';
  if (!couponHtml && (d17 > 0 || d18 > 0)) {
    mealHtml =
      '<div style="background:#f0fdf4;border-left:4px solid #16a34a;padding:14px 16px;margin:14px 0;border-radius:0 8px 8px 0">' +
        '<p style="margin:0 0 8px;font-weight:700;color:#15803d;font-size:13px">สิทธิ์อาหาร / Meal Entitlement</p>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
          '<tr style="background:#dcfce7">' +
            '<th style="padding:6px 10px;text-align:left;color:#166534">วัน / Date</th>' +
            '<th style="padding:6px 10px;text-align:center;color:#166534">จำนวน / Qty</th>' +
          '</tr>' +
          (d17 > 0
            ? '<tr><td style="padding:6px 10px;border-top:1px solid #eee">27 พฤษภาคม (May 27)</td>' +
              '<td style="padding:6px 10px;border-top:1px solid #eee;text-align:center;font-weight:700;color:#0891b2">' + d17 + ' คน</td></tr>'
            : '') +
          (d18 > 0
            ? '<tr><td style="padding:6px 10px;border-top:1px solid #eee">28 พฤษภาคม (May 28)</td>' +
              '<td style="padding:6px 10px;border-top:1px solid #eee;text-align:center;font-weight:700;color:#b45309">' + d18 + ' คน</td></tr>'
            : '') +
        '</table>' +
      '</div>';
  }

  couponHtml = couponHtml || mealHtml;
  couponHtml = _schedHtmlFromSheets(regId) + couponHtml;
  var html =
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%"><tr><td align="center"><table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;font-family:Sarabun,Arial,sans-serif;color:#1f2937;border:1px solid #e3d3a8;border-radius:8px;border-collapse:separate;overflow:hidden"><tr><td style="text-align:left">' +

    // ── Header ──
    '<div style="background:#fbf6e9;padding:20px 20px;border-bottom:4px solid #b8862b">' +
      '<p style="margin:0 0 4px;font-size:9px;color:#b8862b;letter-spacing:.1em;text-transform:uppercase">Payment Confirmed</p>' +
      '<h1 style="margin:0 0 3px;font-size:15px;font-weight:700;color:#0b5844">' + CONF_NAME_TH + '</h1>' +
      '<p style="margin:0 0 2px;font-size:12px;color:#0b5844">' + CONF_NAME_EN + '</p>' +
      '<p style="margin:0 0 12px;font-size:11px;color:#6b5a3a">' + CONF_DATE_TH + ' · ' + CONF_VENUE_TH + '</p>' +
      '<div style="background:#fff;border:1px solid #e3d3a8;border-radius:6px;padding:8px 12px">' +
        '<p style="margin:0 0 2px;font-size:9px;color:#b8862b;letter-spacing:.08em;text-transform:uppercase">REG ID</p>' +
        '<p style="margin:0;font-family:\'Courier New\',monospace;font-size:14px;font-weight:700;color:#b8862b;letter-spacing:.08em;word-break:break-all">' + eRid + '</p>' +
      '</div>' +
    '</div>' +

    // ── Body ──
    '<div style="padding:20px 20px;background:#fff">' +
      '<p style="font-size:14px;margin:0 0 14px">เรียน / Dear <b>' + eName + '</b>,</p>' +

      // ✅ Success banner
      '<div style="background:#f0fdf4;border:2px solid #16a34a;border-radius:8px;padding:16px 20px;margin:0 0 18px;text-align:center">' +
        '<p style="margin:0 0 4px;font-size:22px">✅</p>' +
        '<p style="margin:0 0 4px;font-size:16px;font-weight:700;color:#15803d">ยืนยันการชำระเงินเรียบร้อยแล้ว</p>' +
        '<p style="margin:0;font-size:13px;color:#166534">Payment has been confirmed successfully</p>' +
      '</div>' +

      // Info table
      '<div style="background:#f6efe1;border:1px solid #e3d3a8;border-radius:8px;padding:14px 14px;margin:0 0 16px">' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
          '<tr><td style="color:#64748b;padding:5px 0;width:38%">ชื่อ / Name</td>' +
              '<td style="font-weight:600;padding:5px 0">' + eName + '</td></tr>' +
          '<tr><td style="color:#64748b;padding:5px 0">สังกัด / Org</td>' +
              '<td style="padding:5px 0">' + eOrg + '</td></tr>' +
          '<tr><td style="color:#64748b;padding:5px 0">ประเภท / Type</td>' +
              '<td style="padding:5px 0">' + eType + '</td></tr>' +
          '<tr><td style="color:#64748b;padding:5px 0">ยอดชำระ / Amount</td>' +
              '<td style="font-weight:700;color:#15803d;font-size:15px;padding:5px 0">' + total + ' บาท (THB)</td></tr>' +
          '<tr><td style="color:#64748b;padding:5px 0">วันที่อนุมัติ / Approved</td>' +
              '<td style="font-family:monospace;padding:5px 0">' + dateStr + '</td></tr>' +
          '<tr><td style="color:#64748b;padding:5px 0">สถานะ / Status</td>' +
              '<td style="padding:5px 0"><span style="background:#dcfce7;color:#15803d;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:700">ชำระแล้ว ✓</span></td></tr>' +
        '</table>' +
      '</div>' +

      couponHtml +

      // Notice
      '<div style="background:#fbf6ec;border-left:4px solid #b8862b;padding:12px 14px;margin:14px 0;border-radius:0 8px 8px 0;font-size:12px;color:#6b5a3a">' +
        '<b>หมายเหตุ:</b> กรุณาแสดงอีเมลนี้หรือ REG ID ในวันงาน<br>' +
        '<b>Note:</b> Please present this email or your REG ID on the event day.' +
      '</div>' +

      // Button
      '<div style="text-align:center;margin-top:20px">' +
        '<a href="' + url + '" style="display:inline-block;background:#0b5844;color:#fff;padding:12px 28px;text-decoration:none;font-weight:700;font-size:14px;border-radius:6px">' +
          'ตรวจสอบสถานะ / Check Status →' +
        '</a>' +
      '</div>' +
      (pdfAtt ? _printAttachNote() : '') +
      _printButtonHtml(regId) +
    '</div>' +

    // ── Footer ──
    '<div style="background:#f6efe1;padding:12px 20px;font-size:11px;color:#6b5a3a;border-top:1px solid #eee;text-align:center">' +
      CONF_VENUE_TH + '<br>' +
      CONF_VENUE_EN + '<br>' +
      'อีเมลนี้ออกโดยระบบอัตโนมัติ · Auto-generated — please do not reply' +
    '</div>' +

    '</td></tr></table></td></tr></table>';

  var mailOpt = { name: CONF_NAME_TH, htmlBody: html };
  if (pdfAtt) mailOpt.attachments = [pdfAtt];
  GmailApp.sendEmail(
    email,
    '[ยืนยันการชำระเงิน / Payment Confirmed] Nursing & Innovation 2027 — ' + regId,
    'เรียน ' + fullName + '\n\nการชำระเงิน REG ID: ' + regId + ' ได้รับการยืนยันแล้ว\nยอดชำระ: ' + total + ' บาท\nวันที่: ' + dateStr + '\n\nขอบคุณที่ลงทะเบียน ' + CONF_NAME_TH,
    mailOpt
  );
}


// (testPaymentConfirmEmail ย้ายไปรวมกับตัวที่ใช้งานจริงด้านล่าง — เดิมประกาศซ้ำ 2 ที่)

function _rejectPayment(regId, rejectedBy){
  regId=String(regId||'').trim();
  if(!regId) return{success:false,message:'ไม่ระบุ REG ID'};
  var by=rejectedBy||'staff';
  var lock=LockService.getScriptLock();
  try{ lock.tryLock(15000); }catch(le){}
  if(!lock.hasLock()) return{success:false,message:'ระบบไม่ว่าง กรุณาลองใหม่อีกครั้ง'};
  var em='',nm='';
  try{
    var ss=SpreadsheetApp.openById(SS_ID);
    var info=_getRegInfo(ss,regId);
    if(!info) return{success:false,message:'ไม่พบ REG ID'};
    if(info.status==='ชำระแล้ว') return{success:false,message:'REG ID นี้ชำระแล้ว ไม่สามารถปฏิเสธได้'};
    var hit=false;
    var psh=ss.getSheetByName(SH_PAY);
    if(psh&&psh.getLastRow()>1){
      var pd=psh.getDataRange().getValues();
      for(var i=1;i<pd.length;i++){
        if(String(pd[i][0]).trim()===regId&&String(pd[i][6])==='รอตรวจสอบ'){
          // G=สถานะ, H=ผู้อนุมัติ, I=วันอนุมัติ
          psh.getRange(i+1,7,1,3).setValues([['ปฏิเสธ',by,Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm')]]);
          hit=true;
        }
      }
    }
    if(!hit) return{success:false,message:'ไม่พบสลิปที่รอตรวจสอบของ REG ID นี้'};
    _updatePayStatus(ss,regId,'รอชำระเงิน');
    _clearMealCache();
    em=info.email; nm=info.name;
  }catch(e){
    return{success:false,message:e.message};
  }finally{
    lock.releaseLock();
  }
  var emailQueued=false;
  if(em){
    try{ _queueEmail('reject',regId,{email:em,name:nm}); emailQueued=true; }
    catch(me){ Logger.log('REJECT EMAIL QUEUE ERROR ['+regId+']: '+me.message); }
  }
  // ── Auto-export: อัปเดต 4 ชีทสรุป/รายชื่ออาหาร ให้ตรงหลังปฏิเสธ ──
  try{ _scheduleAutoExports(); }catch(autoErr){ Logger.log('schedule auto export err: '+autoErr.message); }
  return{success:true,emailQueued:emailQueued};
}


// ── SECTION 7: MEAL CHECK-IN ────────────────────────────────

function _getMealCheckIn(p) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'mealCheckIn';
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SH_CHECKIN);
    if (!sh || sh.getLastRow() < 2) {
      return { success: true, rows: [], total17: 0, total18: 0 };
    }
    var lastRow = sh.getLastRow();
    var data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
    var total17 = 0, total18 = 0;
    var rows = data.map(function(r) {
      var dayLbl = String(r[3] || '');
      var cnt = Number(r[4]) || 0;
      if (dayLbl.indexOf('27') >= 0) total17 += cnt;
      else if (dayLbl.indexOf('28') >= 0) total18 += cnt;
      var nm = String(r[1] || '').trim() || 'ไม่ระบุชื่อ';
      return { rid: String(r[0]), name: nm, org: String(r[2]),
               dayLabel: dayLbl, count: cnt,
               receivedBy: String(r[5]), by: String(r[6]), time: String(r[7]) };
    });
    var result = { success: true, rows: rows, total17: total17, total18: total18 };
    try { cache.put(cacheKey, JSON.stringify(result), 60); } catch(e) {}
    return result;
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function _mealCheckIn(p) {
  var rid      = String(p.regId  || '').trim();
  var day      = String(p.day    || '').trim(); // '17' or '18'
  var name     = String(p.name   || '').trim();
  var org      = String(p.org    || '').trim();
  var cnt      = Number(p.count  || 1);
  var recvBy   = String(p.receivedBy || '').trim();
  var by       = String(p.by     || 'staff').trim();
  var giver    = String(p.giver  || '').trim();
  if (giver) by += ' (แจก: ' + giver + ')';

  if (!rid || !day) return { success: false, message: 'ข้อมูลไม่ครบ' };
  if (day !== '17' && day !== '18') return { success: false, message: 'วันไม่ถูกต้อง' };
  if (!isFinite(cnt) || Math.floor(cnt) !== cnt || cnt < 1) return { success: false, message: 'จำนวนไม่ถูกต้อง' };

  var lock = LockService.getScriptLock();
  try { lock.tryLock(15000); } catch (le) {}
  if (!lock.hasLock()) return { success: false, message: 'ระบบไม่ว่าง กรุณาลองใหม่อีกครั้ง' };
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var sh  = ss.getSheetByName(SH_CHECKIN);
    if (!sh) return { success: false, message: 'ไม่พบ MEAL_CHECKIN sheet' };

    var info = _getRegInfo(ss, rid);
    if (!info) return { success: false, message: 'ไม่พบ REG ID' };
    var quota = day === '17' ? info.d17 : info.d18;
    if (quota < 1) return { success: false, message: 'REG ID นี้ไม่มีสิทธิ์อาหารวันที่ ' + day };
    if (cnt > quota) return { success: false, message: 'จำนวนเกินสิทธิ์ (สิทธิ์ ' + quota + ' คน)' };

    // ตรวจซ้ำ
    if (sh.getLastRow() > 1) {
      var existing = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === rid &&
            String(existing[i][3]).indexOf(day==='17'?'27':'28') >= 0) {
          return { success: false, duplicate: true, message: 'เช็คอินซ้ำ' };
        }
      }
    }

    var now      = new Date();
    var dayLabel = day === '17'
      ? 'วันที่ 27 พฤษภาคม 2570'
      : 'วันที่ 28 พฤษภาคม 2570';

    sh.appendRow([rid, _safe(name), _safe(org), dayLabel, cnt, _safe(recvBy), _safe(by),
      Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')]);
    SpreadsheetApp.flush();
    _clearMealCache(); // clear GAS cache หลัง check-in
    _updateMealSummaryCheckin(rid, day); // อัป MEAL_SUMMARY ทันที

    return { success: true, rid: rid, day: day, name: name, count: cnt };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _getMealSummary(){
  try{
    // ── GAS Cache (60 วิ) ──
    var cache=CacheService.getScriptCache();
    var cached=cache.get('mealSummary');
    if(cached){
      try{ return JSON.parse(cached); }catch(e){}
    }

    var ss=SpreadsheetApp.openById(SS_ID);
    var rsh=ss.getSheetByName(SH_REG);
    if(!rsh||rsh.getLastRow()<2) return{success:true,d17:0,d18:0,rows:[]};

    var lastRow=rsh.getLastRow();
    var hdrRow=rsh.getRange(1,1,1,rsh.getLastColumn()).getValues()[0];
    var iN=hdrRow.indexOf('ชื่อเต็ม');
    var iPre=hdrRow.indexOf('คำนำหน้า'),iFn=hdrRow.indexOf('ชื่อ'),iLn=hdrRow.indexOf('นามสกุล');
    var iO=hdrRow.indexOf('สังกัด/หน่วยงาน');
    var iCard=hdrRow.indexOf('เลขบัตรประชาชน');
    var i17=hdrRow.indexOf('อาหาร_27พค'),i18=hdrRow.indexOf('อาหาร_28พค');
    var iRid=0;

    var maxCol=Math.max(iN+1,iPre+1,iFn+1,iLn+1,iO+1,iCard+1,i17+1,i18+1);
    var data=rsh.getRange(2,1,lastRow-1,maxCol).getValues();

    var t17=0,t18=0,rows=[];
    data.forEach(function(r){
      var d17=Number(r[i17]||0),d18=Number(r[i18]||0);
      t17+=d17;t18+=d18;
      if(d17||d18){
        var pre=String(r[iPre]||'').trim();
        var fn=String(r[iFn]||'').trim();
        var ln=String(r[iLn]||'').trim();
        var fullNm=String(r[iN]||'').trim();
        if(!fullNm) fullNm=((pre+' '+fn+' '+ln).trim().replace(/\s+/g,' '));
        rows.push({
          rid:String(r[iRid]),
          prefix:pre,fname:fn,lname:ln,name:fullNm,
          org:String(r[iO]||''),
          idcard:String(r[iCard]||''),
          d17:d17,d18:d18,got17:0,got18:0
        });
      }
    });

    // MEAL_CHECKIN — อ่านเฉพาะ col A(REG_ID), D(วัน), E(จำนวน)
    var csh=ss.getSheetByName(SH_CHECKIN);
    var checked17={},checked18={};
    if(csh&&csh.getLastRow()>1){
      var ciData=csh.getRange(2,1,csh.getLastRow()-1,5).getValues();
      ciData.forEach(function(cr){
        var rid=String(cr[0]),dayLbl=String(cr[3]||''),cnt=Number(cr[4]||0);
        if(dayLbl.indexOf('27')>=0) checked17[rid]=(checked17[rid]||0)+cnt;
        else if(dayLbl.indexOf('28')>=0) checked18[rid]=(checked18[rid]||0)+cnt;
      });
    }
    rows.forEach(function(r){
      r.got17=checked17[r.rid]||0;
      r.got18=checked18[r.rid]||0;
    });
    var recv17=Object.keys(checked17).reduce(function(s,k){return s+checked17[k];},0);
    var recv18=Object.keys(checked18).reduce(function(s,k){return s+checked18[k];},0);

    var result={success:true,d17:t17,d18:t18,
      recv17:recv17,recv18:recv18,
      remain17:t17-recv17,remain18:t18-recv18,rows:rows};

    // เก็บ cache 60 วิ
    try{ cache.put('mealSummary',JSON.stringify(result),60); }catch(e){}
    return result;
  }catch(e){return{success:false,message:e.message};}
}

function _clearMealCache(){
  try{
    var c=CacheService.getScriptCache();
    c.remove('mealSummary');
    c.remove('mealCheckIn');
  }catch(e){}
}
function _sendConfirmEmail(fullName,org,typeLbl,typeLblEn,email,regId,fee,isFree,d,meal){
  // ไม่ครอบ try/catch ทั้งหมด — ให้ error โยนออกไปให้ _registerUser จัดการ
    var pdfAtt=_printPdfBlob(regId); // PDF แนบ (null = ไม่แนบ) — เรียกจากคิว/ส่งซ้ำเท่านั้น
    var url=(_siteUrl()?_siteUrl()+'/#st':_webAppUrl());
    var total=Number(fee).toLocaleString();
    var dateStr=Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm');
    var eName=_esc(fullName), eOrg=_esc(org), eType=_esc(typeLbl), eTypeEn=_esc(typeLblEn), eRid=_esc(regId); // กัน HTML injection

    // Meal table rows — กรอง Online (mealDay=none) ออก
    var mealRowsHtml='';
    var hasMeal=false;
    (meal.summaryRows||[]).forEach(function(r){
      if(!r.d17&&!r.d18) return; // ไม่มีสิทธิ์อาหาร ข้าม
      hasMeal=true;
      var d17c=r.d17?'<span style="color:#15803d;font-weight:700">✓</span>':'<span style="color:#ccc">—</span>';
      var d18c=r.d18?'<span style="color:#b45309;font-weight:700">✓</span>':'<span style="color:#ccc">—</span>';
      mealRowsHtml+='<tr><td style="padding:6px 10px;border-top:1px solid #eee;font-size:12px">'+_esc(r.nm)+'</td>'+
        '<td style="padding:6px 10px;border-top:1px solid #eee;text-align:center;font-size:13px;color:#555">'+r.count+'</td>'+
        '<td style="padding:6px 10px;border-top:1px solid #eee;text-align:center">'+d17c+'</td>'+
        '<td style="padding:6px 10px;border-top:1px solid #eee;text-align:center">'+d18c+'</td></tr>';
    });
    var compRowsHtml='';
    if(!isFree&&(d.comps||[]).length){
      (d.comps||[]).forEach(function(cid){
        var c=COMP[cid];if(!c)return;
        var days=(d.compDays||{})[cid]||1;
        var f=days===2&&c.d2?c.d2:c.d1;
        compRowsHtml+='<tr><td style="padding:6px 10px;border-top:1px solid #eee;font-size:12px">'+_esc(c.nm)+'<div style="font-size:11px;color:#6b7280;margin-top:2px">'+_esc(_schedLine(cid))+'</div></td>'+
          '<td style="padding:6px 10px;border-top:1px solid #eee;text-align:right;font-weight:600">'+f.toLocaleString()+' ฿</td></tr>';
      });
      compRowsHtml+='<tr style="background:#f8f8f8"><td style="padding:8px 10px;font-weight:700">Total / รวม</td>'+
        '<td style="padding:8px 10px;text-align:right;font-weight:700;color:#0b5844;font-size:15px">'+total+' ฿</td></tr>';
    }

    var couponBlk='';
    try{
      var mc=_mealCoupons(_couponListFromPayload(d,isFree),isFree,fullName);
      if(mc.d17.length||mc.d18.length) couponBlk=_couponHtml(regId,mc,isFree);
    }catch(ce){ Logger.log('COUPON ERROR ['+regId+']: '+ce.message); }
    var payBlock=isFree
      ?'<p style="color:#15803d;font-weight:700">✅ ไม่มีค่าธรรมเนียม / No registration fee required</p>'+
         '<p style="font-size:12px;color:#555;margin:4px 0 12px">ลงทะเบียนในฐานะกรรมการจากสถาบันร่วมจัด (ไม่มีค่าลงทะเบียน)</p>'
      :'<div style="background:#fffbeb;border-left:4px solid #d97706;padding:14px 16px;margin:12px 0;border-radius:0 8px 8px 0">'+
         '<p style="font-weight:700;color:#92400e;margin:0 0 8px">ยอดที่ต้องชำระ / Amount Due: <span style="font-size:18px">'+total+' บาท (THB)</span></p>'+
         '<table style="font-size:13px;width:100%;border-collapse:collapse">'+
           '<tr><td style="color:#666;width:38%;padding:3px 0">ธนาคาร / Bank</td><td>ธนาคารกรุงไทย (Krungthai Bank)</td></tr>'+
           '<tr><td style="color:#666;padding:3px 0">เลขบัญชี / Account</td><td style="font-family:monospace;font-size:15px;font-weight:700">986-5-34173-5</td></tr>'+
           '<tr><td style="color:#666;padding:3px 0">ชื่อบัญชี / Name</td><td>มหาวิทยาลัยราชภัฏนครปฐม บริการทางวิชาการ</td></tr>'+
         '</table>'+
         '<p style="margin:8px 0 0;font-size:11px;color:#92400e">กรุณาชำระภายใน <b>7 วัน</b> นับจากวันลงทะเบียน / Please pay within <b>7 days</b> of registration</p>'+
       '</div>';

    var html=
      '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%"><tr><td align="center"><table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;font-family:Sarabun,Arial,sans-serif;color:#1f2937;border:1px solid #e3d3a8;border-radius:8px;border-collapse:separate;overflow:hidden"><tr><td style="text-align:left">'+
      '<div style="background:#fbf6e9;padding:20px 20px;border-bottom:4px solid #b8862b">'+
        '<p style="margin:0 0 4px;font-size:9px;color:#b8862b;letter-spacing:.1em;text-transform:uppercase">Registration Confirmation</p>'+
        '<h1 style="margin:0 0 3px;font-size:15px;font-weight:700;color:#0b5844">'+CONF_NAME_TH+'</h1>'+
        '<p style="margin:0 0 2px;font-size:12px;color:#0b5844">'+CONF_NAME_EN+'</p>'+
        '<p style="margin:0 0 12px;font-size:11px;color:#6b5a3a">'+CONF_DATE_TH+' · '+CONF_VENUE_TH+'</p>'+
        '<div style="background:#fff;border:1px solid #e3d3a8;border-radius:6px;padding:8px 12px">'+
          '<p style="margin:0 0 2px;font-size:9px;color:#b8862b;letter-spacing:.08em;text-transform:uppercase">REG ID</p>'+
          '<p style="margin:0;font-family:\'Courier New\',monospace;font-size:14px;font-weight:700;color:#b8862b;letter-spacing:.08em;word-break:break-all">'+eRid+'</p>'+
        '</div>'+
      '</div>'+
      '<div style="padding:20px 20px;background:#fff">'+
        '<p style="font-size:14px;margin:0 0 14px">เรียน / Dear <b>'+eName+'</b>,</p>'+
        '<p style="font-size:13px;color:#555;margin:0 0 16px;line-height:1.7">'+
          'ระบบได้รับข้อมูลการลงทะเบียนของท่านเรียบร้อยแล้ว<br>'+
          'Your registration for <b>Nursing & Innovation 2027</b> has been successfully received.'+
        '</p>'+
        '<div style="background:#fbf6ec;border:1.5px solid #b8862b;border-radius:8px;padding:16px;text-align:center;margin:0 0 18px">'+
          '<p style="margin:0 0 4px;font-size:9px;color:#b8862b;letter-spacing:.1em;text-transform:uppercase">Registration ID</p>'+
          '<p style="margin:0;font-family:\'Courier New\',monospace;font-size:24px;font-weight:700;color:#0b5844;letter-spacing:.1em;word-break:break-all">'+eRid+'</p>'+
          '<p style="margin:6px 0 0;font-size:11px;color:#888">Registered '+dateStr+'</p>'+
        '</div>'+
        '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:14px">'+
          '<tr style="background:#f6efe1"><td style="padding:8px 10px;color:#555;width:38%">ประเภท / Type</td><td style="padding:8px 10px">'+eType+' <span style="color:#888;font-size:12px">('+eTypeEn+')</span></td></tr>'+
          '<tr><td style="padding:8px 10px;color:#555;border-top:1px solid #eee">สังกัด / Affiliation</td><td style="padding:8px 10px;border-top:1px solid #eee">'+eOrg+'</td></tr>'+
        '</table>'+
        (compRowsHtml?'<p style="font-size:10px;color:#888;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px">รายการที่ลงทะเบียน / Events</p>'+
          '<table style="width:100%;border-collapse:collapse;margin-bottom:14px">'+compRowsHtml+'</table>':'')+
        couponBlk+
        payBlock+
        '<div style="text-align:center;margin-top:20px">'+
          '<a href="'+url+'" style="display:inline-block;background:#0b5844;color:#fff;padding:12px 28px;text-decoration:none;font-weight:700;font-size:14px;border-radius:6px">'+
            'ตรวจสอบสถานะและแนบสลิป / Check Status &amp; Upload Slip →</a>'+
        '</div>'+
        (pdfAtt?_printAttachNote():'')+
        _printButtonHtml(regId)+
      '</div>'+
      '<div style="background:#f6efe1;padding:12px 20px;font-size:11px;color:#6b5a3a;border-top:1px solid #eee;text-align:center">'+
        'คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏนครปฐม<br>'+
        'Faculty of Nursing, Nakhon Pathom Rajabhat University<br>'+
        'อีเมลนี้ออกโดยระบบอัตโนมัติ · Auto-generated — please do not reply'+
      '</div>'+
      '</td></tr></table></td></tr></table>';

    // ส่งอีเมลหลัก → ผู้ลงทะเบียน
    var mailOpt={ name: MAIL_SENDER_NAME, htmlBody:html,name:CONF_NAME_EN};
    if(pdfAtt) mailOpt.attachments=[pdfAtt];
    GmailApp.sendEmail(email,
      '[Registration Confirmed / ยืนยันการลงทะเบียน] Nursing & Innovation 2027 — '+regId,
      'Dear '+fullName+'\nREG ID: '+regId+'\nAmount: '+(isFree?'Free':total+' THB')+
      '\nMeal 27 May: '+meal.d17Count+' person(s) | 28 May: '+meal.d18Count+' person(s)'+
      '\nCheck status: '+url,
      mailOpt);

  // ─────────────────────────────────────────────────────────
  // ลบ team email logic (ไม่มี email field ในฟอร์มสมาชิก)
}

// ── SECTION 10: COMPETITION SHEETS ──────────────────────────

function _writeCompSheets(ss, regId, d, fullName, org, totalFee) {
  var regDate = new Date();

  (d.comps||[]).forEach(function(cid) {
    var cfg = COMP[cid]; if(!cfg) return;
    var shName = COMP_SHEET[cid]; if(!shName) return;

    var days = (d.compDays||{})[cid] || 1;
    var fee  = days===2&&cfg.d2 ? cfg.d2 : cfg.d1;

    var sh = _compSheet(ss, shName, cfg.nm, cfg.nm_en, cid==='c00' ? COMP_ONLINE_HDR : null);
    var teamNo = Math.max(sh.getLastRow() - 1, 1); // แถว 1-2 = title + header

    // ── c00 Online — 8 cols เท่านั้น (ไม่มีทีม) ──────────────
    if(cid==='c00'){
      sh.appendRow([
        regId, regDate, _safe(fullName), _safe(org),
        _safe(d.email||''), _safe(d.phone||''), 'รอชำระเงิน', regDate
      ]);
      return;
    }

    // ── competition ทั่วไป — 36 cols ─────────────────────────
    var teamNameVal = _safe((d.teamName||{})[cid] || '');
    var slots=[];
    for(var si=0;si<4;si++) slots.push({pre:'',fn:'',ln:'',id:'',lic:'',al:''});

    slots[0].pre=_safe((d.prefix||'').trim());
    slots[0].fn=_safe((d.fname||'').trim());
    slots[0].ln=_safe((d.lname||'').trim());
    slots[0].id=_safe(d.idcard||'');
    slots[0].lic=_safe(d.license||'');
    slots[0].al=_safe(d.allergy||'');

    var team=(d.members||{})[cid]||[];
    for(var i=1;i<4;i++){
      var m=team[i];
      if(m&&(m.fname||m.name||'').trim()){
        slots[i].pre=_safe((m.prefix||'').trim());
        slots[i].fn=_safe((m.fname||m.name||'').trim());
        slots[i].ln=_safe((m.lname||'').trim());
        slots[i].id=_safe(m.idcard||'');
        slots[i].lic=_safe(m.license||'');
        slots[i].al=_safe(m.allergy||'');
      }
    }

    sh.appendRow([
      teamNo, regId, regDate, teamNameVal,
      _safe(fullName), _safe(org), _safe(d.email||''), _safe(d.phone||''),
      days, fee,
      slots[0].pre, slots[0].fn, slots[0].ln, slots[0].id, slots[0].lic, slots[0].al,
      slots[1].pre, slots[1].fn, slots[1].ln, slots[1].id, slots[1].lic, slots[1].al,
      slots[2].pre, slots[2].fn, slots[2].ln, slots[2].id, slots[2].lic, slots[2].al,
      slots[3].pre, slots[3].fn, slots[3].ln, slots[3].id, slots[3].lic, slots[3].al,
      'รอชำระเงิน', regDate
    ]);
  });
}

// สร้างหรือดึง competition sheet พร้อม header
var COMP_ONLINE_HDR = ['REG_ID','วันที่ลงทะเบียน','ชื่อ-นามสกุล','สังกัด','อีเมล','โทรศัพท์','สถานะชำระ','วันที่บันทึก'];
// ตั้งความกว้างคอลัมน์ — รวมคอลัมน์ติดกันที่กว้างเท่ากันเป็นคำสั่งเดียว (ลดจำนวน call)
function _setWidths(sh, widths) {
  var i = 0;
  while (i < widths.length) {
    var j = i;
    while (j + 1 < widths.length && widths[j + 1] === widths[i]) j++;
    if (widths[i]) sh.setColumnWidths(i + 1, j - i + 1, widths[i]);
    i = j + 1;
  }
}
function _compSheet(ss, shName, nmTH, nmEN, hdr) {
  var sh = ss.getSheetByName(shName);
  if(!sh){
    var H = hdr || COMP_SHEET_HDR;
    sh = ss.insertSheet(shName);
    // ชีทใหม่มีแค่ 26 คอลัมน์ × 1000 แถว — ขยายก่อนเขียน header 36 คอลัมน์/format ถึงแถว 1002
    if (sh.getMaxColumns() < H.length) sh.insertColumnsAfter(sh.getMaxColumns(), H.length - sh.getMaxColumns());
    if (sh.getMaxRows() < 1002) sh.insertRowsAfter(sh.getMaxRows(), 1002 - sh.getMaxRows());
    // Title row (row 1) — ชื่อรายการ
    sh.getRange(1,1,1,H.length)
      .merge()
      .setValue(nmTH + ' · ' + nmEN)
      .setBackground('#0f1923')
      .setFontColor('#fbbf24')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center');
    // Header row (row 2)
    var hdrRange = sh.getRange(2,1,1,H.length);
    hdrRange.setValues([H])
      .setBackground('#162544')
      .setFontColor('#94a3b8')
      .setFontWeight('bold')
      .setFontSize(10);
    sh.setFrozenRows(2);
    // Column widths ตาม layout จริง (Online 8 คอลัมน์ / ทีม 36 คอลัมน์)
    _setWidths(sh, hdr ? [130,130,180,200,180,100,110,130]
      : [60,120,130,160,160,160,160,100,70,110,
         70,100,100,110,90,110, 70,100,100,110,90,110,
         70,100,100,110,90,110, 70,100,100,110,90,110,
         100,130]);
    if(hdr){
      // Online (8 cols): date = col 2 และ col 8
      sh.getRange(3,2,1000,1).setNumberFormat('dd/MM/yyyy HH:mm');
      sh.getRange(3,8,1000,1).setNumberFormat('dd/MM/yyyy HH:mm');
    } else {
      // 36 cols: วันที่ลงทะเบียน = col 3, ค่าธรรมเนียม = col 10, วันที่บันทึก = col 36
      sh.getRange(3,3,1000,1).setNumberFormat('dd/MM/yyyy HH:mm');
      sh.getRange(3,36,1000,1).setNumberFormat('dd/MM/yyyy HH:mm');
      sh.getRange(3,10,1000,1).setNumberFormat('#,##0');
    }
  }
  return sh;
}

// อัปเดตสถานะชำระในชีทรายการทั้งหมด (เรียกจาก _updatePayStatus)
// รหัสรายการ (cid) ที่ REG_ID นี้ลงทะเบียน อ่านจากชีท COMPETITIONS รอบเดียว — null ถ้าไม่มีแถว (ข้อมูลเก่า/ฟรี)
function _regCids(ss, regId){
  var sh=ss.getSheetByName(SH_COMP);
  if(!sh||sh.getLastRow()<2) return null;
  var v=sh.getRange(1,1,sh.getLastRow(),Math.max(3,Math.min(sh.getLastColumn(),4))).getValues();
  var iR=v[0].indexOf('REG_ID'); if(iR<0) iR=0;
  var iC=v[0].indexOf('รหัสรายการ'); if(iC<0) iC=2;
  var id=String(regId||'').trim(), out=[];
  for(var i=1;i<v.length;i++){
    if(String(v[i][iR]).trim()===id){ var c=String(v[i][iC]||'').trim(); if(c&&COMP_SHEET[c]&&out.indexOf(c)<0) out.push(c); }
  }
  return out.length?out:null;
}
function _updateCompSheetsPayStatus(ss, regId, status) {
  // เปิดเฉพาะชีทของผู้ลงทะเบียนคนนี้ (เดิมวนอ่านทั้ง 13 ชีท) — ถ้าไม่มีแถวใน COMPETITIONS ใช้ลูปเดิมทั้งหมด
  var cids=_regCids(ss,regId);
  var wantOnline = !cids || cids.indexOf('c00')>=0;
  var onlineSh = wantOnline ? ss.getSheetByName('ประชุม_Online') : null;
  // อัปเดตสถานะชำระใน ประชุม_Online (col G = index 6)
  if(onlineSh && onlineSh.getLastRow()>1){
    var od = onlineSh.getRange(2,1,onlineSh.getLastRow()-1,7).getValues();
    for(var i=0;i<od.length;i++){
      if(String(od[i][0]).trim()===String(regId).trim()){
        onlineSh.getRange(i+2,7).setValue(status); // col G
      }
    }
  }

  // อัปเดตสถานะชำระใน competition sheets ทุกอัน
  // HDR34: REG_ID=col B(index 1), สถานะชำระ=col AI(index 34)
  var compSheetNames = [
    'ประชุม_Onsite','SIM_MAN','SIM_MOM','SIM_PED',
    'แนวปฏิบัติ_สถาบัน','แนวปฏิบัติ_สถานบริการ',
    'บทความวิจัย_บุคลากร','บทความวิจัย_นศ',
    'ตอบคำถาม_พยาบาล','ตอบคำถาม_กายวิภาค',
    'นวัตกรรม_นศ','นวัตกรรม_อาจารย์'
  ];
  if(cids){
    compSheetNames=cids.filter(function(c){return c!=='c00';}).map(function(c){return COMP_SHEET[c];});
  }
  compSheetNames.forEach(function(name){
    var sh=ss.getSheetByName(name);
    if(!sh||sh.getLastRow()<2) return;
    // header อยู่แถว 2 (แถว 1 = title merge) — รองรับชีทเก่าที่ header แถว 1 ด้วย
    var top2=sh.getRange(1,1,2,sh.getLastColumn()).getValues();
    var hdr=top2[0].indexOf('REG_ID')>=0 ? top2[0] : top2[1];
    // หา REG_ID col และ สถานะชำระ col จาก header
    var iRid=hdr.indexOf('REG_ID');
    var iStat=hdr.indexOf('สถานะชำระ');
    if(iRid<0) iRid=1;   // fallback col B
    if(iStat<0) iStat=34; // fallback col AI
    var data=sh.getRange(2,iRid+1,sh.getLastRow()-1,1).getValues();
    for(var i=0;i<data.length;i++){
      if(String(data[i][0]).trim()===String(regId).trim()){
        sh.getRange(i+2,iStat+1).setValue(status);
      }
    }
  });

  // อัปเดต MEAL_SUMMARY col I (สถานะชำระเงิน)
  var mealSh=ss.getSheetByName(SH_MEAL);
  if(mealSh&&mealSh.getLastRow()>1){
    var mhdr=mealSh.getRange(1,1,1,mealSh.getLastColumn()).getValues()[0];
    var miRid=mhdr.indexOf('REG_ID');
    var miStat=mhdr.indexOf('สถานะชำระเงิน');
    if(miRid<0) miRid=0;
    if(miStat<0) miStat=8; // fallback col I
    var mdata=mealSh.getRange(2,miRid+1,mealSh.getLastRow()-1,1).getValues();
    for(var i=0;i<mdata.length;i++){
      if(String(mdata[i][0]).trim()===String(regId).trim()){
        mealSh.getRange(i+2,miStat+1).setValue(status);
      }
    }
  }
}


// ── UTILITIES ────────────────────────────────────────────────

function _genRegId(ss){
  // เลขรันนิ่งต่อเนื่องทุกวัน (prefix MMdd เป็นแค่ตัวประกอบ) — ใช้ max suffix + 1 ไม่ใช้จำนวนแถว
  var sh=ss.getSheetByName(SH_REG);
  var max=0;
  if(sh&&sh.getLastRow()>1){
    sh.getRange(2,1,sh.getLastRow()-1,1).getValues().forEach(function(r){
      var m=String(r[0]).match(/-(\d{4,})$/);
      if(m) max=Math.max(max,parseInt(m[1],10));
    });
  }
  return REG_PREFIX+'-'+Utilities.formatDate(new Date(),'Asia/Bangkok','MMdd')+'-'+String(max+1).padStart(4,'0');
}

function _sheet(ss,name,headers){
  var sh=ss.getSheetByName(name);
  if(!sh){
    sh=ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── MANUAL TOOLS (run from GAS editor) ──────────────────────

function getMealReport(){
  var r=_getMealSummary();
  SpreadsheetApp.getUi().alert(
    'สรุปอาหาร Nursing & Innovation 2027\n\n'+
    '27 พฤษภาคม: '+r.d17+' คน\n'+
    '28 พฤษภาคม: '+r.d18+' คน\n\n'+
    'ดูรายละเอียดใน sheet MEAL_SUMMARY'
  );
}

// ================================================================
//  SETUP: สร้างชีทและคอลัมน์ทั้งหมด
//  วิธีใช้: เปิด GAS Editor → เลือก setupSheets → กด ▶ Run
//  ควรรันครั้งแรกก่อน deploy หรือเมื่อต้องการสร้างชีทใหม่
// ================================================================

// ================================================================
//  SETUP NEW INSTANCE: ติดตั้งระบบบนบัญชีใหม่ (เช่น it.nurse) ครั้งเดียว
//  1) สร้าง Google Sheet เปล่า → ส่วนขยาย → Apps Script
//  2) วาง Code.gs + ไฟล์ HTML ชื่อ Index → เลือก setupNewInstance → ▶ Run → อนุญาตสิทธิ์
//  ทำให้: ตั้ง SS_ID/DRIVE_FOLDER_ID ใน Script Properties · สร้างโฟลเดอร์ + SLIPS ·
//         สร้างชีททั้งหมดตามโครงเดิม (criteria/ค่าธรรมเนียม/โควต้าอยู่ใน COMP ในโค้ดแล้ว)
//  ไม่ดึงข้อมูลผู้ลงทะเบียนปีก่อนมา
// ================================================================
function setupNewInstance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ต้องรันจาก Apps Script ที่เปิดจาก Google Sheet (ส่วนขยาย → Apps Script)');
  // ไม่ใช้ ui.alert เลย — dialog เปิดในแท็บ Sheet แล้วรอคนกด OK ทำให้ editor ดูเหมือนค้าง
  // ผลลัพธ์ดูได้ใน "บันทึกการดำเนินการ" ของ editor + toast มุมขวาล่างของ Sheet
  var props = PropertiesService.getScriptProperties();

  var cur = props.getProperty('SS_ID');
  if (cur && cur !== ss.getId()) {
    // Properties ค้างค่าของ Sheet อื่น (เช่น ก๊อปโปรเจกต์มาจากปีก่อน) → ใช้ Sheet ที่สคริปต์ผูกอยู่เป็นหลัก สร้างโฟลเดอร์ใหม่ให้
    Logger.log('⚠️ Script Properties เคยชี้ไป Sheet อื่น (' + cur + ') → เปลี่ยนมาใช้ Sheet นี้: ' + ss.getId());
    props.deleteProperty('DRIVE_FOLDER_ID');
  }
  var reg = ss.getSheetByName(SH_REG);
  if (reg && reg.getLastRow() > 1)
    throw new Error('ยกเลิก: Sheet นี้มีข้อมูลผู้ลงทะเบียนอยู่แล้ว — setupNewInstance ใช้กับ Sheet ใหม่เท่านั้น');

  // โฟลเดอร์ Drive (เก็บ Sheet + สลิป)
  var folder = null, fid = props.getProperty('DRIVE_FOLDER_ID');
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; } }
  if (!folder) {
    folder = DriveApp.createFolder(CONF_NAME_EN + ' — ไฟล์ระบบ');
    try { DriveApp.getFileById(ss.getId()).moveTo(folder); } catch (e) { Logger.log('moveTo: ' + e.message); }
  }
  if (!folder.getFoldersByName('SLIPS').hasNext()) folder.createFolder('SLIPS');

  props.setProperties({ SS_ID: ss.getId(), DRIVE_FOLDER_ID: folder.getId() });
  try { _printSecret(); } catch (pe) { Logger.log('PRINT_SECRET: ' + pe.message); }   // Print.gs
  SS_ID = ss.getId(); DRIVE_FOLDER_ID = folder.getId();

  try { ss.setSpreadsheetTimeZone('Asia/Bangkok'); } catch (tze) { Logger.log('timezone: ' + tze.message); }
  setupSheets(true);

  // ลบชีทเปล่าที่ Google สร้างมาให้
  ['Sheet1', 'แผ่น1', 'แผ่นงาน1', 'ชีต1', 'Hoja 1', 'Feuille 1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  var todo = [];
  if (!_cfg('STAFF_PASSWORD')) todo.push('• กรอก STAFF_PASSWORD ในชีท "ตั้งค่า" (ยังว่าง = staff เข้าระบบไม่ได้)');
  if (!_cfg('PROMPTPAY_ID'))   todo.push('• กรอก PROMPTPAY_ID ในชีท "ตั้งค่า"');
  Logger.log('✅ ติดตั้งเสร็จ\nSheet: ' + ss.getUrl() + '\nโฟลเดอร์: ' + folder.getUrl() + '\n' +
    (todo.length ? 'ยังต้องทำ:\n' + todo.join('\n') + '\n' : '') +
    'จากนั้น Deploy → New deployment → Web app (Execute as: Me · Who has access: Anyone)');
  try { ss.toast(todo.length ? 'กรอกรหัส staff + PromptPay ในชีท "ตั้งค่า" แล้ว Deploy ได้เลย' : 'พร้อม Deploy', '✅ ติดตั้งเสร็จ', 10); } catch (e) {}
}

function setupSheets(silent) {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var log = [];

  var STYLES = {
    REGISTRATIONS: { bg:'#0f1923', fg:'#fbbf24' },
    COMPETITIONS:  { bg:'#0f1923', fg:'#67e8f9' },
    PAYMENTS:      { bg:'#0f1923', fg:'#86efac' },
    MEAL_SUMMARY:  { bg:'#0f1923', fg:'#fdba74' },
    MEAL_CHECKIN:  { bg:'#0f1923', fg:'#c4b5fd' },
  };

  // opts.onlyIfEmpty : ไม่แตะ header ถ้าชีทมีเนื้อหาอยู่แล้ว (F1 — ชีทรายการมี title row + header row 2)
  // opts.meal13       : MEAL_SUMMARY — ถ้ามีข้อมูลแล้วให้แก้ header เป็น 13 คอลัมน์โดยไม่ขยับข้อมูล (F2/M1)
  function makeSheet(name, headers, widths, opts) {
    opts = opts || {};
    var existing = ss.getSheetByName(name);
    var sh;
    if (existing) {
      sh = existing;
      if (opts.onlyIfEmpty && sh.getLastRow() > 0) {
        log.push('⏭ [ข้าม — มีข้อมูล/หัวตารางอยู่แล้ว] ' + name);
        return sh;
      }
      if (opts.meal13 && sh.getLastRow() >= 2) {
        var mh = sh.getRange(1,1,1,Math.max(13, sh.getLastColumn())).getValues()[0];
        var l2 = sh.getRange(2,12).getValue();
        if (String(mh[12]||'').trim() === '' && !(l2 instanceof Date)) {
          // ข้อมูลเป็นแบบ 13 คอลัมน์ (L=แพ้อาหาร, M=วันที่) → แก้ header L,M เท่านั้น
          sh.getRange(1,12,1,2).setValues([['การแพ้อาหาร','วันที่บันทึก']]);
          sh.getRange(1,13).setBackground(sh.getRange(1,12).getBackground())
            .setFontColor(sh.getRange(1,12).getFontColor()).setFontWeight('bold')
            .setHorizontalAlignment('center');
          sh.setColumnWidth(12,160); sh.setColumnWidth(13,130);
          log.push('✓ [อัปเดต header col L/M เท่านั้น] ' + name);
        } else {
          log.push('⏭ [ข้าม — layout เดิม ตรวจเอง] ' + name);
        }
        return sh;
      }
      var oldCols = sh.getLastColumn();
      if (oldCols > headers.length) {
        sh.getRange(1,1,1,oldCols).clearContent().clearFormat();
      }
      log.push('✓ [อัปเดต] ' + name);
    } else {
      sh = ss.insertSheet(name);
      log.push('✅ [สร้างใหม่] ' + name);
    }
    var range = sh.getRange(1,1,1,headers.length);
    range.setValues([headers]);
    var style = STYLES[name] || {bg:'#1e293b',fg:'#ffffff'};
    range.setBackground(style.bg).setFontColor(style.fg)
         .setFontWeight('bold').setFontSize(11)
         .setVerticalAlignment('middle').setHorizontalAlignment('center')
         .setWrap(false);
    sh.setRowHeight(1,36);
    if(sh.getFrozenRows()<1) sh.setFrozenRows(1);
    if(widths) _setWidths(sh, widths);
    headers.forEach(function(h,i){
      if(['วันที่ลงทะเบียน','วันที่ส่ง','วันอนุมัติ','วันที่บันทึก'].indexOf(h)>=0){
        if (sh.getMaxRows() < 1001) sh.insertRowsAfter(sh.getMaxRows(), 1001 - sh.getMaxRows());
        sh.getRange(2,i+1,1000,1).setNumberFormat('dd/MM/yyyy HH:mm');
      }
    });
    return sh;
  }

  // ══ 1. REGISTRATIONS (A–U, 21 cols) ═══════════════════════════
  // ข้อมูลผู้ลงทะเบียนหลักทุกคน (ทั่วไป + กรรมการ)
  makeSheet('REGISTRATIONS',[
    'REG_ID',               // A  — INHI26-MMDD-XXXX
    'วันที่ลงทะเบียน',     // B  — timestamp
    'คำนำหน้า',             // C
    'ชื่อ',                 // D
    'นามสกุล',              // E
    'ชื่อเต็ม',             // F
    'อีเมล',                // G
    'โทรศัพท์',             // H
    'เลขบัตรประชาชน',      // I  — 13 หลัก (บังคับ)
    'ประเภทผู้ลงทะเบียน',  // J
    'ฟรี/ชำระ',             // K
    'สังกัด/หน่วยงาน',     // L
    'ตัดสินประเภท',         // M  — เฉพาะกรรมการตัดสิน
    'การแพ้อาหาร',          // N
    'ที่อยู่ใบเสร็จ',      // O
    'เลขใบประกอบวิชาชีพ',  // P  — 10 หลัก (ไม่บังคับ)
    'ยอดรวม',               // Q  — บาท
    'สถานะชำระ',            // R
    'อาหาร_27พค',           // S  — จำนวนกล่อง
    'อาหาร_28พค',           // T  — จำนวนกล่อง
    'คนได้ข้าว_รวม',       // U
  ],[130,130,70,90,100,160,180,100,130,150,70,180,150,120,220,140,80,110,80,80,90]);

  // ══ 2. COMPETITIONS (A–H, 8 cols) ════════════════════════════
  // บันทึกรายการที่ลงทะเบียน (summary — รายละเอียดอยู่ในชีทรายการ)
  makeSheet('COMPETITIONS',[
    'REG_ID',               // A
    'ลำดับ',                // B
    'รหัสรายการ',           // C
    'ชื่อรายการ',           // D
    'จำนวนวัน',             // E
    'โควต้าทีม',            // F
    'รายชื่อสมาชิก',        // G
    'วันที่บันทึก',         // H
  ],[130,60,90,260,70,80,300,130]);

  // ══ 3. PAYMENTS (A–J, 10 cols) ═══════════════════════════════
  makeSheet('PAYMENTS',[
    'REG_ID',               // A
    'วันที่ส่ง',            // B
    'จำนวนเงิน',            // C  — บาท
    'วิธีชำระ',             // D
    'วันที่โอน',            // E  — วันที่โอน (_submitSlip เขียน transferDate||ref)
    'URL_สลิป',             // F  — Google Drive link
    'สถานะ',                // G  — รอตรวจสอบ|ชำระแล้ว|ปฏิเสธ
    'ผู้อนุมัติ',            // H  — staffName
    'วันอนุมัติ',            // I
    'ที่อยู่ใบเสร็จ',      // J
  ],[130,130,90,90,140,280,100,140,130,260]);

  // ══ 4. MEAL_SUMMARY (A–L, 12 cols) ══════════════════════════
  // สรุปสิทธิ์อาหารต่อรายการ (ไม่รวม Online c00)
  makeSheet('MEAL_SUMMARY',[
    'REG_ID',               // A
    'ชื่อ-นามสกุล',        // B
    'เลขบัตรประชาชน',      // C
    'ประเภท',               // D
    'สังกัด',               // E
    'รายการ/ที่มา',         // F
    'อาหาร_27พค',           // G  — จำนวนกล่อง
    'อาหาร_28พค',           // H  — จำนวนกล่อง
    'สถานะชำระเงิน',       // I  — VLOOKUP PAYMENTS
    'รับ_27พค',             // J  — ✓ รับแล้ว / ว่าง
    'รับ_28พค',             // K  — ✓ รับแล้ว / ว่าง
    'การแพ้อาหาร',          // L
    'วันที่บันทึก',         // M
  ],[130,160,130,120,180,200,80,80,110,90,90,160,130],{meal13:true});

  // ══ 5. MEAL_CHECKIN (A–H, 8 cols) ═══════════════════════════
  // บันทึกการรับอาหารจริง
  makeSheet('MEAL_CHECKIN',[
    'REG_ID',               // A
    'ชื่อ-นามสกุล',        // B
    'สังกัด',               // C
    'วันที่รับอาหาร',       // D  — "27 พฤษภาคม 2570" / "28 พฤษภาคม 2570"
    'จำนวนคน',              // E  — จำนวนกล่อง
    'ชื่อผู้รับอาหาร (ตัวแทนทีม)', // F — ผู้มารับ (ผู้แจกอยู่ในคอลัมน์ G)
    'บันทึกโดย (Staff)',   // G  — staffName ที่ login
    'เวลาบันทึก',           // H
  ],[130,160,180,160,80,160,150,140]);

  // ══ 6. Competition sheets (A–AB, 28 cols) ════════════════════
  // แต่ละรายการแข่งขัน มีสมาชิก 1-4 คน พร้อมเลขบัตร ใบประกอบ แพ้อาหาร
  // F1: ใช้ layout จริง 36 คอลัมน์ (COMP_SHEET_HDR) เหมือน _writeCompSheets
  var COMP_HDR_28 = COMP_SHEET_HDR;
  var COMP_W_28 = [60,120,130,160,160,160,160,100,70,110,
    70,100,100,110,90,110, 70,100,100,110,90,110,
    70,100,100,110,90,110, 70,100,100,110,90,110,
    100,130];

  // Online — 8 cols เท่านั้น (ไม่มีทีม)
  var COMP_HDR_ONLINE = [
    'REG_ID','วันที่ลงทะเบียน','ชื่อ-นามสกุล','สังกัด',
    'อีเมล','โทรศัพท์','สถานะชำระ','วันที่บันทึก'
  ];
  var COMP_W_ONLINE = [130,130,180,200,180,100,110,130];

  var compDefs = [
    {nm:'ประชุม_Online',  hdr:COMP_HDR_ONLINE, w:COMP_W_ONLINE},
    {nm:'ประชุม_Onsite',  hdr:COMP_HDR_28,     w:COMP_W_28},
    {nm:'SIM_MAN',        hdr:COMP_HDR_28,     w:COMP_W_28},
    {nm:'SIM_MOM',        hdr:COMP_HDR_28,     w:COMP_W_28},
    {nm:'SIM_PED',        hdr:COMP_HDR_28,     w:COMP_W_28},
    {nm:'แนวปฏิบัติ_สถาบัน',    hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'แนวปฏิบัติ_สถานบริการ',hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'บทความวิจัย_บุคลากร',  hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'บทความวิจัย_นศ',       hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'ตอบคำถาม_พยาบาล',      hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'ตอบคำถาม_กายวิภาค',    hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'นวัตกรรม_นศ',          hdr:COMP_HDR_28, w:COMP_W_28},
    {nm:'นวัตกรรม_อาจารย์',     hdr:COMP_HDR_28, w:COMP_W_28},
  ];
  // ชีทที่ยังไม่มี → สร้างด้วย _compSheet (title แถว 1 + header แถว 2) ให้ตรงกับที่ runtime/รายงานคาดไว้
  var compKeyByName = {};
  Object.keys(COMP_SHEET).forEach(function(k){ compKeyByName[COMP_SHEET[k]] = k; });
  compDefs.forEach(function(c){
    if (ss.getSheetByName(c.nm)) { makeSheet(c.nm, c.hdr, c.w, {onlyIfEmpty:true}); return; }
    var cfg = COMP[compKeyByName[c.nm]] || {nm:c.nm, nm_en:c.nm};
    _compSheet(ss, c.nm, cfg.nm, cfg.nm_en, c.hdr === COMP_HDR_ONLINE ? COMP_ONLINE_HDR : null);
    log.push('✅ [สร้างใหม่ title+header แถว 2] ' + c.nm);
  });

  // ชีท "ตั้งค่า" (รหัส staff / PromptPay / วันหมดเขต)
  _ensureCfgSheet(ss, log);

  // ชีทจากไฟล์อื่น (เช่น Submissions.gs)
  if (typeof _extraSetupSheets === 'function') { try { _extraSetupSheets(ss, makeSheet, log); } catch (e) { log.push('⚠️ extra: ' + e.message); } }

  // จัดลำดับ sheets
  var order = [
    SH_CFG,'REGISTRATIONS','COMPETITIONS','PAYMENTS','MEAL_SUMMARY','MEAL_CHECKIN',
    'ประชุม_Online','ประชุม_Onsite',
    'SIM_MAN','SIM_MOM','SIM_PED',
    'แนวปฏิบัติ_สถาบัน','แนวปฏิบัติ_สถานบริการ',
    'บทความวิจัย_บุคลากร','บทความวิจัย_นศ',
    'ตอบคำถาม_พยาบาล','ตอบคำถาม_กายวิภาค',
    'นวัตกรรม_นศ','นวัตกรรม_อาจารย์'
  ];
  order.forEach(function(name,pos){
    var sh=ss.getSheetByName(name);
    if(sh && sh.getIndex()!==pos+1) ss.setActiveSheet(sh), ss.moveActiveSheet(pos+1);
  });

  // ไม่เปิด dialog เลย — dialog เปิดในแท็บ Sheet แล้วรอคนกด OK ทำให้ editor ดูเหมือนค้างจนหมดเวลา
  // ผลดูที่ "บันทึกการดำเนินการ" ใน editor (+ toast มุมขวาล่างของ Sheet)
  Logger.log('✅ setupSheets เสร็จสมบูรณ์ (ข้อมูลไม่ถูกลบ)\n' + log.join('\n'));
  try { ss.toast('สร้าง/อัปเดต ' + log.length + ' ชีท — ดูรายละเอียดในบันทึกของ Apps Script', '✅ setupSheets เสร็จ', 8); } catch (e) {}
  return log;
}

function updateSheets() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  var STYLES = {
    REGISTRATIONS: { bg:'#0f1923', fg:'#fbbf24' },
    COMPETITIONS:  { bg:'#0f1923', fg:'#67e8f9' },
    PAYMENTS:      { bg:'#0f1923', fg:'#86efac' },
    MEAL_SUMMARY:  { bg:'#0f1923', fg:'#fdba74' },
    MEAL_CHECKIN:  { bg:'#0f1923', fg:'#c4b5fd' },
    'ประชุม_Online':{ bg:'#1e3a5f', fg:'#93c5fd' },
  };

  function updateHeader(name, headers, widths) {
    var sh = ss.getSheetByName(name);
    if (!sh) { log.push('⚠️ ไม่พบ: ' + name); return; }
    var oldCols = sh.getLastColumn();
    if (oldCols > 0) sh.getRange(1,1,1,Math.max(oldCols,headers.length)).clearContent().clearFormat();
    var range = sh.getRange(1,1,1,headers.length);
    range.setValues([headers]);
    var style = STYLES[name] || { bg:'#1a1a2e', fg:'#a5f3fc' };
    range.setBackground(style.bg).setFontColor(style.fg)
         .setFontWeight('bold').setFontSize(11)
         .setVerticalAlignment('middle').setHorizontalAlignment('center')
         .setWrap(false);
    sh.setRowHeight(1, 38);
    if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
    if (widths) widths.forEach(function(w,i){ if(w) sh.setColumnWidth(i+1,w); });
    headers.forEach(function(h,i){
      if(['วันที่ลงทะเบียน','วันที่ส่ง','วันอนุมัติ','วันที่บันทึก'].indexOf(h)>=0)
        sh.getRange(2,i+1,Math.max(sh.getLastRow()-1,1),1).setNumberFormat('dd/MM/yyyy HH:mm');
    });
    log.push('✅ ' + name + ' (' + headers.length + ' cols)');
  }

  // ══ 1. REGISTRATIONS (A–U, 21 cols) ═════════════════════════
  updateHeader('REGISTRATIONS',[
    'REG_ID','วันที่ลงทะเบียน','คำนำหน้า','ชื่อ','นามสกุล','ชื่อเต็ม',
    'อีเมล','โทรศัพท์','เลขบัตรประชาชน','ประเภทผู้ลงทะเบียน','ฟรี/ชำระ',
    'สังกัด/หน่วยงาน','ตัดสินประเภท','การแพ้อาหาร','ที่อยู่ใบเสร็จ',
    'เลขใบประกอบวิชาชีพ','ยอดรวม','สถานะชำระ','อาหาร_27พค','อาหาร_28พค','คนได้ข้าว_รวม'
  ],[130,130,70,90,100,160,180,100,130,150,70,180,150,120,220,140,80,110,80,80,90]);

  // ══ 2. COMPETITIONS (A–H, 8 cols) ════════════════════════════
  updateHeader('COMPETITIONS',[
    'REG_ID','ลำดับ','รหัสรายการ','ชื่อรายการ','จำนวนวัน','โควต้าทีม','รายชื่อสมาชิก','วันที่บันทึก'
  ],[130,60,90,260,70,80,300,130]);

  // ══ 3. PAYMENTS (A–J, 10 cols) ═══════════════════════════════
  // col E = วันที่โอน (เปลี่ยนจากเลขอ้างอิง)
  updateHeader('PAYMENTS',[
    'REG_ID','วันที่ส่ง','จำนวนเงิน','วิธีชำระ','วันที่โอน',
    'URL_สลิป','สถานะ','ผู้อนุมัติ','วันอนุมัติ','ที่อยู่ใบเสร็จ'
  ],[130,130,90,90,140,280,100,140,130,260]);

  // ══ 4. MEAL_SUMMARY (A–L, 12 cols) ══════════════════════════
  updateHeader('MEAL_SUMMARY',[
    'REG_ID','ชื่อ-นามสกุล','เลขบัตรประชาชน','ประเภท','สังกัด','รายการ/ที่มา',
    'อาหาร_27พค','อาหาร_28พค','สถานะชำระเงิน','รับ_27พค','รับ_28พค','การแพ้อาหาร','วันที่บันทึก'
  ],[130,160,130,120,180,200,80,80,110,90,90,160,130]);

  // ══ 5. MEAL_CHECKIN (A–H, 8 cols) ════════════════════════════
  updateHeader('MEAL_CHECKIN',[
    'REG_ID','ชื่อ-นามสกุล','สังกัด','วันที่รับอาหาร','จำนวนคน',
    'ชื่อผู้รับอาหาร (ตัวแทนทีม)','บันทึกโดย (Staff)','เวลาบันทึก'
  ],[130,160,180,160,80,160,150,140]);

  // ══ 6. ประชุม_Online (A–H, 8 cols) ══════════════════════════
  updateHeader('ประชุม_Online',[
    'REG_ID','วันที่ลงทะเบียน','ชื่อ-นามสกุล','สังกัด',
    'อีเมล','โทรศัพท์','สถานะชำระ','วันที่บันทึก'
  ],[130,130,180,200,180,100,110,130]);

  // ══ 7-18. Competition sheets (A–AJ, 34 cols) ═════════════════
  // สมาชิก 1-4 × (คำนำหน้า + ชื่อ + นามสกุล + เลขบัตร + ใบประกอบ + แพ้อาหาร)
  var HDR34 = [
    'ลำดับทีม','REG_ID','วันที่ลงทะเบียน','ชื่อทีม',
    'ชื่อหัวหน้าทีม','สังกัด','อีเมล','โทรศัพท์','จำนวนวัน','ค่าธรรมเนียม (บาท)',
    'คำนำหน้า 1','ชื่อ 1','นามสกุล 1','เลขบัตร 1','ใบประกอบ 1','แพ้อาหาร 1',
    'คำนำหน้า 2','ชื่อ 2','นามสกุล 2','เลขบัตร 2','ใบประกอบ 2','แพ้อาหาร 2',
    'คำนำหน้า 3','ชื่อ 3','นามสกุล 3','เลขบัตร 3','ใบประกอบ 3','แพ้อาหาร 3',
    'คำนำหน้า 4','ชื่อ 4','นามสกุล 4','เลขบัตร 4','ใบประกอบ 4','แพ้อาหาร 4',
    'สถานะชำระ','วันที่บันทึก'
  ];
  var W34 = [60,120,130,160,160,160,160,100,70,110,
    70,100,100,110,90,110, 70,100,100,110,90,110,
    70,100,100,110,90,110, 70,100,100,110,90,110,
    100,130];
  ['ประชุม_Onsite','SIM_MAN','SIM_MOM','SIM_PED',
   'แนวปฏิบัติ_สถาบัน','แนวปฏิบัติ_สถานบริการ',
   'บทความวิจัย_บุคลากร','บทความวิจัย_นศ',
   'ตอบคำถาม_พยาบาล','ตอบคำถาม_กายวิภาค',
   'นวัตกรรม_นศ','นวัตกรรม_อาจารย์'
  ].forEach(function(n){
    // ชีทแข่งขันจริงมี title (merge) แถว 1 + header แถว 2 — ถ้ามีข้อมูลแล้วห้ามเขียนทับแถว 1
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) { log.push('⏭️ ข้าม (มีข้อมูลแล้ว): ' + n); return; }
    updateHeader(n, HDR34, W34);
  });

  SpreadsheetApp.flush();
  ui.alert('✅ updateSheets เสร็จสมบูรณ์',
    log.join('\n')+'\n\n✅ แก้แค่ header row 1 — ข้อมูลทุกแถวยังครบ',
    ui.ButtonSet.OK);
}


function fixDoubleHeader() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheetNames = [
    'ประชุม_Onsite','SIM_MAN','SIM_MOM','SIM_PED',
    'แนวปฏิบัติ_สถาบัน','แนวปฏิบัติ_สถานบริการ',
    'บทความวิจัย_บุคลากร','บทความวิจัย_นศ',
    'ตอบคำถาม_พยาบาล','ตอบคำถาม_กายวิภาค',
    'นวัตกรรม_นศ','นวัตกรรม_อาจารย์'
  ];
  var fixed = [];
  sheetNames.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    // ตรวจว่า row 2 เป็น header เก่าไหม (มีคำว่า 'สมาชิก')
    if (sh.getLastRow() >= 2) {
      var row2 = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
      // F16: ต้องมี 'REG_ID' อยู่ใน row 2 จึงถือเป็น header ซ้ำ (กันลบข้อมูลจริงที่มีคำว่า 'สมาชิก')
      var isOldHeader = row2.some(function(v) {
        return String(v).trim() === 'REG_ID';
      });
      if (isOldHeader) {
        sh.deleteRow(2);
        fixed.push(name);
      }
    }
  });
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '✅ แก้ไขเสร็จ',
    fixed.length > 0
      ? 'ลบ header row 2 ออกจาก:\n' + fixed.join('\n') + '\n\nรัน setupSheets() ต่อได้เลย'
      : 'ไม่พบ header ซ้ำ — รัน setupSheets() ได้เลย',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════
//  resetAndSetupSheets()
//  ▶ ลบทุก sheet แล้วสร้างใหม่ทั้งหมด พร้อม header + format
//  ⚠️  ข้อมูลทั้งหมดจะหายไป — ใช้เฉพาะตอน reset เริ่มต้น
// ════════════════════════════════════════════════════════════════
function resetAndSetupSheets() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  // F4: นับแถวที่จะหาย + ต้องพิมพ์ RESET + สำรองไฟล์ก่อนลบ
  var cntLines = [], cntTotal = 0;
  ss.getSheets().forEach(function(sh0) {
    var n0 = Math.max(sh0.getLastRow() - 1, 0);
    if (n0 > 0) { cntLines.push(sh0.getName() + ': ' + n0); cntTotal += n0; }
  });
  var pr = ui.prompt(
    '⚠️ ยืนยันการ Reset — ข้อมูลทั้งหมดจะถูกลบ',
    'แถวข้อมูลที่จะหาย รวม ' + cntTotal + ' แถว\n' + cntLines.join('\n') +
    '\n\nระบบจะสำรองไฟล์ให้ก่อนลบ\nพิมพ์คำว่า RESET (ตัวพิมพ์ใหญ่) เพื่อยืนยัน',
    ui.ButtonSet.OK_CANCEL
  );
  if (pr.getSelectedButton() !== ui.Button.OK || String(pr.getResponseText()).trim() !== 'RESET') {
    ui.alert('ยกเลิกการ Reset');
    return;
  }
  var bk = DriveApp.getFileById(SS_ID).makeCopy(
    'BACKUP before reset ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmm'));
  Logger.log('Backup created: ' + bk.getUrl());

  // ── สีหัวตาราง ───────────────────────────────────────────────
  var STYLE = {
    REG:    { bg:'#0f1923', fg:'#fbbf24' },
    COMP:   { bg:'#0f1923', fg:'#67e8f9' },
    PAY:    { bg:'#0f1923', fg:'#86efac' },
    MEAL_S: { bg:'#0f1923', fg:'#fdba74' },
    MEAL_C: { bg:'#0f1923', fg:'#c4b5fd' },
    ONLINE: { bg:'#1e3a5f', fg:'#93c5fd' },
    COMPSH: { bg:'#1a1a2e', fg:'#a5f3fc' },
  };

  // ── Helper: สร้าง sheet ใหม่ fresh ───────────────────────────
  function createSheet(name, headers, widths, style) {
    // ลบถ้ามีอยู่แล้ว
    var old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
    Utilities.sleep(100);

    var sh = ss.insertSheet(name);
    var r  = sh.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setBackground(style.bg)
     .setFontColor(style.fg)
     .setFontWeight('bold')
     .setFontSize(11)
     .setVerticalAlignment('middle')
     .setHorizontalAlignment('center')
     .setWrap(false);
    sh.setRowHeight(1, 38);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, widths[0] || 130);
    widths.forEach(function(w, i) { if (w) sh.setColumnWidth(i+1, w); });

    // format วันที่
    headers.forEach(function(h, i) {
      if (['วันที่ลงทะเบียน','วันที่ส่ง','วันอนุมัติ','วันที่บันทึก'].indexOf(h) >= 0)
        sh.getRange(2, i+1, 2000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    });

    // format ตัวเลข
    headers.forEach(function(h, i) {
      if (['ยอดรวม','จำนวนเงิน','ค่าธรรมเนียม (บาท)'].indexOf(h) >= 0)
        sh.getRange(2, i+1, 2000, 1).setNumberFormat('#,##0');
    });

    return sh;
  }

  // ════════════════════════════════════════════════════════════
  //  1. REGISTRATIONS — A…U (21 cols)
  // ════════════════════════════════════════════════════════════
  createSheet('REGISTRATIONS', [
    'REG_ID',               // A  INHI26-MMDD-XXXX
    'วันที่ลงทะเบียน',     // B
    'คำนำหน้า',             // C
    'ชื่อ',                 // D
    'นามสกุล',              // E
    'ชื่อเต็ม',             // F
    'อีเมล',                // G
    'โทรศัพท์',             // H
    'เลขบัตรประชาชน',      // I  13 หลัก (บังคับ)
    'ประเภทผู้ลงทะเบียน',  // J
    'ฟรี/ชำระ',             // K
    'สังกัด/หน่วยงาน',     // L
    'ตัดสินประเภท',         // M  เฉพาะกรรมการตัดสิน
    'การแพ้อาหาร',          // N
    'ที่อยู่ใบเสร็จ',      // O
    'เลขใบประกอบวิชาชีพ',  // P  10 หลัก (ไม่บังคับ)
    'ยอดรวม',               // Q  บาท
    'สถานะชำระ',            // R
    'อาหาร_27พค',           // S  จำนวนกล่อง
    'อาหาร_28พค',           // T  จำนวนกล่อง
    'คนได้ข้าว_รวม',       // U
  ], [130,130,70,90,100,160,180,100,130,150,70,180,150,120,220,140,80,110,80,80,90],
  STYLE.REG);

  // ════════════════════════════════════════════════════════════
  //  2. COMPETITIONS — A…H (8 cols)  summary per registration
  // ════════════════════════════════════════════════════════════
  createSheet('COMPETITIONS', [
    'REG_ID',               // A
    'ลำดับ',                // B
    'รหัสรายการ',           // C
    'ชื่อรายการ',           // D
    'จำนวนวัน',             // E
    'โควต้าทีม',            // F
    'รายชื่อสมาชิก',        // G
    'วันที่บันทึก',         // H
  ], [130,60,90,260,70,80,300,130], STYLE.COMP);

  // ════════════════════════════════════════════════════════════
  //  3. PAYMENTS — A…J (10 cols)
  // ════════════════════════════════════════════════════════════
  createSheet('PAYMENTS', [
    'REG_ID',               // A
    'วันที่ส่ง',            // B
    'จำนวนเงิน',            // C  บาท
    'วิธีชำระ',             // D
    'วันที่โอน',            // E  วันที่โอน (_submitSlip เขียน transferDate||ref)
    'URL_สลิป',             // F  Google Drive link
    'สถานะ',                // G  รอตรวจสอบ|ชำระแล้ว|ปฏิเสธ
    'ผู้อนุมัติ',            // H  staffName
    'วันอนุมัติ',            // I
    'ที่อยู่ใบเสร็จ',      // J
  ], [130,130,90,90,140,280,100,140,130,260], STYLE.PAY);

  // ════════════════════════════════════════════════════════════
  //  4. MEAL_SUMMARY — A…M (13 cols)  ไม่รวม Online
  // ════════════════════════════════════════════════════════════
  createSheet('MEAL_SUMMARY', [
    'REG_ID',               // A
    'ชื่อ-นามสกุล',        // B
    'เลขบัตรประชาชน',      // C
    'ประเภท',               // D
    'สังกัด',               // E
    'รายการ/ที่มา',         // F
    'อาหาร_27พค',           // G  จำนวนกล่อง
    'อาหาร_28พค',           // H  จำนวนกล่อง
    'สถานะชำระเงิน',       // I  VLOOKUP PAYMENTS
    'รับ_27พค',             // J  ✓รับแล้ว / ว่าง
    'รับ_28พค',             // K  ✓รับแล้ว / ว่าง
    'การแพ้อาหาร',          // L
    'วันที่บันทึก',         // M
  ], [130,160,130,120,180,200,80,80,110,90,90,160,130], STYLE.MEAL_S);

  // ════════════════════════════════════════════════════════════
  //  5. MEAL_CHECKIN — A…H (8 cols)
  // ════════════════════════════════════════════════════════════
  createSheet('MEAL_CHECKIN', [
    'REG_ID',               // A
    'ชื่อ-นามสกุล',        // B
    'สังกัด',               // C
    'วันที่รับอาหาร',       // D  "27 พฤษภาคม 2570"
    'จำนวนคน',              // E  จำนวนกล่อง
    'ชื่อผู้รับอาหาร (ตัวแทนทีม)', // F
    'บันทึกโดย (Staff)',   // G  staffName ที่ login
    'เวลาบันทึก',           // H
  ], [130,160,180,160,80,160,150,140], STYLE.MEAL_C);

  // ════════════════════════════════════════════════════════════
  //  6. ประชุม_Online — A…H (8 cols)  ไม่มีทีม
  // ════════════════════════════════════════════════════════════
  createSheet('ประชุม_Online', [
    'REG_ID',               // A
    'วันที่ลงทะเบียน',     // B
    'ชื่อ-นามสกุล',        // C
    'สังกัด',               // D
    'อีเมล',                // E
    'โทรศัพท์',             // F
    'สถานะชำระ',            // G
    'วันที่บันทึก',         // H
  ], [130,130,180,200,180,100,110,130], STYLE.ONLINE);

  // ════════════════════════════════════════════════════════════
  //  7-18. Competition sheets — A…AB (28 cols) สมาชิก 1-4
  // ════════════════════════════════════════════════════════════
  var HDR28 = [
    'ลำดับทีม',             // A
    'REG_ID',               // B
    'วันที่ลงทะเบียน',     // C
    'ชื่อทีม',              // D
    'ชื่อหัวหน้าทีม',      // E
    'สังกัด',               // F
    'อีเมล',                // G
    'โทรศัพท์',             // H
    'จำนวนวัน',             // I
    'ค่าธรรมเนียม (บาท)',   // J
    'คำนำหน้า 1',          // K
    'ชื่อ 1',               // L
    'นามสกุล 1',            // M
    'เลขบัตร 1',            // N  13 หลัก
    'ใบประกอบ 1',           // O  10 หลัก (ไม่บังคับ)
    'แพ้อาหาร 1',           // P
    'คำนำหน้า 2',          // Q
    'ชื่อ 2',               // R
    'นามสกุล 2',            // S
    'เลขบัตร 2',            // T
    'ใบประกอบ 2',           // U
    'แพ้อาหาร 2',           // V
    'คำนำหน้า 3',          // W
    'ชื่อ 3',               // X
    'นามสกุล 3',            // Y
    'เลขบัตร 3',            // Z
    'ใบประกอบ 3',           // AA
    'แพ้อาหาร 3',           // AB
    'คำนำหน้า 4',          // AC
    'ชื่อ 4',               // AD
    'นามสกุล 4',            // AE
    'เลขบัตร 4',            // AF
    'ใบประกอบ 4',           // AG
    'แพ้อาหาร 4',           // AH
    'สถานะชำระ',            // AI
    'วันที่บันทึก',         // AJ
  ];
  var W34 = [60,120,130,160,160,160,160,100,70,110,
    70,110,100,110,90,110,70,110,100,110,90,110,70,110,100,110,90,110,70,110,100,110,90,110,100,130];
  var W28 = W34; // alias

  var compSheets = [
    'ประชุม_Onsite','SIM_MAN','SIM_MOM','SIM_PED',
    'แนวปฏิบัติ_สถาบัน','แนวปฏิบัติ_สถานบริการ',
    'บทความวิจัย_บุคลากร','บทความวิจัย_นศ',
    'ตอบคำถาม_พยาบาล','ตอบคำถาม_กายวิภาค',
    'นวัตกรรม_นศ','นวัตกรรม_อาจารย์'
  ];
  compSheets.forEach(function(name) {
    createSheet(name, HDR28, W28, STYLE.COMPSH);
  });

  // ── จัดลำดับ sheet ───────────────────────────────────────────
  var ORDER = [
    'REGISTRATIONS','COMPETITIONS','PAYMENTS','MEAL_SUMMARY','MEAL_CHECKIN',
    'ประชุม_Online','ประชุม_Onsite',
    'SIM_MAN','SIM_MOM','SIM_PED',
    'แนวปฏิบัติ_สถาบัน','แนวปฏิบัติ_สถานบริการ',
    'บทความวิจัย_บุคลากร','บทความวิจัย_นศ',
    'ตอบคำถาม_พยาบาล','ตอบคำถาม_กายวิภาค',
    'นวัตกรรม_นศ','นวัตกรรม_อาจารย์'
  ];
  ORDER.forEach(function(name, pos) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(pos + 1); }
  });

  // ── ลบ Sheet1 default ถ้ามี ──────────────────────────────────
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1') || ss.getSheetByName('ชีต1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.flush();
  ui.alert('✅ Reset เสร็จสมบูรณ์',
    '18 sheets พร้อมใช้งาน\nไม่มีข้อมูลเดิม\nพร้อมรับการลงทะเบียน',
    ui.ButtonSet.OK);
}

// ── อัปเดต MEAL_SUMMARY เมื่อมีการ check-in ─────────────────
function _updateMealSummaryCheckin(rid, day){
  try{
    var ss=SpreadsheetApp.openById(SS_ID);
    var sh=ss.getSheetByName(SH_MEAL);
    if(!sh||sh.getLastRow()<2) return;
    var data=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    var col=day==='17'?10:11; // J=10, K=11
    for(var i=0;i<data.length;i++){
      if(String(data[i][0]).trim()===String(rid).trim()){
        // F5: ไม่เขียนทับสูตร COUNTIFS (สูตรสะท้อน MEAL_CHECKIN อยู่แล้ว)
        var cc=sh.getRange(i+2,col);
        if(!cc.getFormula()) cc.setValue('✓ รับแล้ว');
      }
    }
  }catch(e){}
}

// ── แก้สูตร/ค่าใน MEAL_SUMMARY ทุกแถว (รัน 1 ครั้งเพื่อแก้ข้อมูลเก่า) ──
function fixMealSummaryFormulas(){
  var ss=SpreadsheetApp.openById(SS_ID);
  var sh=ss.getSheetByName(SH_MEAL);
  var csh=ss.getSheetByName(SH_CHECKIN);
  if(!sh||sh.getLastRow()<2){
    SpreadsheetApp.getUi().alert('ไม่มีข้อมูลใน MEAL_SUMMARY');return;
  }
  var lastRow=sh.getLastRow();
  var mealData=sh.getRange(2,1,lastRow-1,1).getValues();
  var ciRows=[];
  if(csh&&csh.getLastRow()>1){
    ciRows=csh.getRange(2,1,csh.getLastRow()-1,5).getValues();
  }
  var got17={},got18={};
  ciRows.forEach(function(r){
    var rid=String(r[0]).trim();
    var day=String(r[3]||'');
    if(day.indexOf('27')>=0) got17[rid]=true;
    if(day.indexOf('28')>=0) got18[rid]=true;
  });
  // F5: batch + คืนสูตร COUNTIFS ใน J/K (ไม่เขียนข้อความตายตัวทับ)
  var fI=[],fJ=[],fK=[];
  for(var i=0;i<mealData.length;i++){
    var rn=i+2;
    fI.push(['=IFERROR(VLOOKUP(A'+rn+',PAYMENTS!$A:$G,7,0),"รอชำระเงิน")']);
    fJ.push(['=IFERROR(IF(COUNTIFS(MEAL_CHECKIN!$A:$A,A'+rn+',MEAL_CHECKIN!$D:$D,"*27*")>0,"✓ รับแล้ว"),"")']);
    fK.push(['=IFERROR(IF(COUNTIFS(MEAL_CHECKIN!$A:$A,A'+rn+',MEAL_CHECKIN!$D:$D,"*28*")>0,"✓ รับแล้ว"),"")']);
  }
  sh.getRange(2,9,fI.length,1).setFormulas(fI);
  sh.getRange(2,10,fJ.length,1).setFormulas(fJ);
  sh.getRange(2,11,fK.length,1).setFormulas(fK);
  SpreadsheetApp.flush();
  var msg='✅ อัปเดตแล้ว ('+(lastRow-1)+' แถว)\n';
  msg+='MEAL_CHECKIN: '+ciRows.length+' แถว\n';
  msg+='รับ 27 พ.ค.: '+Object.keys(got17).length+' ทีม\n';
  msg+='รับ 28 พ.ค.: '+Object.keys(got18).length+' ทีม';
  SpreadsheetApp.getUi().alert(msg);
}

// ── Keep GAS warm (ตั้ง trigger ทุก 5 นาที) ─────────────────
function keepWarm(){
  SpreadsheetApp.openById(SS_ID).getName();
}

// ── DEBUG: ตรวจสอบ/แก้ไขสถานะชำระใน REGISTRATIONS ────────────
// รันใน GAS Editor แล้ว check Logger เพื่อ debug
function debugPayStatus(){
  var ss=SpreadsheetApp.openById(SS_ID);
  var rsh=ss.getSheetByName(SH_REG);
  var psh=ss.getSheetByName(SH_PAY);
  var ui=SpreadsheetApp.getUi();

  if(!rsh||!psh){ui.alert('ไม่พบ sheet');return;}

  var rhdr=rsh.getDataRange().getValues()[0];
  var phdr=psh.getDataRange().getValues()[0];

  // หา col สถานะชำระ
  var iStat=rhdr.indexOf('สถานะชำระ');
  var msg='REG headers: '+rhdr.join(', ')+'\n\n';
  msg+='สถานะชำระ col index: '+iStat+' (col '+String.fromCharCode(65+iStat)+')\n';
  msg+='PAY headers: '+phdr.join(', ')+'\n\n';

  // ดึง REG IDs ที่ชำระแล้วใน PAYMENTS
  var pdata=psh.getDataRange().getValues();
  var paidMap={};
  pdata.slice(1).forEach(function(r){
    if(String(r[6]||'')==='ชำระแล้ว') paidMap[String(r[0])]=true;
  });
  msg+='REG IDs ที่ชำระแล้วใน PAYMENTS: '+Object.keys(paidMap).join(', ')+'\n\n';

  // ตรวจ REGISTRATIONS
  var rdata=rsh.getDataRange().getValues();
  var mismatch=[];
  rdata.slice(1).forEach(function(r){
    var rid=String(r[0]);
    if(paidMap[rid]&&String(r[iStat]||'')==='รอชำระเงิน'){
      mismatch.push(rid);
    }
  });

  if(mismatch.length>0){
    msg+='พบ '+mismatch.length+' รายการที่ PAYMENTS=ชำระแล้ว แต่ REGISTRATIONS ยังรอ:\n'+mismatch.join('\n');
    var fix=ui.alert('พบข้อมูลไม่ตรง',msg+'\n\nกด OK เพื่อแก้ไขอัตโนมัติ',ui.ButtonSet.OK_CANCEL);
    if(fix===ui.Button.OK){
      mismatch.forEach(function(rid){
        _updatePayStatus(ss,rid,'ชำระแล้ว');
      });
      SpreadsheetApp.flush();
      ui.alert('✅ แก้ไขแล้ว '+mismatch.length+' รายการ');
    }
  }else{
    ui.alert('ตรวจสอบเสร็จ',msg+'✅ ไม่พบความไม่ตรงกัน',ui.ButtonSet.OK);
  }
}

// (testUpdatePayStatus ฉบับแรกถูกลบ — ซ้ำชื่อกับฉบับท้ายไฟล์ซึ่งเป็นฉบับที่ใช้งานจริง)

// ── แก้ข้อมูลเก่าใน ประชุม_Online ที่บันทึก 36 cols มาเป็น 8 cols ──
function fixOnlineSheet(){
  var ss=SpreadsheetApp.openById(SS_ID);
  var sh=ss.getSheetByName('ประชุม_Online');
  var ui=SpreadsheetApp.getUi();
  if(!sh||sh.getLastRow()<2){ui.alert('ไม่มีข้อมูล');return;}

  var data=sh.getDataRange().getValues();
  var hdr=data[0];
  ui.alert('Header cols: '+hdr.length+'\nData row 2 cols: '+(data.length>1?data[1].length:0)+
           '\nRow 2 preview: '+JSON.stringify(data[1]||[]).slice(0,200));
}

// ── แก้จริง: เขียน data ใหม่ให้ตรง 8 cols ────────────────────
function fixOnlineSheetData(){
  var ss=SpreadsheetApp.openById(SS_ID);
  var sh=ss.getSheetByName('ประชุม_Online');
  var ui=SpreadsheetApp.getUi();
  if(!sh||sh.getLastRow()<2){ui.alert('ไม่มีข้อมูล');return;}

  if(sh.getLastColumn()<20){ui.alert('ชีทเป็น 8 คอลัมน์อยู่แล้ว ไม่ต้องแก้');return;}
  var conf=ui.alert('ยืนยัน','จะเขียน data ใหม่ให้ตรง 8 cols\n(backup ไว้ก่อนถ้าต้องการ)',ui.ButtonSet.OK_CANCEL);
  if(conf!==ui.Button.OK) return;

  var data=sh.getDataRange().getValues();
  var rsh=ss.getSheetByName('REGISTRATIONS');
  var regData=rsh?rsh.getDataRange().getValues():[];
  var regHdr=regData[0]||[];
  var iN=regHdr.indexOf('ชื่อเต็ม');
  var iPre=regHdr.indexOf('คำนำหน้า'),iFn=regHdr.indexOf('ชื่อ'),iLn=regHdr.indexOf('นามสกุล');
  var iOrg=regHdr.indexOf('สังกัด/หน่วยงาน');
  var iEmail=regHdr.indexOf('อีเมล'),iPhone=regHdr.indexOf('โทรศัพท์');
  var iStat=regHdr.indexOf('สถานะชำระ');

  // สร้าง lookup จาก REGISTRATIONS
  var regMap={};
  regData.slice(1).forEach(function(r){
    var rid=String(r[0]).trim();
    var nm=String(r[iN]||'').trim();
    if(!nm) nm=((String(r[iPre]||'')+' '+String(r[iFn]||'')+' '+String(r[iLn]||'')).trim());
    regMap[rid]={name:nm,org:String(r[iOrg]||''),email:String(r[iEmail]||''),
                 phone:String(r[iPhone]||''),stat:String(r[iStat]||'รอชำระเงิน')};
  });

  // สร้าง new data
  var newRows=[['REG_ID','วันที่ลงทะเบียน','ชื่อ-นามสกุล','สังกัด','อีเมล','โทรศัพท์','สถานะชำระ','วันที่บันทึก']];
  data.slice(1).forEach(function(r){
    // REG_ID — ถ้า 36 cols: col B(index 1), ถ้า 8 cols: col A(index 0)
    var rid='';
    if(data[0].length>=20) rid=String(r[1]||'').trim(); // 36 cols format
    else rid=String(r[0]||'').trim(); // 8 cols format
    if(!rid) return;
    var info=regMap[rid]||{};
    // F11: 36-col → วันที่ลงทะเบียน = col C (r[2]); 8-col → col B (r[1]) และวันที่บันทึก = r[7]
    var is36=data[0].length>=20;
    var regDate=(is36?r[2]:r[1])||new Date();
    var saved=(!is36&&r[7])?r[7]:new Date();
    newRows.push([rid,regDate,info.name||'',info.org||'',info.email||'',info.phone||'',info.stat||'รอชำระเงิน',saved]);
  });

  // เขียนใหม่
  sh.clearContents();
  sh.getRange(1,1,newRows.length,8).setValues(newRows);
  SpreadsheetApp.flush();
  ui.alert('✅ แก้ไขแล้ว '+(newRows.length-1)+' แถว');
}

// ================================================================
//  PATCH: ส่งอีเมลยืนยันชำระเงินย้อนหลัง
//  วิธีใช้:
//    1. รัน previewApprovedEmails() ก่อน → ดูรายชื่อใน Logger
//    2. ถ้าถูกต้อง → รัน resendApprovedEmails() เพื่อส่งจริง
//
//  *** ต้องมีฟังก์ชัน _sendPaymentConfirmEmail อยู่ใน script แล้ว
//      (จากไฟล์ payment_confirm_patch.gs ที่ติดตั้งก่อนหน้า) ***
// ================================================================


// ── STEP 1: ดูรายชื่อก่อนส่ง (ไม่ส่งอีเมลจริง) ─────────────────
function previewApprovedEmails() {
  var list = _getApprovedList();
  if (list.length === 0) {
    Logger.log('ไม่พบรายการที่ชำระแล้ว');
    return;
  }
  Logger.log('=== พบ ' + list.length + ' รายการที่ชำระแล้ว ===');
  list.forEach(function(r, i) {
    Logger.log((i+1) + '. ' + r.regId + ' | ' + r.name + ' | ' + r.email + ' | ' + r.fee + ' บาท | อนุมัติ: ' + r.approvedAt);
  });
  Logger.log('=== รัน resendApprovedEmails() เพื่อส่งอีเมลจริง ===');
}


// ── STEP 2: ส่งอีเมลจริงทุกคนที่ชำระแล้ว ────────────────────────
function resendApprovedEmails() {
  var ui = SpreadsheetApp.getUi();
  var list = _getApprovedList();

  if (list.length === 0) {
    ui.alert('ไม่พบรายการที่ชำระแล้ว');
    return;
  }

  var confirm = ui.alert(
    'ยืนยันการส่งอีเมลย้อนหลัง',
    'พบ ' + list.length + ' รายการที่ชำระแล้ว\nต้องการส่งอีเมลยืนยันให้ทุกคนใช่ไหม?\n\n(ดูรายชื่อได้ใน Logger จาก previewApprovedEmails)',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var sent = 0, skipped = 0, failed = 0;
  var failLog = [];
  // F10: เช็คโควต้าอีเมล + หยุดก่อนเวลา 6 นาที
  var quota = MailApp.getRemainingDailyQuota();
  if (quota < list.length) {
    var qc = ui.alert('โควต้าอีเมลไม่พอ',
      'โควต้าที่เหลือวันนี้: ' + quota + ' ฉบับ แต่มี ' + list.length + ' รายการ\nจะส่งได้ไม่เกินโควต้า ต้องการดำเนินการต่อหรือไม่?',
      ui.ButtonSet.OK_CANCEL);
    if (qc !== ui.Button.OK) return;
  }
  var t0 = Date.now();
  var stoppedAt = '';
  list.forEach(function(r, idx) {
    if (stoppedAt) return;
    if (Date.now() - t0 > 4.5 * 60 * 1000) {
      stoppedAt = 'หยุดก่อนครบเวลา ที่ลำดับ ' + (idx + 1) + '/' + list.length + ' (' + r.regId + ') — ไม่ได้ส่งตั้งแต่รายการนี้เป็นต้นไป';
      Logger.log('STOP: ' + stoppedAt);
      return;
    }
    if (sent >= quota) {
      stoppedAt = 'หยุดเพราะโควต้าอีเมลหมด ที่ลำดับ ' + (idx + 1) + '/' + list.length + ' (' + r.regId + ')';
      Logger.log('STOP: ' + stoppedAt);
      return;
    }
    if (!r.email) {
      Logger.log('SKIP (no email): ' + r.regId);
      skipped++;
      return;
    }
    try {
      _sendPaymentConfirmEmail(
        r.name,
        r.org,
        r.typeLbl,
        r.email,
        r.regId,
        r.fee,
        r.d17,
        r.d18,
        r.approvedAt
      );
      Logger.log('SENT: ' + r.regId + ' to ' + r.email);
      sent++;
      Utilities.sleep(300); // หน่วงเล็กน้อยป้องกัน quota หมด
    } catch(e) {
      Logger.log('ERROR: ' + r.regId + ' -- ' + e.message);
      failLog.push(r.regId + ' (' + e.message + ')');
      failed++;
    }
  });

  var msg = 'สำเร็จ: ' + sent + ' คน\n';
  if (skipped > 0) msg += 'ข้าม (ไม่มีอีเมล): ' + skipped + ' คน\n';
  if (failed  > 0) msg += 'ส่งไม่สำเร็จ: ' + failed + ' คน\n\n' + failLog.join('\n');
  if (stoppedAt) msg += '\n\n⚠️ ' + stoppedAt + '\n(ไม่มีระบบจำว่าส่งแล้ว — อย่ารันซ้ำทั้งหมด จะส่งซ้ำคนเดิม)';

  ui.alert('ผลการส่งอีเมล', msg, ui.ButtonSet.OK);
}


// ── Helper: ดึงรายชื่อทั้งหมดที่ชำระแล้ว ────────────────────────
function _getApprovedList() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var psh = ss.getSheetByName(SH_PAY);
  var rsh = ss.getSheetByName(SH_REG);
  if (!psh || psh.getLastRow() < 2) return [];

  // สร้าง lookup จาก REGISTRATIONS
  var regMap = {};
  if (rsh && rsh.getLastRow() > 1) {
    var rd   = rsh.getDataRange().getValues();
    var rh   = rd[0];
    var iE   = rh.indexOf('อีเมล');
    var iN   = rh.indexOf('ชื่อเต็ม');
    var iPre = rh.indexOf('คำนำหน้า');
    var iFn  = rh.indexOf('ชื่อ');
    var iLn  = rh.indexOf('นามสกุล');
    var iOrg = rh.indexOf('สังกัด/หน่วยงาน');
    var iFee = rh.indexOf('ยอดรวม');
    var iD17 = rh.indexOf('อาหาร_27พค');
    var iD18 = rh.indexOf('อาหาร_28พค');
    var iTy  = rh.indexOf('ประเภทผู้ลงทะเบียน');
    rd.slice(1).forEach(function(r) {
      var rid = String(r[0]).trim();
      if (!rid) return;
      var name = String(r[iN] || '').trim();
      if (!name) name = ((String(r[iPre]||'')+' '+String(r[iFn]||'')+' '+String(r[iLn]||'')).trim());
      regMap[rid] = {
        name   : name,
        email  : String(r[iE]   || '').trim(),
        org    : String(r[iOrg] || ''),
        fee    : Number(r[iFee] || 0),
        d17    : Number(r[iD17] || 0),
        d18    : Number(r[iD18] || 0),
        typeLbl: String(r[iTy]  || '')
      };
    });
  }

  // กรองเฉพาะ สถานะ = ชำระแล้ว จาก PAYMENTS (เอาแถวล่าสุดต่อ REG ID)
  var payMap = {};
  var pd = psh.getDataRange().getValues();
  pd.slice(1).forEach(function(r) {
    var rid    = String(r[0] || '').trim();
    var status = String(r[6] || '');
    if (!rid || status !== 'ชำระแล้ว') return;
    var approvedAt = '';
    if (r[8]) {
      approvedAt = (r[8] instanceof Date)
        ? Utilities.formatDate(r[8], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')
        : String(r[8]);
    }
    payMap[rid] = { approvedAt: approvedAt };
  });

  // รวม PAYMENTS + REGISTRATIONS
  return Object.keys(payMap).map(function(rid) {
    var reg = regMap[rid] || {};
    return {
      regId     : rid,
      name      : reg.name    || '',
      email     : reg.email   || '',
      org       : reg.org     || '',
      fee       : reg.fee     || 0,
      d17       : reg.d17     || 0,
      d18       : reg.d18     || 0,
      typeLbl   : reg.typeLbl || '',
      approvedAt: payMap[rid].approvedAt || ''
    };
  });
}

function debugRegHeaders() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var rsh = ss.getSheetByName(SH_REG);
  var hdr = rsh.getRange(1,1,1,rsh.getLastColumn()).getValues()[0];
  hdr.forEach(function(h, i) {
    Logger.log(String.fromCharCode(65+i) + ' (' + i + '): [' + h + ']');
  });
}

function debugPayHeaders() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var psh = ss.getSheetByName(SH_PAY);
  var hdr = psh.getRange(1,1,1,psh.getLastColumn()).getValues()[0];
  hdr.forEach(function(h, i) {
    Logger.log(String.fromCharCode(65+i) + ' (' + i + '): [' + h + ']');
  });
  // ดูข้อมูล 3 แถวแรกด้วย
  if (psh.getLastRow() > 1) {
    var rows = psh.getRange(2,1,Math.min(3, psh.getLastRow()-1), psh.getLastColumn()).getValues();
    rows.forEach(function(r, i) {
      Logger.log('Row '+(i+2)+': REG_ID=['+r[0]+'] สถานะ=['+r[6]+'] วันอนุมัติ=['+r[8]+']');
    });
  }
}

function debugAllPayRows() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var psh = ss.getSheetByName(SH_PAY);
  if (psh.getLastRow() < 2) { Logger.log('ไม่มีข้อมูล'); return; }
  var rows = psh.getDataRange().getValues();
  rows.slice(1).forEach(function(r, i) {
    Logger.log('Row '+(i+2)+': ['+r[0]+'] | สถานะ=['+r[6]+']');
  });
}

function findRealRegIds() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var rsh = ss.getSheetByName(SH_REG);
  var rd  = rsh.getDataRange().getValues();
  var rh  = rd[0];
  var iN  = rh.indexOf('ชื่อเต็ม');
  var iCard = rh.indexOf('เลขบัตรประชาชน');

  // ค้นหาทั้ง 4 คน + เลขบัตร
  var targets = ['3769900241621','ผกามาศ พีธรากร','ดวงหทัย ยอดทอง','พิศสมร เดชดวง','ฐิติชญา ฤทธิ์คำรพ'];

  rd.slice(1).forEach(function(r) {
    var rid  = String(r[0]).trim();
    var name = String(r[iN] || '').trim();
    var card = String(r[iCard] || '').trim();
    targets.forEach(function(t) {
      if (name.indexOf(t) >= 0 || card === t) {
        Logger.log('พบ: REG_ID=[' + rid + '] ชื่อ=[' + name + '] บัตร=[' + card + ']');
      }
    });
  });
}

// ================================================================
//  PATCH: แก้ไข REG_ID ผิดใน PAYMENTS sheet
//  รัน fixPaymentRegIds() ครั้งเดียว แล้วรัน resendApprovedEmails() ต่อได้เลย
// ================================================================

function fixPaymentRegIds() {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var psh = ss.getSheetByName(SH_PAY);
  var ui  = SpreadsheetApp.getUi();

  // mapping ค่าผิด → REG_ID ที่ถูกต้อง
  var fixMap = {
    '3769900241621'   : 'INHI26-0602-0002',  // สรัลรัตน์ พลอินทร์
    'ผกามาศ พีธรากร' : 'INHI26-0602-0003',  // ผกามาศ พีธรากร
    'ดวงหทัย ยอดทอง' : 'INHI26-0602-0004',  // ดวงหทัย ยอดทอง
    'พิศสมร เดชดวง'  : 'INHI26-0603-0005',  // พิศสมร เดชดวง
    'ฐิติชญา ฤทธิ์คำรพ': 'INHI26-0608-0006' // ฐิติชญา ฤทธิ์คำรพ
  };

  var data = psh.getDataRange().getValues();
  var fixed = 0;
  var log = [];

  for (var i = 1; i < data.length; i++) {
    var current = String(data[i][0]).trim();
    if (fixMap[current]) {
      var newId = fixMap[current];
      psh.getRange(i + 1, 1).setValue(newId);
      log.push('Row ' + (i+1) + ': [' + current + '] → [' + newId + ']');
      fixed++;
    }
  }

  SpreadsheetApp.flush();

  var msg = 'แก้ไขแล้ว ' + fixed + ' แถว\n\n' + log.join('\n');
  Logger.log(msg);
  ui.alert('แก้ไข PAYMENTS เสร็จ', msg, ui.ButtonSet.OK);
}

function testPaymentConfirmEmail() {
  _sendPaymentConfirmEmail(
    'นางสาว ทดสอบ ระบบ',
    'มหาวิทยาลัยทดสอบ',
    'ผู้เข้าร่วมทั่วไป',
    Session.getActiveUser().getEmail(), // ส่งหาผู้รันเอง
    'INHI26-TEST-0001',
    1800, 2, 2,
    '09/06/2027 10:00'
  );
}

function resendOneEmail() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('ส่งอีเมลยืนยันชำระเงิน', 'ใส่ REG ID:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var regId = resp.getResponseText().trim();
  if (!regId) { ui.alert('ไม่ได้ใส่ REG ID'); return; }

  var list = _getApprovedList().filter(function(r) {
    return r.regId === regId;
  });

  if (list.length === 0) {
    ui.alert('ไม่พบ REG ID นี้ในรายการที่ชำระแล้ว\n\n' + regId);
    return;
  }

  var r = list[0];
  if (!r.email) {
    ui.alert('ไม่พบอีเมลของ ' + r.name);
    return;
  }

  var confirm = ui.alert(
    'ยืนยันการส่งอีเมล',
    'ชื่อ: ' + r.name + '\nอีเมล: ' + r.email + '\nREG ID: ' + r.regId + '\n\nส่งอีเมลยืนยันการชำระเงินให้คนนี้?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  try {
    _sendPaymentConfirmEmail(r.name, r.org, r.typeLbl, r.email, r.regId, r.fee, r.d17, r.d18, r.approvedAt);
    ui.alert('✅ ส่งสำเร็จ\n\n' + r.name + '\n' + r.email);
    Logger.log('RESEND ONE: ' + r.regId + ' → ' + r.email);
  } catch(e) {
    ui.alert('❌ ส่งไม่สำเร็จ\n\n' + e.message);
    Logger.log('RESEND ONE ERROR: ' + r.regId + ' -- ' + e.message);
  }
}

// ── ส่งอีเมลยืนยันซ้ำ (ค้นด้วย REG ID / ชื่อ) ──
// หมายเหตุ: ฟังก์ชันนี้ใช้ ui.prompt/ui.alert ได้ เพราะรันจากเมนูในชีท "รายงานสรุป"
// (SpreadsheetApp.getUi() ใช้ได้เฉพาะตอนเปิดจาก Sheet) — ห้ามรันจาก Apps Script editor
// และห้ามเอาแพทเทิร์น alert นี้ไปใส่ในฟังก์ชันที่รันจาก editor
function resendConfirmationEmail() {
  var ui = SpreadsheetApp.getUi();
  try {
    var resp = ui.prompt('ส่งอีเมลยืนยันซ้ำ', 'ใส่ REG ID (เช่น INHI26-1023-0003 หรือ 1023-0003) หรือชื่อ/นามสกุล:', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    var q = String(resp.getResponseText() || '').trim();
    if (!q) { ui.alert('ไม่ได้ใส่ข้อมูลค้นหา'); return; }

    var ss  = SpreadsheetApp.openById(SS_ID);
    var rsh = ss.getSheetByName(SH_REG);
    if (!rsh || rsh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลในชีท ' + SH_REG); return; }
    var rd = rsh.getDataRange().getValues(), H = rd[0];
    var col = function(n) { return H.indexOf(n); };
    var iFull = col('ชื่อเต็ม'), iFn = col('ชื่อ'), iLn = col('นามสกุล'), iPre = col('คำนำหน้า');
    var iMail = col('อีเมล'), iSt = col('สถานะชำระ'), iOrg = col('สังกัด/หน่วยงาน'), iFee = col('ยอดรวม'), iTy = col('ประเภทผู้ลงทะเบียน');
    var iFree = col('ฟรี/ชำระ'), iD17 = col('อาหาร_27พค'), iD18 = col('อาหาร_28พค');

    var ql = q.toLowerCase().replace(/\s+/g, ' ');
    var qNoSp = ql.replace(/\s/g, '');
    var byId = [], byName = [];
    for (var i = 1; i < rd.length; i++) {
      var id = String(rd[i][0] || '').trim(); if (!id) continue;
      var idl = id.toLowerCase();
      var nm = String(rd[i][iFull] || '').trim() || ((String(rd[i][iPre] || '') + ' ' + String(rd[i][iFn] || '') + ' ' + String(rd[i][iLn] || '')).trim());
      var rec = {
        regId: id, name: nm, email: String(rd[i][iMail] || '').trim(), status: String(rd[i][iSt] || '').trim(),
        org: String(rd[i][iOrg] || ''), fee: Number(rd[i][iFee] || 0), typeLbl: String(rd[i][iTy] || ''),
        d17: Number(rd[i][iD17] || 0), d18: Number(rd[i][iD18] || 0), free: String(rd[i][iFree] || '') === 'ฟรี'
      };
      if (idl === qNoSp || (qNoSp.length >= 6 && idl.slice(-9) === qNoSp)) { byId.push(rec); continue; }
      var hay = [nm, String(rd[i][iFn] || ''), String(rd[i][iLn] || ''), String(rd[i][iFn] || '') + ' ' + String(rd[i][iLn] || '')]
        .map(function(s) { return s.toLowerCase().replace(/\s+/g, ' '); });
      if (hay.some(function(s) { return s.indexOf(ql) >= 0; })) byName.push(rec);
    }
    var found = byId.length ? byId : byName;

    if (found.length === 0) { ui.alert('ไม่พบผู้ลงทะเบียน\n\nคำค้น: ' + q); return; }
    if (found.length > 1) {
      var lines = found.slice(0, 10).map(function(r) { return r.regId + ' · ' + r.name + ' · ' + (r.email || '(ไม่มีอีเมล)'); });
      ui.alert('พบ ' + found.length + ' รายการ — กรุณารันใหม่โดยใส่ REG ID',
        lines.join('\n') + (found.length > 10 ? '\n… และอีก ' + (found.length - 10) + ' รายการ' : ''), ui.ButtonSet.OK);
      return;
    }

    var r = found[0];
    if (!r.email) { ui.alert('ไม่พบอีเมลของ ' + r.name + ' (' + r.regId + ')'); return; }
    var paid = r.status === 'ชำระแล้ว';
    var which = paid ? 'อีเมลยืนยันการชำระเงิน (Payment Confirmed + คูปองอาหาร)' : 'อีเมลยืนยันการลงทะเบียน (Registration Confirmed)';
    var c = ui.alert('ยืนยันการส่งอีเมลซ้ำ',
      'ชื่อ: ' + r.name + '\nอีเมล: ' + r.email + '\nREG ID: ' + r.regId + '\nสถานะ: ' + (r.status || '-') +
      '\n\nจะส่ง: ' + which + '\n\nส่งเลย?', ui.ButtonSet.OK_CANCEL);
    if (c !== ui.Button.OK) return;

    try {
      if (paid) {
        var approvedAt = '';
        var psh = ss.getSheetByName(SH_PAY);
        if (psh && psh.getLastRow() > 1) {
          var pd = psh.getDataRange().getValues(), PH = pd[0], iSt2 = PH.indexOf('สถานะ'), iAp = PH.indexOf('วันอนุมัติ');
          for (var p = pd.length - 1; p >= 1; p--) {
            if (String(pd[p][0]).trim() === r.regId && String(pd[p][iSt2]) === 'ชำระแล้ว' && pd[p][iAp]) {
              var av = pd[p][iAp];
              approvedAt = (av instanceof Date) ? Utilities.formatDate(av, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : String(av);
              break;
            }
          }
        }
        if (!approvedAt) approvedAt = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
        _sendPaymentConfirmEmail(r.name, r.org, r.typeLbl, r.email, r.regId, r.fee, r.d17, r.d18, approvedAt);
      } else {
        var d = _rebuildPayloadFromSheets(ss, r.regId);
        if (!d) throw new Error('สร้างข้อมูลลงทะเบียนจากชีทไม่ได้');
        var isFree = d.type === 'committee_dir' || d.type === 'committee_judge' || r.free;
        var typeLbl = TYPE_LBL[d.type] || r.typeLbl;
        var typeLblEn = TYPE_LBL_EN[d.type] || d.type;
        var meal = _calcMeal(d, isFree, r.name, r.org, typeLbl, r.regId);
        _sendConfirmEmail(r.name, r.org, typeLbl, typeLblEn, r.email, r.regId, r.fee, isFree, d, meal);
      }
      Logger.log('RESEND CONFIRM: ' + r.regId + ' → ' + r.email + ' (' + (paid ? 'payment' : 'registration') + ')');
      ui.alert('✅ ส่งสำเร็จ\n\n' + r.name + '\n' + r.email + '\n' + which);
    } catch (e) {
      Logger.log('RESEND CONFIRM ERROR: ' + r.regId + ' -- ' + e.message);
      ui.alert('❌ ส่งไม่สำเร็จ\n\n' + e.message);
    }
  } catch (err) {
    Logger.log('resendConfirmationEmail ERROR: ' + err.message);
    try { ui.alert('❌ เกิดข้อผิดพลาด\n\n' + err.message); } catch (e2) {}
  }
}

// สร้าง payload `d` (รูปแบบเดียวกับตอนลงทะเบียน) จากชีท REGISTRATIONS + COMPETITIONS + ชีทรายการแข่งขัน
function _rebuildPayloadFromSheets(ss, regId) {
  var id = String(regId || '').trim();
  var rsh = ss.getSheetByName(SH_REG);
  if (!rsh || rsh.getLastRow() < 2) return null;
  var rd = rsh.getDataRange().getValues(), H = rd[0], row = null;
  for (var i = 1; i < rd.length; i++) if (String(rd[i][0]).trim() === id) { row = rd[i]; break; }
  if (!row) return null;
  var g = function(n) { var k = H.indexOf(n); return k < 0 ? '' : String(row[k] == null ? '' : row[k]).trim(); };
  var typeKey = 'other';
  Object.keys(TYPE_LBL).forEach(function(k) { if (TYPE_LBL[k] === g('ประเภทผู้ลงทะเบียน')) typeKey = k; });
  var org = g('สังกัด/หน่วยงาน');
  var d = {
    type: typeKey, prefix: g('คำนำหน้า'), fname: g('ชื่อ'), lname: g('นามสกุล'),
    email: g('อีเมล'), phone: g('โทรศัพท์'), idcard: g('เลขบัตรประชาชน'),
    org: org, dirOrg: typeKey === 'committee_dir' ? org : '', judgeOrg: typeKey === 'committee_judge' ? org : '',
    allergy: g('การแพ้อาหาร'), addr: g('ที่อยู่ใบเสร็จ'), license: g('เลขใบประกอบวิชาชีพ'),
    comps: [], compDays: {}, members: {}, teamName: {}
  };
  if (typeKey !== 'other') return d; // กรรมการ: ไม่มีรายการแข่งขัน
  var csh = ss.getSheetByName(SH_COMP);
  if (csh && csh.getLastRow() > 1) {
    var cd = csh.getDataRange().getValues(), CH = cd[0];
    var iC = CH.indexOf('รหัสรายการ'), iD = CH.indexOf('จำนวนวัน'), iO = CH.indexOf('ลำดับ');
    var list = [];
    for (var j = 1; j < cd.length; j++) if (String(cd[j][0]).trim() === id) list.push({ cid: String(cd[j][iC]).trim(), days: Number(cd[j][iD]) === 2 ? 2 : 1, o: Number(cd[j][iO]) || 0 });
    list.sort(function(a, b) { return a.o - b.o; });
    list.forEach(function(x) { d.comps.push(x.cid); d.compDays[x.cid] = x.days; });
  }
  d.comps.forEach(function(cid) {
    if (cid === 'c00' || !COMP_SHEET[cid]) return;
    var sh = ss.getSheetByName(COMP_SHEET[cid]); if (!sh || sh.getLastRow() < 2) return;
    var top = sh.getRange(1, 1, Math.min(3, sh.getLastRow()), sh.getLastColumn()).getValues(), hr = -1;
    for (var t = 0; t < top.length && hr < 0; t++) if (top[t].indexOf('REG_ID') >= 0) hr = t;
    if (hr < 0 || sh.getLastRow() <= hr + 1) return;
    var v = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, sh.getLastColumn()).getValues(), CH2 = v[0], iR = CH2.indexOf('REG_ID'), iTn = CH2.indexOf('ชื่อทีม');
    for (var k = 1; k < v.length; k++) if (String(v[k][iR]).trim() === id) {
      var gg = function(x) { return x < 0 ? '' : String(v[k][x] == null ? '' : v[k][x]).trim(); };
      var mem = [];
      for (var n = 1; n <= 4; n++) {
        var iFn = CH2.indexOf('ชื่อ ' + n); if (iFn < 0) iFn = CH2.indexOf('ชื่อ สมาชิก ' + n);
        mem.push({ prefix: gg(CH2.indexOf('คำนำหน้า ' + n)), fname: gg(iFn), lname: gg(CH2.indexOf('นามสกุล ' + n)),
                   idcard: gg(CH2.indexOf('เลขบัตร ' + n)), license: gg(CH2.indexOf('ใบประกอบ ' + n)), allergy: gg(CH2.indexOf('แพ้อาหาร ' + n)) });
      }
      d.members[cid] = mem;
      d.teamName[cid] = gg(iTn);
      break;
    }
  });
  return d;
}

// ── รันครั้งเดียวเพื่อตรวจสอบและแก้ไข REGISTRATIONS header + data ──
function diagnoseAndFixRegistrations() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH_REG);
  var ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('ไม่พบ REGISTRATIONS sheet'); return; }

  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var msg = 'Header ปัจจุบัน (' + hdr.length + ' cols):\n';
  hdr.forEach(function(h, i) {
    msg += String.fromCharCode(65 + i) + ': [' + h + ']\n';
  });

  var iStat = hdr.indexOf('สถานะชำระ');
  msg += '\nสถานะชำระ อยู่ col index: ' + iStat + ' (' + (iStat >= 0 ? String.fromCharCode(65 + iStat) : 'ไม่พบ!') + ')';

  // ตรวจว่า _updatePayStatus จะเขียนถูก col ไหม
  var expected = 17; // col R (0-based) ตาม setupSheets ล่าสุด
  if (iStat === expected) {
    msg += '\n✅ Header ตรงแล้ว — _updatePayStatus จะทำงานถูกต้อง';
  } else if (iStat >= 0) {
    msg += '\n⚠️ Header ไม่ตรง! indexOf คืน ' + iStat + ' แต่ hardcode fallback = ' + expected;
    msg += '\n→ _updatePayStatus อาจเขียนผิด col เมื่อ indexOf fail';
  } else {
    msg += '\n❌ ไม่พบ header สถานะชำระ เลย!';
  }

  ui.alert('Diagnose Result', msg, ui.ButtonSet.OK);
}

function testUpdatePayStatus() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('ใส่ REG ID ที่ชำระแล้ว', 'เช่น INHI26-0521-0004', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var regId = resp.getResponseText().trim();
  if (!regId) { ui.alert('ไม่ได้ใส่ REG ID'); return; }

  var ss = SpreadsheetApp.openById(SS_ID);
  // ตรวจว่ามี REG_ID นี้จริง + ขอยืนยันก่อนเขียน (ฟังก์ชันนี้แก้ข้อมูลจริง)
  var rsh0 = ss.getSheetByName(SH_REG);
  var exists = false;
  if (rsh0 && rsh0.getLastRow() > 1) {
    rsh0.getRange(2, 1, rsh0.getLastRow() - 1, 1).getValues().forEach(function(r0) {
      if (String(r0[0]).trim() === regId) exists = true;
    });
  }
  if (!exists) { ui.alert('ไม่พบ REG ID ใน REGISTRATIONS: ' + regId); return; }
  var cf = ui.alert('ยืนยัน', 'จะเปลี่ยนสถานะ ' + regId + ' เป็น "ชำระแล้ว" ใน REGISTRATIONS และชีทรายการ\n(ไม่แก้ PAYMENTS / ไม่ส่งอีเมล)\n\nดำเนินการต่อ?', ui.ButtonSet.OK_CANCEL);
  if (cf !== ui.Button.OK) return;
  _updatePayStatus(ss, regId, 'ชำระแล้ว');
  SpreadsheetApp.flush();

  // ตรวจผล
  var sh = ss.getSheetByName(SH_REG);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === regId) {
      ui.alert('ผล', 'REG ID: ' + regId + '\nสถานะชำระ (col R): [' + data[i][17] + ']', ui.ButtonSet.OK);
      return;
    }
  }
  ui.alert('ไม่พบ REG ID: ' + regId);
}

function diagApprove() {
  var ss = SpreadsheetApp.openById(SS_ID);
  
  // 1. ดู PAYMENTS — มี REG ID อะไรบ้าง + สถานะ
  var psh = ss.getSheetByName(SH_PAY);
  var pd = psh.getDataRange().getValues();
  Logger.log('=== PAYMENTS ===');
  pd.slice(1).forEach(function(r, i) {
    Logger.log('Row '+(i+2)+': REG_ID=['+r[0]+'] สถานะ=['+r[6]+']');
  });
  
  // 2. ดู REGISTRATIONS — col A มีอะไร
  var rsh = ss.getSheetByName(SH_REG);
  var rd = rsh.getDataRange().getValues();
  Logger.log('=== REGISTRATIONS (col A) ===');
  rd.slice(1).forEach(function(r, i) {
    Logger.log('Row '+(i+2)+': ['+r[0]+']');
  });
  
  // 3. ดู header REGISTRATIONS
  Logger.log('=== REG HEADERS ===');
  rd[0].forEach(function(h, i) {
    Logger.log(String.fromCharCode(65+i)+' ('+i+'): ['+h+']');
  });
}

function fixPaymentRegIds2() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var psh = ss.getSheetByName(SH_PAY);
  var rsh = ss.getSheetByName(SH_REG);
  var ui = SpreadsheetApp.getUi();

  // สร้าง lookup จาก REGISTRATIONS: ชื่อ/เลขบัตร → REG_ID
  var rd = rsh.getDataRange().getValues();
  var rh = rd[0];
  var iName = rh.indexOf('ชื่อเต็ม');
  var iFn = rh.indexOf('ชื่อ');
  var iCard = rh.indexOf('เลขบัตรประชาชน');

  var iPre = rh.indexOf('คำนำหน้า');
  var iLn = rh.indexOf('นามสกุล');
  function normNm(x) { return String(x || '').replace(/\s+/g, '').toLowerCase(); }

  var nameMap = {};  // ชื่อเต็มที่ normalize → [rid,...]
  var cardMap = {};  // เลขบัตร → rid

  function addName(nm, rid) {
    nm = normNm(nm);
    if (!nm) return;
    if (!nameMap[nm]) nameMap[nm] = [];
    if (nameMap[nm].indexOf(rid) < 0) nameMap[nm].push(rid);
  }

  rd.slice(1).forEach(function(r) {
    var rid = String(r[0]).trim();
    if (!rid) return;
    var fullName = iName >= 0 ? String(r[iName] || '').trim() : '';
    var fname = iFn >= 0 ? String(r[iFn] || '').trim() : '';
    var lname = iLn >= 0 ? String(r[iLn] || '').trim() : '';
    var pre = iPre >= 0 ? String(r[iPre] || '').trim() : '';
    var card = String(r[iCard] || '').trim();
    addName(fullName, rid);
    if (fname && lname) { addName(pre + fname + lname, rid); addName(fname + lname, rid); }
    if (card) cardMap[card] = rid;
  });

  // แก้ PAYMENTS (ต้องตรงทั้งชื่อ+นามสกุล; ถ้าซ้ำหลายคนให้ข้าม)
  var pd = psh.getDataRange().getValues();
  var fixed = 0;
  var log = [];

  for (var i = 1; i < pd.length; i++) {
    var current = String(pd[i][0]).trim();
    // ถ้าไม่ใช่ format INHI26-xxxx-xxxx
    if (!/^INHI26-/.test(current)) {
      var cands = nameMap[normNm(current)] || [];
      var newId = null;
      if (cands.length === 1) newId = cands[0];
      else if (cands.length === 0) newId = cardMap[current] || null;
      if (cands.length > 1) {
        log.push('Row ' + (i+1) + ': [' + current + '] → ⚠️ ชื่อซ้ำหลายคน (' + cands.join(', ') + ') — ข้าม');
      } else if (newId) {
        psh.getRange(i + 1, 1).setValue(newId);
        log.push('Row ' + (i+1) + ': [' + current + '] → [' + newId + ']');
        fixed++;
      } else {
        log.push('Row ' + (i+1) + ': [' + current + '] → ❌ หา REG ID ไม่เจอ');
      }
    }
  }

  SpreadsheetApp.flush();
  var msg = 'แก้แล้ว ' + fixed + ' แถว\n\n' + log.join('\n');
  Logger.log(msg);
  ui.alert('แก้ PAYMENTS เสร็จ', msg, ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: ชีทกรรมการ (แยก) + ชีทรายชื่ออาหารรายบุคคล
//  รันด้วยตนเองจาก GAS Editor เมื่อต้องการอัปเดตข้อมูลล่าสุด
//  ไม่แก้ไข/ไม่ลบข้อมูลในชีทอื่นใดๆ ทั้งสิ้น
// ================================================================

var SH_DIR      = 'กรรมการอำนวยการ';
var SH_JUDGE    = 'กรรมการตัดสินการแข่งขัน';
var SH_MEALLIST = 'รายชื่ออาหารรายบุคคล';

// ── ชีทกรรมการอำนวยการ / กรรมการตัดสินการแข่งขัน (แยกจาก REGISTRATIONS) ──
// silent=true → ไม่แสดง popup (ใช้ตอนถูกเรียกอัตโนมัติจาก _registerUser)
function exportCommitteeSheets(silent) {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var reg = ss.getSheetByName(SH_REG);
  if (!reg || reg.getLastRow() < 2) {
    if (!silent) SpreadsheetApp.getUi().alert('ไม่พบข้อมูลใน REGISTRATIONS');
    return;
  }

  var data = reg.getDataRange().getValues();
  var hdr  = data[0];
  var iType  = hdr.indexOf('ประเภทผู้ลงทะเบียน');
  var iDate  = hdr.indexOf('วันที่ลงทะเบียน');
  var iPre   = hdr.indexOf('คำนำหน้า');
  var iFn    = hdr.indexOf('ชื่อ');
  var iLn    = hdr.indexOf('นามสกุล');
  var iEmail = hdr.indexOf('อีเมล');
  var iPhone = hdr.indexOf('โทรศัพท์');
  var iOrg   = hdr.indexOf('สังกัด/หน่วยงาน');
  var iComp  = hdr.indexOf('ตัดสินประเภท');
  var iAl    = hdr.indexOf('การแพ้อาหาร');

  var dirHdr   = ['REG_ID','วันที่ลงทะเบียน','คำนำหน้า','ชื่อ','นามสกุล','อีเมล','โทรศัพท์','สังกัด/หน่วยงาน','การแพ้อาหาร'];
  var judgeHdr = ['REG_ID','วันที่ลงทะเบียน','คำนำหน้า','ชื่อ','นามสกุล','อีเมล','โทรศัพท์','สังกัด/หน่วยงาน','ตัดสินรายการ','การแพ้อาหาร'];

  var dirRows = [], judgeRows = [];

  for (var i = 1; i < data.length; i++) {
    var row  = data[i];
    var type = String(row[iType] || '').trim();
    if (type === TYPE_LBL.committee_dir) {
      dirRows.push([row[0], row[iDate], row[iPre], row[iFn], row[iLn],
                     row[iEmail], row[iPhone], row[iOrg], row[iAl]]);
    } else if (type === TYPE_LBL.committee_judge) {
      judgeRows.push([row[0], row[iDate], row[iPre], row[iFn], row[iLn],
                       row[iEmail], row[iPhone], row[iOrg], row[iComp], row[iAl]]);
    }
  }

  _rebuildReportSheet(ss, SH_DIR, dirHdr, dirRows);
  _rebuildReportSheet(ss, SH_JUDGE, judgeHdr, judgeRows);

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      'อัปเดตชีทกรรมการเสร็จแล้ว\n\n' +
      'กรรมการอำนวยการ: ' + dirRows.length + ' คน\n' +
      'กรรมการตัดสินการแข่งขัน: ' + judgeRows.length + ' คน'
    );
  }
}

// ── ชีทรายชื่ออาหารรายบุคคล: แยกตามรายการแข่งขัน/ทีม/วัน + สรุปยอดรวม 27/28 กค ──
// ★ ทำงานได้ด้วยตัวเอง 100% ไม่ต้องพึ่งฟังก์ชันอื่น ★
// silent=true → ไม่แสดง popup (ใช้ตอนถูกเรียกอัตโนมัติจาก _registerUser)
function exportMealListByCompetition(silent) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var rows = [];
  var debugLog = [];
  var incompleteTeams = []; // ทีมที่กรอกสมาชิกไม่ครบโควต้าที่จ่ายเงินไว้
  var suspiciousBlankRows = []; // แถวที่ REG_ID อ่านไม่ได้ แต่มีข้อมูลคอลัมน์อื่นอยู่ (รวมทุกห้อง)

  Object.keys(COMP_SHEET).forEach(function(cid) {
    var cfg = COMP[cid];
    if (!cfg || cfg.mealDay === 'none') return; // c00 ออนไลน์ไม่มีอาหาร

    var shName = COMP_SHEET[cid];
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) { debugLog.push(shName + ': ไม่มีข้อมูล'); return; }

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1;
    var hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0) { debugLog.push(shName + ': หา header (REG_ID) ไม่เจอ'); return; }
    if (sh.getLastRow() <= headerRowIdx) { debugLog.push(shName + ': มี header แต่ไม่มีข้อมูล'); return; }

    var iRegId = hdr.indexOf('REG_ID');
    var iTeam  = hdr.indexOf('ชื่อทีม');
    var iDays  = hdr.indexOf('จำนวนวัน');

    // รองรับ header ได้ทั้ง 2 รูปแบบ: แยกคำนำหน้า/ชื่อ/นามสกุล หรือ "ชื่อ สมาชิก N" รวมช่องเดียว
    var slotCols = [];
    for (var n = 1; n <= 4; n++) {
      var pre = hdr.indexOf('คำนำหน้า ' + n);
      var fn  = hdr.indexOf('ชื่อ ' + n);
      var ln  = hdr.indexOf('นามสกุล ' + n);
      var combined = hdr.indexOf('ชื่อ สมาชิก ' + n);
      var al  = hdr.indexOf('แพ้อาหาร ' + n);
      slotCols.push({ pre: pre, fn: fn, ln: ln, combined: combined, al: al });
    }

    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
    var foundInSheet = 0;
    var quota = cfg.team || 1;

    data.forEach(function(r, rIdx) {
      var regId = iRegId >= 0 ? r[iRegId] : '';
      if (!String(regId||'').trim()) {
        // เช็คว่าแถวนี้ "ว่างจริง" หรือมีข้อมูลอื่นอยู่แต่ REG_ID อ่านไม่ได้ (เช่น cell ถูก merge/พิมพ์ผิดคอลัมน์)
        var hasOtherData = r.some(function(cellVal){ return String(cellVal||'').trim() !== ''; });
        if (hasOtherData) {
          var sheetRowNum = headerRowIdx + 1 + rIdx;
          suspiciousBlankRows.push(shName + ' แถวที่ ' + sheetRowNum + ' | REG_ID อ่านไม่ได้ แต่มีข้อมูลอื่น: [' +
            r.slice(0, 6).map(function(v){ return String(v||'').trim() || '-'; }).join(' | ') + ' ...]');
        }
        return; // แถวว่างเปล่า ข้าม
      }
      var teamName = iTeam >= 0 ? (r[iTeam] || '') : '';
      var days = iDays >= 0 ? (Number(r[iDays]) || 1) : 1;
      var getsD17 = cfg.mealDay === 'both' || cfg.mealDay === 'd17' || (cfg.mealDay === 'd18' && days === 2);
      var getsD18 = cfg.mealDay === 'both' || cfg.mealDay === 'd18' || (cfg.mealDay === 'd17' && days === 2);
      var dayLbl = (getsD17 ? '27 พ.ค.' : '') + (getsD17 && getsD18 ? ' + ' : '') + (getsD18 ? '28 พ.ค.' : '');

      var filledInTeam = 0;
      slotCols.forEach(function(col, idx) {
        var fullName = '';
        var al = '';
        if (col.fn >= 0) {
          var pre = col.pre >= 0 ? r[col.pre] : '';
          var ln  = col.ln  >= 0 ? r[col.ln]  : '';
          fullName = (String(pre||'')+' '+String(r[col.fn]||'')+' '+String(ln||'')).trim();
          al = col.al >= 0 ? r[col.al] : '';
        } else if (col.combined >= 0) {
          fullName = String(r[col.combined] || '').trim();
          al = col.al >= 0 ? r[col.al] : '';
        } else {
          return; // ชีทนี้ไม่มีคอลัมน์ของสมาชิกคนนี้เลย
        }
        if (!fullName) return; // ข้ามช่องที่ไม่มีชื่อจริง
        foundInSheet++;
        filledInTeam++;
        rows.push([cfg.nm, teamName, regId, idx+1, fullName, dayLbl, al || '', getsD17, getsD18]);
      });

      // ทีมนี้กรอกสมาชิกไม่ครบโควต้าที่จ่ายเงินไว้ → บันทึกไว้แจ้งเตือน
      if (filledInTeam < quota) {
        incompleteTeams.push([cfg.nm, teamName || '(ไม่ระบุชื่อทีม)', regId, quota, filledInTeam, quota - filledInTeam]);
      }
    });

    debugLog.push(shName + ': ดึงได้ ' + foundInSheet + ' คน');
  });

// กรรมการ (ได้อาหารทั้ง 2 วันเสมอ) — แยกห้องด้วยถ้าเป็นกรรมการตัดสิน
  var reg = ss.getSheetByName(SH_REG);
  if (reg && reg.getLastRow() > 1) {
    var rdata = reg.getDataRange().getValues();
    var rhdr  = rdata[0];
    var iType = rhdr.indexOf('ประเภทผู้ลงทะเบียน');
    var iName = rhdr.indexOf('ชื่อเต็ม');
    var iOrg  = rhdr.indexOf('สังกัด/หน่วยงาน');
    var iAl   = rhdr.indexOf('การแพ้อาหาร');
    var iComp = rhdr.indexOf('ตัดสินประเภท');
    for (var i = 1; i < rdata.length; i++) {
      var type = String(rdata[i][iType] || '').trim();
      if (type === TYPE_LBL.committee_dir || type === TYPE_LBL.committee_judge) {
        var category = type;
        if (type === TYPE_LBL.committee_judge && iComp >= 0) {
          var judgeRoom = String(rdata[i][iComp] || '').trim();
          if (judgeRoom) category = type + ' — ' + judgeRoom;
        }
        rows.push([category, rdata[i][iOrg] || '', rdata[i][0], 1,
                   rdata[i][iName], '27 พ.ค. + 28 พ.ค.', rdata[i][iAl] || '', true, true]);
      }
    }
  }

// ── สรุปยอดรวมต่อรายการ ──────────────────────────────────
  var summaryMap = {}; // { รายการ: {d17:N, d18:N, teamSet:{regId:true}} }
  var grand17 = 0, grand18 = 0, grandTeams = 0;
  rows.forEach(function(r) {
    var nm = r[0];
    var regId = r[2];
    if (!summaryMap[nm]) summaryMap[nm] = { d17: 0, d18: 0, teamSet: {} };
    if (regId && !summaryMap[nm].teamSet[regId]) summaryMap[nm].teamSet[regId] = true;
    if (r[7]) { summaryMap[nm].d17++; grand17++; }
    if (r[8]) { summaryMap[nm].d18++; grand18++; }
  });

  // ── เขียนตารางหลัก (คอลัมน์ A-G เท่านั้น ตัด flag d17/d18 ออก) ──
  var headers = ['รายการแข่งขัน/ประเภท','ชื่อทีม/สังกัด','REG_ID','ลำดับ','ชื่อ-นามสกุล','วันที่ได้รับอาหาร','การแพ้อาหาร'];
  var sh = ss.getSheetByName(SH_MEALLIST);
  if (!sh) {
    sh = ss.insertSheet(SH_MEALLIST);
  } else {
    sh.clearContents();
    sh.clearFormats();
  }
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length)
    .setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);

  var dataRows = rows.map(function(r){ return r.slice(0,7); }); // ตัด flag ออก เหลือ 7 คอลัมน์
  if (dataRows.length) {
    sh.getRange(2,1,dataRows.length,headers.length).setValues(dataRows);
  }
  sh.autoResizeColumns(1, headers.length);

  // ── เขียนตารางสรุปยอด ต่อท้าย 2 แถวว่าง (ตำแหน่งขยับอัตโนมัติตามจำนวนแถวจริง) ──
  var sumStartRow = dataRows.length + 3;
  sh.getRange(sumStartRow, 1).setValue('สรุปจำนวนอาหารตามรายการ')
    .setFontWeight('bold').setFontSize(12);

var sumHeaderRow = sumStartRow + 1;
  sh.getRange(sumHeaderRow, 1, 1, 4).setValues([['รายการ/ประเภท','จำนวนทีม','27 พ.ค. (คน)','28 พ.ค. (คน)']])
    .setBackground('#162544').setFontColor('#fff').setFontWeight('bold');

  var sumKeys = Object.keys(summaryMap).sort();
  var sumRows = sumKeys.map(function(k){
    var teamCount = Object.keys(summaryMap[k].teamSet).length;
    grandTeams += teamCount;
    return [k, teamCount, summaryMap[k].d17, summaryMap[k].d18];
  });
  if (sumRows.length) {
    sh.getRange(sumHeaderRow + 1, 1, sumRows.length, 4).setValues(sumRows);
  }

  var totalRow = sumHeaderRow + 1 + sumRows.length;
  sh.getRange(totalRow, 1, 1, 4).setValues([['รวมทั้งหมด', grandTeams, grand17, grand18]])
    .setBackground('#f0fdf4').setFontWeight('bold');

  sh.autoResizeColumns(1, 4);

  // ── ตารางทีมที่กรอกสมาชิกไม่ครบโควต้าที่จ่ายเงินไว้ (เว้น 2 แถวถัดจากตารางสรุป) ──
  var warnStartRow = totalRow + 3;
  sh.getRange(warnStartRow, 1).setValue('⚠️ ทีมที่กรอกข้อมูลสมาชิกไม่ครบโควต้า (ควรทวงถามให้กรอกเพิ่ม)')
    .setFontWeight('bold').setFontSize(12).setFontColor('#b45309');

  var warnHeaderRow = warnStartRow + 1;
  var warnHeaders = ['รายการแข่งขัน','ชื่อทีม','REG_ID','โควต้า (คน)','กรอกแล้ว (คน)','ขาดอีก (คน)'];
  sh.getRange(warnHeaderRow, 1, 1, warnHeaders.length).setValues([warnHeaders])
    .setBackground('#78350f').setFontColor('#fff').setFontWeight('bold');

  if (incompleteTeams.length) {
    sh.getRange(warnHeaderRow + 1, 1, incompleteTeams.length, warnHeaders.length).setValues(incompleteTeams);
    sh.getRange(warnHeaderRow + 1, 1, incompleteTeams.length, warnHeaders.length).setBackground('#fffbeb');
  } else {
    sh.getRange(warnHeaderRow + 1, 1).setValue('✅ ทุกทีมกรอกสมาชิกครบโควต้าแล้ว');
  }
  sh.autoResizeColumns(1, warnHeaders.length);

  Logger.log(debugLog.join('\n'));
  if (suspiciousBlankRows.length) {
    Logger.log('\n=== ⚠️ แถวที่หลุดจากรายชื่ออาหาร (REG_ID อ่านไม่ได้ แต่มีข้อมูลอื่นอยู่) ===\n' + suspiciousBlankRows.join('\n'));
  }
  if (!silent) {
    SpreadsheetApp.getUi().alert(
      'อัปเดตชีท "รายชื่ออาหารรายบุคคล" เสร็จแล้ว',
      'รวมทั้งหมด: ' + rows.length + ' คน\n' +
      '27 พฤษภาคม: ' + grand17 + ' คน\n' +
      '28 พฤษภาคม: ' + grand18 + ' คน\n' +
      'ทีมที่กรอกไม่ครบโควต้า: ' + incompleteTeams.length + ' ทีม\n' +
      (suspiciousBlankRows.length ? '⚠️ แถวที่หลุดหายเพราะ REG_ID อ่านไม่ได้: ' + suspiciousBlankRows.length + ' แถว (ดูรายละเอียดที่ View > Logs)\n' : '') +
      '\nดูตารางสรุป/รายชื่อทีมไม่ครบโควต้าได้ที่ท้ายชีท (แถวที่ ' + sumStartRow + ' เป็นต้นไป)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ลบเนื้อหาเดิมเฉพาะ "ชีทที่ระบุ" แล้วเขียนใหม่ทั้งหมด (ไม่แตะชีทอื่น)
function _rebuildReportSheet(ss, name, headers, rows) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clearContents();
    sh.clearFormats();
  }
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length)
    .setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  }
  sh.autoResizeColumns(1, headers.length);
}

// ── DEBUG: ตรวจสอบว่าชีทรายการแข่งขันแต่ละอันมีอยู่จริงไหม มีกี่แถว ──
function debugCompSheetsStatus() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var msg = [];
  Object.keys(COMP_SHEET).forEach(function(cid) {
    var name = COMP_SHEET[cid];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      msg.push(name + ' → ยังไม่มีชีทนี้');
    } else {
      msg.push(name + ' → มีอยู่ | แถวทั้งหมด: ' + sh.getLastRow() +
                ' | ทีมที่ลงทะเบียน: ' + Math.max(0, sh.getLastRow() - 2));
    }
  });
  SpreadsheetApp.getUi().alert('สถานะชีทรายการแข่งขัน', msg.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── ตัดแถวฟอร์แมตเปล่าๆ ท้ายชีทออก (ไม่กระทบข้อมูลจริง) ──────
// เกิดจากตอน resetAndSetupSheets ตั้งฟอร์แมตวันที่ล่วงหน้าไว้ถึง 2000 แถว
// ทำให้ getLastRow() รายงานแถวเกินจริง — ฟังก์ชันนี้ตัดเฉพาะแถวว่างส่วนเกินออก
function trimEmptyFormattedRows() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var targets = Object.keys(COMP_SHEET).map(function(cid){ return COMP_SHEET[cid]; });
  var log = [];

  targets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { log.push(name + ' → ไม่พบชีท'); return; }

    var lastRow = sh.getLastRow();
    if (lastRow < 3) { log.push(name + ' → ไม่มีข้อมูล'); return; }

    // F13: หาแถว header จริง (สแกนแถว 1-3) และห้ามลบแถวที่มีค่าใดๆ
    var hdrRow = 0;
    var lastCol0 = sh.getLastColumn();
    for (var hr = 1; hr <= Math.min(3, lastRow); hr++) {
      var cand = sh.getRange(hr, 1, 1, lastCol0).getValues()[0];
      if (cand.indexOf('REG_ID') >= 0) { hdrRow = hr; break; }
    }
    if (!hdrRow) { log.push(name + ' → ไม่พบแถว header (REG_ID) — ข้าม'); return; }
    if (lastRow <= hdrRow) { log.push(name + ' → ไม่มีข้อมูล'); return; }

    var all = sh.getRange(1, 1, lastRow, lastCol0).getValues();
    var realLastRow = hdrRow;
    for (var i = hdrRow; i < all.length; i++) {
      var hasVal = all[i].some(function(v) { return String(v === null || v === undefined ? '' : v).trim() !== ''; });
      if (hasVal) realLastRow = i + 1;
    }

    if (realLastRow < lastRow) {
      sh.deleteRows(realLastRow + 1, lastRow - realLastRow);
      log.push(name + ' → ตัดจากแถว ' + lastRow + ' เหลือ ' + realLastRow + ' (ลบฟอร์แมตเปล่า ' + (lastRow - realLastRow) + ' แถว)');
    } else {
      log.push(name + ' → ไม่มีแถวเกิน (' + lastRow + ' แถว เป็นข้อมูลจริงทั้งหมด)');
    }
  });

  SpreadsheetApp.getUi().alert('ผลการตัดแต่งชีท', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ================================================================
//  ADD-ON: ชีทสรุปการลงทะเบียนทั้งหมด (ทุกห้อง/ทุกรายการแข่งขัน/ทุกกรรมการ/ทุกคน)
//  รันด้วยตนเองจาก GAS Editor — ไม่แก้ไข/ไม่ลบข้อมูลในชีทอื่นใดๆ ทั้งสิ้น
//  สร้าง/อัปเดตเฉพาะชีท "สรุปการลงทะเบียนทั้งหมด" เท่านั้น
// ================================================================

var SH_FULLSUMMARY = 'สรุปการลงทะเบียนทั้งหมด';

// silent=true → ไม่แสดง popup (เผื่อจะเรียกอัตโนมัติในอนาคต)
// ================================================================
//  REPLACEMENT: exportFullRegistrationSummary
//  แก้ไข 2 เรื่อง:
//   1. บั๊กสถานะ "ฟรี" เคยตกไปนับเป็น "รอดำเนินการ" ผิด → แก้ให้นับเป็น
//      "ชำระแล้ว/จบแล้ว" ถูกต้อง (ทั้งจุด c00 ออนไลน์ และรายการแข่งขันทั่วไป)
//   2. แยกยอดเงินเป็น 2 คอลัมน์:
//      - "ยอดเงินเก็บได้จริง (บาท)"   → นับเฉพาะสถานะ "ชำระแล้ว" เท่านั้น
//      - "ยอดเงินรวมทั้งหมด (บาท)"   → รวมทุกสถานะ (พฤติกรรมเดิม)
//
//  วิธีติดตั้ง: ลบฟังก์ชัน exportFullRegistrationSummary เวอร์ชันเก่าออกทั้งหมด
//  (ตั้งแต่ "function exportFullRegistrationSummary(silent) {" บรรทัดแรก
//   ไปจนถึง "}" ปิดท้ายของฟังก์ชัน) แล้ววางฟังก์ชันนี้แทนที่ตำแหน่งเดิม
//  ไม่ต้องแตะ var SH_FULLSUMMARY = 'สรุปการลงทะเบียนทั้งหมด'; ที่อยู่ด้านบนฟังก์ชัน — ใช้ตัวเดิมได้เลย
// ================================================================

function exportFullRegistrationSummary(silent) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var people = []; // รายชื่อทุกคนแบบละเอียด
  var roomStats = {}; // สถิติต่อห้อง/รายการ

  function ensureRoom(name, type) {
    if (!roomStats[name]) roomStats[name] = {
      type: type, teams: 0, quota: 0, named: 0,
      feeCollected: 0, feeTotal: 0,
      paid: 0, pending: 0, rejected: 0
    };
    return roomStats[name];
  }

  // ── สถานะชำระที่แท้จริง — ดึงจาก REGISTRATIONS โดยตรง ──
  // อ่านจากคอลัมน์ "สถานะอัพเดทอัตโนมัติตาม PAYMENTS" (V) ก่อน — ค่าอัปเดตสดเสมอ
  // ถ้าไม่เจอ fallback ไปคอลัมน์ "สถานะชำระ" (R) เหมือนเดิม
  var regStatusMap = {};
  var regSheetForStatus = ss.getSheetByName(SH_REG);
  if (regSheetForStatus && regSheetForStatus.getLastRow() > 1) {
    var statusData = regSheetForStatus.getDataRange().getValues();
    var statusHdr  = statusData[0];
    var iRegStatus = -1;
for (var _h = 0; _h < statusHdr.length; _h++) { if (String(statusHdr[_h]).indexOf('สถานะอัพเดท') >= 0) { iRegStatus = _h; break; } }
    if (iRegStatus < 0) iRegStatus = statusHdr.indexOf('สถานะชำระ');
    if (iRegStatus >= 0) {
      for (var si = 1; si < statusData.length; si++) {
        var srid = String(statusData[si][0] || '').trim();
        if (srid) regStatusMap[srid] = String(statusData[si][iRegStatus] || '').trim();
      }
    }
  }

  // ── 1) รายการแข่งขันทุกอัน (รวม c00/c00s) ──────────────────
  Object.keys(COMP_SHEET).forEach(function(cid) {
    var cfg = COMP[cid];
    if (!cfg) return;
    var shName = COMP_SHEET[cid];
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0 || sh.getLastRow() <= headerRowIdx) return;

    var iRegId  = hdr.indexOf('REG_ID');
    var iTeam   = hdr.indexOf('ชื่อทีม');
    var iOrg    = hdr.indexOf('สังกัด');
    var iFee    = hdr.indexOf('ค่าธรรมเนียม (บาท)');
    var quota   = cfg.team || 1;
    var stat    = ensureRoom(cfg.nm, 'รายการแข่งขัน');

    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();

    if (cid === 'c00') {
      // c00 ออนไลน์ — 8 คอลัมน์ ไม่มีทีม ชื่อรวมช่องเดียว ไม่มีคอลัมน์ค่าธรรมเนียม
      var iName = hdr.indexOf('ชื่อ-นามสกุล');
      data.forEach(function(r) {
        var regId = String(r[iRegId] || '').trim();
        if (!regId) return;
        var fullName = String(iName >= 0 ? r[iName] : '' || '').trim();
        var org = iOrg >= 0 ? (r[iOrg] || '') : '';
        var status = regStatusMap[regId] || 'รอชำระเงิน';
        stat.teams++; stat.quota += quota;
        // F14: ออนไลน์ไม่มีคอลัมน์ค่าธรรมเนียม ใช้ราคา d1 ของรายการ
        stat.feeTotal += (cfg.d1 || 0);
        if (status === 'ชำระแล้ว') stat.feeCollected += (cfg.d1 || 0);
        if (fullName) { stat.named++; people.push([cfg.nm, 'ผู้เข้าร่วมทั่วไป', org, regId, 1, '', fullName, '', '', '', '', status]); }
        if (status === 'ชำระแล้ว' || status === 'ฟรี') stat.paid++;
        else if (status === 'ปฏิเสธ') stat.rejected++;
        else stat.pending++;
      });
      return;
    }

    // รูปแบบทีมทั่วไป — รองรับ header ได้ทั้ง 2 แบบเหมือน exportMealListByCompetition
    var slotCols = [];
    for (var n = 1; n <= 4; n++) {
      slotCols.push({
        pre: hdr.indexOf('คำนำหน้า ' + n),
        fn:  hdr.indexOf('ชื่อ ' + n),
        ln:  hdr.indexOf('นามสกุล ' + n),
        combined: hdr.indexOf('ชื่อ สมาชิก ' + n),
        idc: hdr.indexOf('เลขบัตร ' + n),
        lic: hdr.indexOf('ใบประกอบ ' + n),
        al:  hdr.indexOf('แพ้อาหาร ' + n)
      });
    }

    data.forEach(function(r) {
      var regId = String(r[iRegId] || '').trim();
      if (!regId) return;
      var teamName = iTeam >= 0 ? (r[iTeam] || '') : '';
      var fee = iFee >= 0 ? (Number(r[iFee]) || 0) : 0;
      var status = regStatusMap[regId] || 'รอชำระเงิน';

      stat.teams++; stat.quota += quota;
      stat.feeTotal += fee; // ยอดรวมทั้งหมด — นับทุกสถานะ (ไว้ดูยอดค้างรวม)
      if (status === 'ชำระแล้ว' || status === 'ฟรี') {
        stat.paid++;
        if (status === 'ชำระแล้ว') stat.feeCollected += fee; // เก็บได้จริง — เฉพาะที่ชำระแล้วเท่านั้น (ไม่รวมฟรี)
      } else if (status === 'ปฏิเสธ') {
        stat.rejected++;
      } else {
        stat.pending++;
      }

      slotCols.forEach(function(col, idx) {
        var fullName = '', pre = '', ln = '', idc = '', lic = '', al = '';
        if (col.fn >= 0) {
          pre = col.pre >= 0 ? (r[col.pre] || '') : '';
          ln  = col.ln  >= 0 ? (r[col.ln]  || '') : '';
          fullName = String(r[col.fn] || '').trim();
        } else if (col.combined >= 0) {
          fullName = String(r[col.combined] || '').trim();
        } else {
          return;
        }
        if (!fullName) return;
        idc = col.idc >= 0 ? (r[col.idc] || '') : '';
        lic = col.lic >= 0 ? (r[col.lic] || '') : '';
        al  = col.al  >= 0 ? (r[col.al]  || '') : '';
        stat.named++;
        people.push([cfg.nm, 'ผู้เข้าร่วมทั่วไป', teamName, regId, idx + 1, pre, fullName, ln, idc, lic, al, status]);
      });
    });
  });

  // ── 2) กรรมการอำนวยการ + กรรมการตัดสินการแข่งขัน (จาก REGISTRATIONS) ──
  var reg = ss.getSheetByName(SH_REG);
  if (reg && reg.getLastRow() > 1) {
    var rdata = reg.getDataRange().getValues();
    var rhdr  = rdata[0];
    var iType  = rhdr.indexOf('ประเภทผู้ลงทะเบียน');
    var iPre   = rhdr.indexOf('คำนำหน้า');
    var iFn    = rhdr.indexOf('ชื่อ');
    var iLn    = rhdr.indexOf('นามสกุล');
    var iOrg2  = rhdr.indexOf('สังกัด/หน่วยงาน');
    var iCard  = rhdr.indexOf('เลขบัตรประชาชน');
    var iAl2   = rhdr.indexOf('การแพ้อาหาร');
    var iComp2 = rhdr.indexOf('ตัดสินประเภท');
    // อ่านจากคอลัมน์ "สถานะอัพเดทอัตโนมัติตาม PAYMENTS" ก่อน ถ้าไม่เจอ fallback ไปคอลัมน์ "สถานะชำระ" แล้วค่อย hardcode col W (22)
   var iStat2 = -1;
for (var _h2 = 0; _h2 < rhdr.length; _h2++) { if (String(rhdr[_h2]).indexOf('สถานะอัพเดท') >= 0) { iStat2 = _h2; break; } }
    if (iStat2 < 0) iStat2 = rhdr.indexOf('สถานะชำระ');
    if (iStat2 < 0) iStat2 = 22;

    for (var i = 1; i < rdata.length; i++) {
      var type = String(rdata[i][iType] || '').trim();
      if (type !== TYPE_LBL.committee_dir && type !== TYPE_LBL.committee_judge) continue;
      var roomName = type; // ใช้ชื่อประเภทเป็นชื่อ "ห้อง"
      var stat = ensureRoom(roomName, 'กรรมการ');
      stat.teams++; stat.quota += 1; stat.named++;
      var status = String(rdata[i][iStat2] || '');
      // F8: กรรมการไม่มีค่าธรรมเนียม — _registerUser เขียน 'รอชำระเงิน' ให้ ต้องนับเป็นฟรี/ไม่ค้าง
      if (status.indexOf('ปฏิเสธ') >= 0) stat.rejected++;
      else { stat.paid++; if (status.indexOf('ชำระแล้ว') < 0) status = 'ฟรี'; }
      // กรรมการไม่มีค่าธรรมเนียม ไม่ต้องบวกเข้า feeCollected/feeTotal

      people.push([
        roomName, roomName, rdata[i][iOrg2] || '', rdata[i][0], 1,
        rdata[i][iPre] || '', rdata[i][iFn] || '', rdata[i][iLn] || '',
        rdata[i][iCard] || '', (type === TYPE_LBL.committee_judge ? (rdata[i][iComp2] || '') : ''),
        rdata[i][iAl2] || '', status
      ]);
    }
  }

  // ── เขียนชีท "สรุปการลงทะเบียนทั้งหมด" ──────────────────────
  var sh = ss.getSheetByName(SH_FULLSUMMARY);
  if (!sh) { sh = ss.insertSheet(SH_FULLSUMMARY); } else { sh.clearContents(); sh.clearFormats(); }

  var peopleHdr = ['ห้อง/รายการ','ประเภท','ชื่อทีม/สังกัด','REG_ID','ลำดับ','คำนำหน้า','ชื่อ','นามสกุล','เลขบัตร/ตัดสินประเภท','ใบประกอบ','การแพ้อาหาร','สถานะชำระ'];
  sh.appendRow(peopleHdr);
  sh.getRange(1,1,1,peopleHdr.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);
  if (people.length) sh.getRange(2,1,people.length,peopleHdr.length).setValues(people);
  sh.autoResizeColumns(1, peopleHdr.length);

  // ── ตารางสรุปต่อห้อง (ต่อท้าย เว้น 2 แถว) ────────────────────
  var sumStartRow = people.length + 3;
  sh.getRange(sumStartRow, 1).setValue('สรุปจำนวนการลงทะเบียนตามห้อง/รายการ')
    .setFontWeight('bold').setFontSize(12);

  var sumHeaderRow = sumStartRow + 1;
  var sumHdr = ['ห้อง/รายการ','ประเภท','จำนวนทีม/คนลงทะเบียน','โควต้ารวม (คน)','กรอกชื่อจริง (คน)','ยอดเงินเก็บได้จริง (บาท)','ยอดเงินรวมทั้งหมด (บาท)','ชำระแล้ว','รอดำเนินการ','ปฏิเสธ'];
  sh.getRange(sumHeaderRow, 1, 1, sumHdr.length).setValues([sumHdr])
    .setBackground('#162544').setFontColor('#fff').setFontWeight('bold');

  var roomKeys = Object.keys(roomStats).sort();
  var sumRows = roomKeys.map(function(k) {
    var s = roomStats[k];
    return [k, s.type, s.teams, s.quota, s.named, s.feeCollected, s.feeTotal, s.paid, s.pending, s.rejected];
  });
  if (sumRows.length) sh.getRange(sumHeaderRow + 1, 1, sumRows.length, sumHdr.length).setValues(sumRows);

  var grandTotalRow = sumHeaderRow + 1 + sumRows.length;
  var gTeams         = sumRows.reduce(function(s,r){return s+r[2];},0);
  var gQuota         = sumRows.reduce(function(s,r){return s+r[3];},0);
  var gNamed         = sumRows.reduce(function(s,r){return s+r[4];},0);
  var gFeeCollected  = sumRows.reduce(function(s,r){return s+r[5];},0);
  var gFeeTotal      = sumRows.reduce(function(s,r){return s+r[6];},0);
  var gPaid          = sumRows.reduce(function(s,r){return s+r[7];},0);
  var gPend          = sumRows.reduce(function(s,r){return s+r[8];},0);
  var gRej           = sumRows.reduce(function(s,r){return s+r[9];},0);
  sh.getRange(grandTotalRow, 1, 1, sumHdr.length).setValues([['รวมทั้งหมด','', gTeams, gQuota, gNamed, gFeeCollected, gFeeTotal, gPaid, gPend, gRej]])
    .setBackground('#f0fdf4').setFontWeight('bold');

  sh.autoResizeColumns(1, sumHdr.length);

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      'อัปเดตชีท "' + SH_FULLSUMMARY + '" เสร็จแล้ว',
      'รวมทุกคน: ' + people.length + ' รายการ\n' +
      'จำนวนห้อง/รายการทั้งหมด: ' + roomKeys.length + ' รายการ\n' +
      'เก็บเงินได้จริง: ' + gFeeCollected.toLocaleString() + ' บาท\n' +
      'ยอดค้างรวมทั้งหมด: ' + gFeeTotal.toLocaleString() + ' บาท\n\n' +
      'ดูตารางสรุปต่อห้องได้ที่ท้ายชีท (แถวที่ ' + sumStartRow + ' เป็นต้นไป)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ================================================================
//  EMAIL QUEUE: ตอบผู้ใช้ทันที แล้วส่งอีเมลเบื้องหลัง (~1 นาที)
//  ชีทซ่อน EMAIL_QUEUE: id, created, kind, regId, payloadJson, status, attempts, lastError, sentAt
//  kind = register | approve | reject · status = pending | sent | error
// ================================================================
var SH_MAILQ = 'EMAIL_QUEUE';
var MAILQ_HDR = ['id','created','kind','regId','payloadJson','status','attempts','lastError','sentAt'];
var MAILQ_MAX_ATTEMPTS = 3;

function _queueEmail(kind, regId, payloadObj) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH_MAILQ);
  if (!sh) {
    sh = _sheet(ss, SH_MAILQ, MAILQ_HDR);
    try { sh.hideSheet(); } catch (eH) {}
  }
  sh.appendRow([Utilities.getUuid(), new Date(), kind, String(regId), JSON.stringify(payloadObj || {}), 'pending', 0, '', '']);
  try { _ensureMailTrigger(1000); } catch (e) { Logger.log('ensureMailTrigger err: ' + e.message + ' (ใช้เมนู ✉️ ส่งอีเมลค้างในคิวตอนนี้)'); }
}

// สร้าง trigger ครั้งเดียวรัน _processEmailQueue — มี flag กันสร้างซ้ำ (แบบเดียวกับ _scheduleAutoExports)
function _ensureMailTrigger(delayMs) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) { return; }
  try {
    var props = PropertiesService.getScriptProperties();
    var t0 = Number(props.getProperty('mailTriggerPending') || 0);
    if (t0 > 1 && Date.now() - t0 < 5 * 60 * 1000) return; // มีคิวรออยู่แล้ว (ใหม่กว่า 5 นาที)
    var trg = ScriptApp.newTrigger('_processEmailQueue').timeBased().after(delayMs || 1000).create();
    props.setProperty('mailTriggerUid', trg.getUniqueId());
    props.setProperty('mailTriggerPending', String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

// สร้างอีเมลจาก payload ที่เก็บไว้ แล้วส่งเหมือนเดิมทุกประการ (throw ถ้าล้มเหลว)
function _sendQueuedMail(kind, regId, p) {
  if (kind === 'register') {
    var d = p.d || {};
    var meal = _calcMeal(d, !!p.isFree, p.fullName, p.org, p.typeLbl, regId);
    _sendConfirmEmail(p.fullName, p.org, p.typeLbl, p.typeLblEn, p.email, regId, p.fee, !!p.isFree, d, meal);
  } else if (kind === 'approve') {
    _sendPaymentConfirmEmail(p.name, p.org, p.typeLbl, p.email, regId, p.fee, p.d17, p.d18, p.approvedAt);
  } else if (kind === 'reject') {
    var nm = p.name || '';
    GmailApp.sendEmail(p.email,
      '[Nursing & Innovation 2027] กรุณาส่งสลิปใหม่ — ' + regId,
      'เรียน ' + nm + '\nทีมงานไม่สามารถยืนยันสลิปได้ กรุณาส่งสลิปใหม่\nREG ID: ' + regId + '\n\nDear ' + nm + ',\nWe could not verify your payment. Please resubmit.\nREG ID: ' + regId,
      { name: CONF_NAME_EN });
  } else {
    throw new Error('unknown mail kind: ' + kind);
  }
}

// ฟังก์ชันที่ trigger เรียก (และเมนู) — คืน {sent, failed, pending, stopped}
function _processEmailQueue() {
  var props = PropertiesService.getScriptProperties();
  var t0 = Date.now(), out = { sent: 0, failed: 0, pending: 0, stopped: '' };
  // ล้าง flag ตั้งแต่ต้น — อีเมลที่เข้าคิวระหว่างรันจะจัดรอบใหม่ได้; และกันรันซ้อนด้วย mutex ใน Properties
  // (ไม่ถือ ScriptLock ตอนส่ง เพราะส่งหลายฉบับนานหลายสิบวินาที จะบล็อกการลงทะเบียน/อนุมัติ)
  var lk = LockService.getScriptLock(), busy = false;
  try {
    lk.waitLock(10000);
    props.deleteProperty('mailTriggerPending');
    var run0 = Number(props.getProperty('mailProcessing') || 0);
    if (run0 && Date.now() - run0 < 6 * 60 * 1000) busy = true; else props.setProperty('mailProcessing', String(Date.now()));
  } catch (eL) { out.stopped = 'lock'; return out; } finally { try { lk.releaseLock(); } catch (eR) {} }
  if (busy) { out.stopped = 'busy'; return out; }
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SH_MAILQ);
    if (!sh || sh.getLastRow() < 2) return out;
    var v = sh.getRange(1, 1, sh.getLastRow(), MAILQ_HDR.length).getValues();
    var H = v[0], iK = H.indexOf('kind'), iR = H.indexOf('regId'), iP = H.indexOf('payloadJson'), iS = H.indexOf('status'), iA = H.indexOf('attempts');
    var quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (eQ) {}
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][iS]) !== 'pending') continue;
      if (Date.now() - t0 > 270000) { out.stopped = 'time'; break; }            // ~4.5 นาที
      if (quota !== null && quota <= 0) { out.stopped = 'quota'; Logger.log('EMAIL QUEUE: โควตาอีเมลหมด — ค้างในคิว'); break; }
      var attempts = Number(v[i][iA] || 0), row = i + 1, status = 'sent', err = '';
      try {
        _sendQueuedMail(String(v[i][iK]), String(v[i][iR]), JSON.parse(String(v[i][iP] || '{}')));
        if (quota !== null) quota--;
        attempts++;
      } catch (me) {
        attempts++; err = me.message;
        status = attempts >= MAILQ_MAX_ATTEMPTS ? 'error' : 'pending';
        Logger.log('EMAIL QUEUE ERROR [' + v[i][iR] + '] attempt ' + attempts + ': ' + me.message);
      }
      sh.getRange(row, iS + 1, 1, 4).setValues([[status, attempts, err, err ? '' : new Date()]]);
      v[i][iS] = status;
      if (status === 'sent') out.sent++; else if (status === 'error') out.failed++;
    }
    for (var j = 1; j < v.length; j++) if (String(v[j][iS]) === 'pending') out.pending++;
  } finally {
    try { lk.waitLock(5000); props.deleteProperty('mailProcessing'); } catch (e1) {} finally { try { lk.releaseLock(); } catch (e2) {} }
    // ลบ trigger ที่ใช้แล้วไม่ให้สะสม (เก็บอันที่เพิ่งจัดคิวใหม่ระหว่างรัน)
    try {
      var keepUid = props.getProperty('mailTriggerPending') ? (props.getProperty('mailTriggerUid') || '') : '';
      ScriptApp.getProjectTriggers().forEach(function (t) {
        if (t.getHandlerFunction() === '_processEmailQueue' && t.getUniqueId() !== keepUid) ScriptApp.deleteTrigger(t);
      });
    } catch (e3) {}
  }
  // ยังมีงานค้างที่ลองซ้ำได้ (ไม่ใช่โควตาหมด) → จัดรอบใหม่ใน 1 นาที
  if (out.pending > 0 && out.stopped !== 'quota') { try { _ensureMailTrigger(60000); } catch (e4) { Logger.log('requeue mail trigger err: ' + e4.message); } }
  return out;
}

// เมนู: ส่งอีเมลค้างในคิวตอนนี้
function sendPendingEmailsNow() {
  var ui = SpreadsheetApp.getUi();
  var r = _processEmailQueue();
  ui.alert('✉️ ส่งอีเมลค้างในคิว\n\nส่งสำเร็จ: ' + r.sent + '\nล้มเหลวถาวร: ' + r.failed + '\nยังค้าง: ' + r.pending +
    (r.stopped === 'quota' ? '\n\n⚠️ โควตาอีเมลวันนี้หมด — ลองใหม่พรุ่งนี้' : r.stopped === 'time' ? '\n\nหมดเวลารอบนี้ กดอีกครั้งเพื่อส่งที่เหลือ' : r.stopped === 'busy' ? '\n\nมีรอบส่งอื่นกำลังทำงานอยู่ ลองใหม่ภายหลัง' : '') +
    '\n\nรายละเอียดดูชีทซ่อน EMAIL_QUEUE');
}

  // ================================================================
//  ASYNC AUTO-EXPORT: ให้ export 3 ชีท (กรรมการ/อาหาร/สรุปทั้งหมด)
//  รันแยกเบื้องหลัง ไม่บล็อกการตอบกลับผู้ใช้ตอนลงทะเบียน
// ================================================================

// เรียกจาก _registerUser — สร้าง trigger แบบครั้งเดียว รันภายใน ~20 วิ
// ถ้ามีคิวรออยู่แล้ว (เผื่อมีคนลงทะเบียนพร้อมกันหลายคน) จะไม่สร้างคิวซ้ำ
function _scheduleAutoExports() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) { return; } // ไม่ได้ล็อก = มีอีกคำขอกำลังจัดคิวอยู่
  try {
    var props = PropertiesService.getScriptProperties();
    var t0 = Number(props.getProperty('autoExportPending') || 0);
    // มีคิวรอและยังใหม่ (< 3 นาที) → ข้าม; เก่ากว่านั้นถือว่าค้าง → จัดคิวใหม่
    if (t0 > 1 && Date.now() - t0 < 3 * 60 * 1000) return;
    if (t0 === 1) { /* ค่าเก่า '1' จากเวอร์ชันก่อน ถือว่าค้าง → จัดคิวใหม่ */ }
    // F6: สร้าง trigger ก่อน แล้วค่อยตั้ง flag (ถ้า create ล้มเหลว flag จะไม่ค้าง)
    var trg = ScriptApp.newTrigger('_runAutoExportsNow')
      .timeBased()
      .after(20 * 1000) // รอ 20 วินาที ให้ request ลงทะเบียนตอบกลับผู้ใช้เสร็จก่อน
      .create();
    props.setProperty('autoExportTrigUid', trg.getUniqueId());
    props.setProperty('autoExportPending', String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

// ฟังก์ชันที่ trigger เรียกจริง — รัน export 3 ตัว แล้วลบ trigger ตัวเองทิ้งกันสะสม
function _runAutoExportsNow() {
  var props = PropertiesService.getScriptProperties();
  // F6: ล้าง flag ตั้งแต่ต้น — คนลงทะเบียนระหว่างที่ export กำลังรันจะจัดคิวรอบใหม่ได้
  var lk = LockService.getScriptLock();
  try { lk.waitLock(5000); props.deleteProperty('autoExportPending'); } catch (eL) {} finally { try { lk.releaseLock(); } catch (eR) {} }
  try {
    try{ exportCommitteeSheets(true); }catch(e1){ Logger.log('auto exportCommitteeSheets err: '+e1.message); }
    try{ exportMealListByCompetition(true); }catch(e2){ Logger.log('auto exportMealListByCompetition err: '+e2.message); }
    try{ exportFullRegistrationSummary(true); }catch(e3){ Logger.log('auto exportFullRegistrationSummary err: '+e3.message); }
    try{ exportItemFoodSummary(true); }catch(e4){ Logger.log('auto exportItemFoodSummary err: '+e4.message); }
  } finally {
    // ลบ trigger ที่ใช้แล้วตาม handler name ไม่ให้สะสม — แต่เก็บอันที่เพิ่งจัดคิวใหม่ระหว่างรัน (ถ้ามี flag ค้างอยู่)
    var keepUid = props.getProperty('autoExportPending') ? (props.getProperty('autoExportTrigUid') || '') : '';
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t){
      if (t.getHandlerFunction() === '_runAutoExportsNow' && t.getUniqueId() !== keepUid) {
        ScriptApp.deleteTrigger(t);
      }
    });
  }
}

// เผื่อ trigger ค้าง (เช่น สคริปต์ error กลางทางไม่ทันลบ) — รันมือได้เพื่อล้างคิวและ trigger ค้างทั้งหมด
function resetAutoExportQueue() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('autoExportPending');
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t){
    if (t.getHandlerFunction() === '_runAutoExportsNow') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  SpreadsheetApp.getUi().alert('ล้างคิวเรียบร้อย — ลบ trigger ค้าง ' + removed + ' ตัว');
}

// ── อนุมัติการชำระเงินย้อนหลัง เฉพาะ REG_ID ที่ระบุ (ใส่ได้หลายอันคั่นด้วยจุลภาค) ──
// เช็คสถานะปัจจุบันให้ดูก่อน แล้วถามยืนยันก่อนอนุมัติจริง
function approveSpecificRegIds() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'อนุมัติย้อนหลังเฉพาะ REG_ID',
    'ใส่ REG_ID (คั่นด้วยจุลภาคถ้ามีหลายอัน)\nเช่น INHI26-0613-0014, INHI26-0613-0015',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var input = resp.getResponseText().trim();
  if (!input) { ui.alert('ไม่ได้ใส่ REG_ID'); return; }
  var regIds = input.split(',').map(function(s){ return s.trim(); }).filter(Boolean);

  var ss = SpreadsheetApp.openById(SS_ID);
  var psh = ss.getSheetByName(SH_PAY);
  var rsh = ss.getSheetByName(SH_REG);

  // ── ขั้น 1: แสดงสถานะปัจจุบันของแต่ละ REG_ID ก่อนอนุมัติ ──
  var preview = [];
  var foundInReg = {};
  regIds.forEach(function(rid) {
    var regStatus = 'ไม่พบใน REGISTRATIONS';
    var payStatus = 'ไม่มีสลิปเลย (ไม่มีแถวใน PAYMENTS)';
    var name = '';

    if (rsh && rsh.getLastRow() > 1) {
      var rd = rsh.getDataRange().getValues();
      var rh = rd[0];
      var iStat = rh.indexOf('สถานะชำระ');
      var iName = rh.indexOf('ชื่อเต็ม');
      for (var i = 1; i < rd.length; i++) {
        if (String(rd[i][0]).trim() === rid) {
          foundInReg[rid] = true;
          regStatus = String(rd[i][iStat] || '');
          name = String(rd[i][iName] || '');
          break;
        }
      }
    }
    if (psh && psh.getLastRow() > 1) {
      var pd = psh.getDataRange().getValues();
      for (var j = pd.length - 1; j >= 1; j--) { // เอาแถวล่าสุดของ REG_ID นี้
        if (String(pd[j][0]).trim() === rid) {
          payStatus = String(pd[j][6] || '');
          break;
        }
      }
    }
    preview.push(rid + ' (' + name + ') → REGISTRATIONS: [' + regStatus + '] | PAYMENTS ล่าสุด: [' + payStatus + ']');
  });

  var confirmMsg = preview.join('\n') + '\n\nยืนยันอนุมัติทั้งหมดนี้เป็น "ชำระแล้ว" เลยไหม?\n(จะส่งอีเมลยืนยันการชำระเงินให้แต่ละคนด้วย)';
  var confirm = ui.alert('ตรวจสอบก่อนอนุมัติ', confirmMsg, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  // ── ขั้น 2: อนุมัติจริง ──
  var results = [];
  regIds.forEach(function(rid) {
    try {
      // F7: ข้าม REG_ID ที่ไม่มีใน REGISTRATIONS (พิมพ์ผิด) ไม่ให้รายงานว่า "บังคับสำเร็จ" ทั้งที่ไม่ได้เขียนอะไร
      if (!foundInReg[rid]) {
        Logger.log('approveSpecificRegIds SKIP (ไม่พบใน REGISTRATIONS): ' + rid);
        results.push(rid + ': ❌ ไม่พบใน REGISTRATIONS — ข้าม');
        return;
      }
      // ถ้ามีแถวใน PAYMENTS ที่ยัง "รอตรวจสอบ" อยู่ ให้ใช้ flow ปกติ (_approvePayment จะจัดการ PAYMENTS+REGISTRATIONS+อีเมลให้ครบ)
      var hasPendingSlip = false;
      if (psh && psh.getLastRow() > 1) {
        var pd2 = psh.getDataRange().getValues();
        for (var k = 1; k < pd2.length; k++) {
          if (String(pd2[k][0]).trim() === rid && String(pd2[k][6]) === 'รอตรวจสอบ') { hasPendingSlip = true; break; }
        }
      }

      if (hasPendingSlip) {
        var r = _approvePayment(rid, 'staff-retro-approve');
        results.push(rid + ': อนุมัติผ่าน flow ปกติ ' + (r.success ? '✅' : '❌ ' + r.message) + (r.emailSent ? ' (ส่งอีเมลแล้ว)' : ' (ไม่ได้ส่งอีเมล)'));
      } else {
        // ไม่มีสลิปเลย (เช่น จ่ายเงินสดหน้างาน/โอนแล้วแต่ไม่ได้แนบสลิปในระบบ) → บังคับสถานะตรงๆ ไม่ผ่านสลิป
        _updatePayStatus(ss, rid, 'ชำระแล้ว');
        results.push(rid + ': ไม่มีสลิปในระบบ → บังคับสถานะเป็น "ชำระแล้ว" ตรงๆ (ไม่ได้ส่งอีเมล)');
      }
    } catch (e) {
      results.push(rid + ': ❌ error — ' + e.message);
    }
  });

SpreadsheetApp.flush();

  // ── Auto-export: อัปเดต 4 ชีทสรุป/รายชื่ออาหาร ให้ตรงหลังอนุมัติย้อนหลัง ──
  try{ _scheduleAutoExports(); }catch(autoErr){ Logger.log('schedule auto export err: '+autoErr.message); }

  ui.alert('ผลการอนุมัติย้อนหลัง', results.join('\n') + '\n\n(กำลังอัปเดตชีทสรุปทั้ง 4 ให้อัตโนมัติ จะเสร็จภายใน ~20 วิ)', ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: ชีทสรุปจำนวนรายการ + อาหารแต่ละวัน (นับตามโควต้าที่พึงได้รับ)
//  วิธีติดตั้ง: copy ฟังก์ชันนี้ทั้งหมด ไปวางต่อท้ายไฟล์ .gs หลักของโปรเจกต์
//              (วางใกล้ๆ exportFullRegistrationSummary / exportMealListByCompetition ก็ได้)
//  วิธีรัน: เปิด Apps Script Editor → เลือกฟังก์ชัน exportItemFoodSummary → กด Run
//           หรือรันจากเมนู custom เดิมถ้ามีอยู่แล้ว
//  ไม่แก้ไข/ไม่ลบข้อมูลในชีทอื่นใดๆ ทั้งสิ้น สร้าง/อัปเดตเฉพาะชีท SH_ITEMFOOD เท่านั้น
// ================================================================

var SH_ITEMFOOD = 'สรุปจำนวนรายการและอาหาร';

// silent=true → ไม่แสดง popup (ใช้เผื่อเรียกอัตโนมัติจาก _runAutoExportsNow ในอนาคต)
function exportItemFoodSummary(silent) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stats = {}; // { ชื่อรายการ/ประเภท: {type, teams, people, food17, food18} }

  function ensure(name, type) {
    if (!stats[name]) stats[name] = { type: type, teams: 0, people: 0, food17: 0, food18: 0 };
    return stats[name];
  }

  // ── 1) รายการแข่งขันทุกอัน (รวม c00 online / c00s onsite) ────────────────
  Object.keys(COMP_SHEET).forEach(function(cid) {
    var cfg = COMP[cid];
    if (!cfg) return;
    var shName = COMP_SHEET[cid];
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    // header อาจอยู่แถว 1 หรือ 2 แล้วแต่ชีท — สแกนหาแถวที่มี 'REG_ID'
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0 || sh.getLastRow() <= headerRowIdx) return;

    var iRegId = hdr.indexOf('REG_ID');
    var iDays  = hdr.indexOf('จำนวนวัน');
    var quota  = cfg.team || 1; // จำนวนคนต่อทีมตามที่กำหนดไว้ใน COMP config
    var stat   = ensure(cfg.nm, 'รายการแข่งขัน');

    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();

    data.forEach(function(r) {
      var regId = String(r[iRegId] || '').trim();
      if (!regId) return; // แถวว่าง ข้าม

      // c00 (ออนไลน์) ไม่มีคอลัมน์ "จำนวนวัน" — แต่ mealDay='none' อยู่แล้วจึงไม่กระทบ
      var days = (iDays >= 0) ? (Number(r[iDays]) || 1) : 1;

      // สูตรเดียวกับตอนคำนวณสิทธิ์อาหารจริงตอนลงทะเบียน (_calcMeal)
      var getsD17 = cfg.mealDay === 'both' || cfg.mealDay === 'd17' || (cfg.mealDay === 'd18' && days === 2);
      var getsD18 = cfg.mealDay === 'both' || cfg.mealDay === 'd18' || (cfg.mealDay === 'd17' && days === 2);

      stat.teams++;
      stat.people += quota;
      if (cfg.mealDay !== 'none') {
        if (getsD17) stat.food17 += quota;
        if (getsD18) stat.food18 += quota;
      }
    });
  });

  // ── 2) กรรมการอำนวยการ + กรรมการตัดสินการแข่งขัน (ได้อาหารทั้ง 2 วันเสมอ) ──
  var reg = ss.getSheetByName(SH_REG);
  if (reg && reg.getLastRow() > 1) {
    var rdata = reg.getDataRange().getValues();
    var rhdr  = rdata[0];
    var iType = rhdr.indexOf('ประเภทผู้ลงทะเบียน');
    for (var i = 1; i < rdata.length; i++) {
      var type = String(rdata[i][iType] || '').trim();
      if (type !== TYPE_LBL.committee_dir && type !== TYPE_LBL.committee_judge) continue;
      var stat = ensure(type, 'กรรมการ');
      stat.teams++;   // นับเป็น 1 หน่วย/คน
      stat.people++;
      stat.food17++;
      stat.food18++;
    }
  }

  // ── เขียนชีทสรุป (ตารางล้วนๆ ไม่มีรายชื่อรายบุคคล) ──────────────────────
  var sh = ss.getSheetByName(SH_ITEMFOOD);
  if (!sh) { sh = ss.insertSheet(SH_ITEMFOOD); } else { sh.clearContents(); sh.clearFormats(); }

  var hdrRow = ['รายการ/ประเภท', 'ประเภท', 'จำนวนทีม/หน่วย', 'จำนวนคนรวม (ตามโควต้า)', 'อาหาร 27 พ.ค. (คน)', 'อาหาร 28 พ.ค. (คน)'];
  sh.appendRow(hdrRow);
  sh.getRange(1, 1, 1, hdrRow.length)
    .setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);

  var keys = Object.keys(stats).sort();
  var rows = keys.map(function (k) {
    var s = stats[k];
    return [k, s.type, s.teams, s.people, s.food17, s.food18];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, hdrRow.length).setValues(rows);

  var totalRow = rows.length + 2;
  var gTeams  = rows.reduce(function (s, r) { return s + r[2]; }, 0);
  var gPeople = rows.reduce(function (s, r) { return s + r[3]; }, 0);
  var g17     = rows.reduce(function (s, r) { return s + r[4]; }, 0);
  var g18     = rows.reduce(function (s, r) { return s + r[5]; }, 0);
  sh.getRange(totalRow, 1, 1, hdrRow.length)
    .setValues([['รวมทั้งหมด', '', gTeams, gPeople, g17, g18]])
    .setBackground('#f0fdf4').setFontWeight('bold');

  sh.autoResizeColumns(1, hdrRow.length);

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      'อัปเดตชีท "' + SH_ITEMFOOD + '" เสร็จแล้ว',
      'จำนวนรายการ/ประเภททั้งหมด: ' + keys.length + '\n' +
      'จำนวนคนรวม (ตามโควต้า): ' + gPeople + '\n' +
      'อาหาร 27 พ.ค.: ' + g17 + ' คน\n' +
      'อาหาร 28 พ.ค.: ' + g18 + ' คน',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ================================================================
//  ADD-ON: หาทีมที่กรอกสมาชิกเกินโควต้า (สาเหตุที่ตัวเลข "กรอกชื่อจริง"
//          กับ "อาหารตามชื่อจริง" มากกว่าตัวเลขตามโควต้า)
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน findOverQuotaTeams แล้วกด Run
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่แสดงผลเป็น popup + Logger
// ================================================================

function findOverQuotaTeams() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var overQuota = []; // [รายการ, ชื่อทีม, REG_ID, โควต้า, กรอกแล้ว, เกินมา]

  Object.keys(COMP_SHEET).forEach(function(cid) {
    var cfg = COMP[cid];
    if (!cfg || cid === 'c00') return; // c00 ออนไลน์ไม่มีทีม ข้าม

    var shName = COMP_SHEET[cid];
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0 || sh.getLastRow() <= headerRowIdx) return;

    var iRegId = hdr.indexOf('REG_ID');
    var iTeam  = hdr.indexOf('ชื่อทีม');
    var quota  = cfg.team || 1;

    // รองรับ header ได้ทั้ง 2 รูปแบบ เหมือนฟังก์ชันอื่นๆ
    var slotCols = [];
    for (var n = 1; n <= 4; n++) {
      slotCols.push({
        fn: hdr.indexOf('ชื่อ ' + n),
        combined: hdr.indexOf('ชื่อ สมาชิก ' + n)
      });
    }

    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();

    data.forEach(function(r) {
      var regId = String(r[iRegId] || '').trim();
      if (!regId) return;
      var teamName = iTeam >= 0 ? (r[iTeam] || '') : '';

      var filled = 0;
      slotCols.forEach(function(col) {
        var fullName = '';
        if (col.fn >= 0) fullName = String(r[col.fn] || '').trim();
        else if (col.combined >= 0) fullName = String(r[col.combined] || '').trim();
        if (fullName) filled++;
      });

      if (filled > quota) {
        overQuota.push([cfg.nm, teamName || '(ไม่ระบุชื่อทีม)', regId, quota, filled, filled - quota]);
      }
    });
  });

  if (overQuota.length === 0) {
    SpreadsheetApp.getUi().alert('✅ ไม่พบทีมที่กรอกเกินโควต้า', 'ทุกทีมกรอกสมาชิกไม่เกินโควต้าที่จ่ายเงินไว้', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var msg = overQuota.map(function(r) {
    return r[0] + ' | ทีม: ' + r[1] + ' | REG_ID: ' + r[2] +
           ' | โควต้า: ' + r[3] + ' คน | กรอกจริง: ' + r[4] + ' คน | เกินมา: ' + r[5] + ' คน';
  }).join('\n');

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(
    '⚠️ พบ ' + overQuota.length + ' ทีมที่กรอกเกินโควต้า',
    msg + '\n\nไปที่ชีทรายการนั้นๆ แล้วลบชื่อสมาชิกคนที่เกินออก (หรือถ้าต้องการให้ครบจริง ต้องคิดค่าธรรมเนียม/อาหารเพิ่ม)\n\nหลังแก้แล้ว รัน exportFullRegistrationSummary, exportMealListByCompetition, exportItemFoodSummary ใหม่อีกครั้ง',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ================================================================
//  ADD-ON: ค้นหา + ลบแถวตาม REG_ID จากทุกชีทต้นทาง (เพื่อให้ตัวเลขสรุปตรง)
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน deleteRowsByRegId แล้วกด Run
//
//  ขั้นตอนการทำงาน:
//   1. ถาม REG_ID ที่ต้องการลบ (ใส่ได้หลายอันคั่นด้วยจุลภาค)
//   2. ค้นหาทุกชีทต้นทางก่อน (ไม่ลบทันที) แล้วแสดง preview ให้ดู
//   3. ถามยืนยันอีกรอบก่อนลบจริง
//   4. ลบจริง แล้วรัน export ชีทสรุปทั้ง 4 ตัวใหม่ให้อัตโนมัติ
//
//  ⚠️ การลบย้อนกลับไม่ได้ — ควรตรวจ preview ให้ดีก่อนกดยืนยัน
//  ไม่แตะชีทสรุป/export โดยตรง (จะถูก regenerate ใหม่หลังลบเสร็จ)
// ================================================================

function deleteRowsByRegId() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'ค้นหา + ลบข้อมูลตาม REG_ID',
    'ใส่ REG_ID ที่ต้องการลบ (คั่นด้วยจุลภาคถ้ามีหลายอัน)\nเช่น INHI26-0613-0013, INHI26-0710-0211',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var input = resp.getResponseText().trim();
  if (!input) { ui.alert('ไม่ได้ใส่ REG_ID'); return; }
  var targetIds = input.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  if (!targetIds.length) { ui.alert('ไม่ได้ใส่ REG_ID'); return; }

  var ss = SpreadsheetApp.openById(SS_ID);

  // ชีทต้นทางทั้งหมดที่อาจมี REG_ID (ไม่รวมชีทสรุป/export เพราะจะ regenerate ใหม่ทีหลัง)
  var sourceSheets = [SH_REG, SH_COMP, SH_PAY, SH_MEAL, SH_CHECKIN]
    .concat(Object.keys(COMP_SHEET).map(function(cid){ return COMP_SHEET[cid]; }));

  // ── ขั้นที่ 1: ค้นหาก่อน ไม่ลบจริง ──────────────────────────
  var found = []; // {sheetName, rowNum, regId, preview}
  sourceSheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 1) return;

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0 || sh.getLastRow() <= headerRowIdx) return;

    var iRegId = hdr.indexOf('REG_ID');
    var iName = hdr.indexOf('ชื่อเต็ม');
    if (iName < 0) iName = hdr.indexOf('ชื่อทีม');
    if (iName < 0) iName = hdr.indexOf('ชื่อ-นามสกุล');

    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
    data.forEach(function(r, idx) {
      var rid = String(r[iRegId] || '').trim();
      if (targetIds.indexOf(rid) >= 0) {
        found.push({
          sheetName: name,
          rowNum: headerRowIdx + 1 + idx,
          regCol: iRegId + 1,
          regId: rid,
          preview: iName >= 0 ? String(r[iName] || '') : ''
        });
      }
    });
  });

  if (found.length === 0) {
    ui.alert('ไม่พบข้อมูล', 'ไม่พบ REG_ID ที่ระบุในชีทต้นทางใดๆ เลย', ui.ButtonSet.OK);
    return;
  }

  var previewMsg = found.map(function(f) {
    return f.sheetName + ' | แถว ' + f.rowNum + ' | ' + f.regId + (f.preview ? ' | ' + f.preview : '');
  }).join('\n');

  var confirm = ui.alert(
    'ยืนยันการลบ (พบ ' + found.length + ' แถว)',
    previewMsg + '\n\n⚠️ การลบนี้ย้อนกลับไม่ได้ ยืนยันลบทั้งหมดนี้ใช่ไหม?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // ── ขั้นที่ 2: ลบจริง — ลบจากแถวล่างขึ้นบนภายในชีทเดียวกัน กันเลขแถวเลื่อน ──
  // [ADM2-08] ล็อกกัน _registerUser / admin อื่นแทรกระหว่างลบ
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { ui.alert('ระบบกำลังประมวลผลอยู่ ลองใหม่อีกครั้ง'); return; }
  var deletedCount = 0, skipped = [], sheetsTouched = {};
  try {
    // [ADM2-09a] เก็บสำเนาแถวที่ลบไว้ในชีท DELETED_ROWS_LOG (กู้ย้อนหลังได้)
    var logSh = ss.getSheetByName('DELETED_ROWS_LOG') || ss.insertSheet('DELETED_ROWS_LOG');
    var bySheet = {};
    found.forEach(function(f) {
      if (!bySheet[f.sheetName]) bySheet[f.sheetName] = [];
      bySheet[f.sheetName].push(f);
    });

    Object.keys(bySheet).forEach(function(name) {
      var sh = ss.getSheetByName(name);
      var items = bySheet[name].sort(function(a, b) { return b.rowNum - a.rowNum; }); // มากไปน้อย
      items.forEach(function(f) {
        // [ADM2-08] ตรวจซ้ำก่อนลบ: แถวนั้นยังเป็น REG_ID เดิมอยู่ไหม (กันแถวเลื่อน)
        var cur = String(sh.getRange(f.rowNum, f.regCol).getValue() || '').trim();
        if (cur !== f.regId) { skipped.push(name + ' แถว ' + f.rowNum + ' (REG_ID เปลี่ยนเป็น "' + cur + '")'); return; }
        try {
          logSh.appendRow([new Date(), name, f.rowNum, f.regId, JSON.stringify(sh.getRange(f.rowNum, 1, 1, sh.getLastColumn()).getValues()[0])]);
        } catch (logErr) { Logger.log('delete log err: ' + logErr.message); }
        sh.deleteRow(f.rowNum);
        deletedCount++;
        sheetsTouched[name] = true;
      });
    });

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // ── ขั้นที่ 3: ตั้ง export ชีทสรุปทั้งหมดใหม่ผ่าน scheduler (ไม่รันตรงๆ กันชน 6 นาที) ──
  try { _scheduleAutoExports(); } catch (e1) { Logger.log('schedule auto export err: ' + e1.message); }

  ui.alert(
    '✅ ลบเสร็จแล้ว',
    'ลบไปทั้งหมด ' + deletedCount + ' แถว จาก ' + Object.keys(sheetsTouched).length + ' ชีท' +
    (skipped.length ? '\n\n⚠️ ข้าม ' + skipped.length + ' แถวเพราะข้อมูลเปลี่ยนไประหว่างรอ:\n' + skipped.join('\n') : '') +
    '\n\nสำเนาแถวที่ลบเก็บไว้ในชีท DELETED_ROWS_LOG' +
    '\nชีทสรุปทั้งหมดจะอัปเดตอัตโนมัติภายใน ~20 วิ (กรรมการ / อาหาร / สรุปการลงทะเบียน / สรุปจำนวนรายการและอาหาร)',
    ui.ButtonSet.OK
  );
}

// ================================================================
//  ADD-ON: บีบอัดแถวข้อมูลในชีทรายการแข่งขันให้ติดกัน (Compact Rows)
//  แก้ปัญหา: ข้อมูลจริงกระโดดไปอยู่ไกล (เช่น แถว 2000+) เพราะ getLastRow()
//            เคยบวมผิดปกติจากฟอร์แมตที่ตั้งไว้ล่วงหน้าตอน setup
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน compactCompSheets แล้วกด Run
//
//  ⚠️ สำคัญ: ฟังก์ชันนี้ "ย้ายตำแหน่งแถว" เท่านั้น ไม่แก้ไข/ไม่ลบค่าใดๆ ในแถว
//  จะแสดง preview ให้ตรวจสอบก่อน แล้วถามยืนยันก่อนดำเนินการจริงเสมอ
// ================================================================

function compactCompSheets() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();
  var targets = Object.keys(COMP_SHEET).map(function(cid){ return COMP_SHEET[cid]; });
  var preview = [];
  var plans = {};

  targets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { preview.push(name + ' → ไม่มีข้อมูล ข้าม'); return; }

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0) { preview.push(name + ' → หา header (REG_ID) ไม่เจอ ข้าม'); return; }

    var iRegId = hdr.indexOf('REG_ID');
    var lastRow = sh.getLastRow();
    if (lastRow <= headerRowIdx) { preview.push(name + ' → ไม่มีข้อมูล ข้าม'); return; }

    var dataRange = sh.getRange(headerRowIdx + 1, 1, lastRow - headerRowIdx, lastCol);
    var data = dataRange.getValues();
    var dataF = dataRange.getFormulasR1C1(); // [ADM2-10] เก็บสูตรไว้ด้วย
    // [ADM2-10] แถวที่ REG_ID ว่างแต่มีข้อมูลอื่น -> ไม่แตะชีทนี้ (ให้ตรวจมือ)
    var orphan = data.filter(function(r) {
      return String(r[iRegId] || '').trim() === '' && r.some(function(c) { return String(c).trim() !== ''; });
    });
    if (orphan.length) { preview.push(name + ' → ⚠️ มี ' + orphan.length + ' แถวที่มีข้อมูลแต่ REG_ID ว่าง — ข้าม (ตรวจมือก่อน)'); return; }
    var realIdx = [];
    data.forEach(function(r, ri) { if (String(r[iRegId] || '').trim() !== '') realIdx.push(ri); });
    var realRows = realIdx.map(function(ri) { return data[ri]; });
    var realFormulas = realIdx.map(function(ri) { return dataF[ri]; });

    if (realRows.length === data.length) {
      preview.push(name + ' → เรียบร้อยดีอยู่แล้ว ไม่มีช่องว่างคั่นกลาง (' + realRows.length + ' แถว)');
      return;
    }

    plans[name] = { headerRowIdx: headerRowIdx, lastCol: lastCol, realRows: realRows, realFormulas: realFormulas, iRegId: iRegId, oldDataRows: data.length };
    preview.push(name + ' → พบช่องว่างคั่นกลาง: ข้อมูลจริง ' + realRows.length + ' แถว จากทั้งหมด ' + data.length + ' แถว (จะบีบอัดให้ติดกัน)');
  });

  if (Object.keys(plans).length === 0) {
    ui.alert('✅ ไม่พบชีทที่มีช่องว่างคั่นกลาง', preview.join('\n'), ui.ButtonSet.OK);
    return;
  }

  var confirm = ui.alert(
    'ยืนยันการบีบอัดข้อมูล (Compact)',
    preview.join('\n') +
    '\n\n⚠️ จะย้ายแถวข้อมูลจริงทั้งหมดให้มาติดกันใต้ header ทันที ' +
    '(ไม่ลบ/ไม่แก้ไขค่าใดๆ ในแถว แค่ย้ายตำแหน่ง) แล้วลบแถวว่างส่วนเกินท้ายชีททิ้ง\n\nยืนยันดำเนินการ?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // [ADM2-08] ล็อกกัน _registerUser เพิ่มแถวระหว่าง clear -> write
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { ui.alert('ระบบกำลังประมวลผลอยู่ ลองใหม่อีกครั้ง'); return; }
  var results = [];
  try {
    Object.keys(plans).forEach(function(name) {
      var p = plans[name];
      var sh = ss.getSheetByName(name);

      // [ADM2-08] ตรวจซ้ำภายใต้ล็อก: ถ้าจำนวนแถวเปลี่ยนไปจากตอน preview ให้ข้ามชีทนี้
      if (sh.getLastRow() - p.headerRowIdx !== p.oldDataRows) {
        results.push(name + ' → ⚠️ ข้าม: ข้อมูลเปลี่ยนไประหว่างรอยืนยัน (รันใหม่)');
        return;
      }

      // ล้างเฉพาะ "เนื้อหา" ของช่วงข้อมูลเดิมทั้งหมด (ไม่แตะ header, ไม่แตะฟอร์แมต)
      sh.getRange(p.headerRowIdx + 1, 1, p.oldDataRows, p.lastCol).clearContent();

      // เขียนข้อมูลจริงกลับไปให้ติดกันใต้ header ทันที
      if (p.realRows.length) {
        sh.getRange(p.headerRowIdx + 1, 1, p.realRows.length, p.lastCol).setValues(p.realRows);
        // [ADM2-10] คืนสูตร (R1C1 จึงอ้างอิงแถวใหม่ถูกต้อง) เฉพาะเซลล์ที่เดิมเป็นสูตร
        p.realFormulas.forEach(function(frow, ri) {
          frow.forEach(function(f, ci) {
            if (f && String(f).charAt(0) === '=') {
              sh.getRange(p.headerRowIdx + 1 + ri, ci + 1).setFormulaR1C1(f);
            }
          });
        });
      }

      // ลบแถวว่างส่วนเกินท้ายชีท (กันไม่ให้ getLastRow บวมซ้ำอีกในอนาคต)
      var newLastRow = p.headerRowIdx + p.realRows.length;
      var currentLastRow = sh.getLastRow();
      if (currentLastRow > newLastRow) {
        sh.deleteRows(newLastRow + 1, currentLastRow - newLastRow);
      }

      results.push(name + ' → บีบอัดเสร็จ: เหลือ ' + p.realRows.length + ' แถวติดกัน (ลบช่องว่างคั่นกลาง ' + (p.oldDataRows - p.realRows.length) + ' แถว)');
    });
  } finally {
    lock.releaseLock();
  }

  SpreadsheetApp.flush();
  ui.alert('✅ บีบอัดข้อมูลเสร็จสมบูรณ์', results.join('\n'), ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: ตรวจสอบ 2 จุดเสี่ยง — (1) header คอลัมน์ V ตรงเป๊ะไหม
//          (2) REGISTRATIONS มีปัญหา getLastRow() บวมเหมือนชีทรายการแข่งขันไหม
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน diagnoseRegistrationsSheet แล้วกด Run
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่อ่านและแสดงผล
// ================================================================

// [ADM2-15b] แปลง index คอลัมน์ (0-based) เป็นตัวอักษร A, B, ... Z, AA, ...
function _colLetter_(i) {
  var s = ''; i++;
  while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

function diagnoseRegistrationsSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH_REG);
  var ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert('ไม่พบ REGISTRATIONS'); return; }

  var msg = [];

  // ── ตรวจที่ 1: header คอลัมน์ V ตรงเป๊ะกับที่โค้ดค้นหาไหม ──────
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  var target = 'สถานะอัพเดทอัตโนมัติตาม PAYMENTS';
  var exactIdx = hdr.indexOf(target);

  // หาแบบ partial ด้วย เผื่อ exact ไม่เจอ
  var partialIdx = -1;
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]).indexOf('สถานะอัพเดท') >= 0) { partialIdx = i; break; }
  }

  msg.push('=== ตรวจ header คอลัมน์สถานะอัพเดทฯ ===');
  if (exactIdx >= 0) {
    msg.push('✅ หาแบบ exact match เจอที่ col ' + _colLetter_(exactIdx) + ' — โค้ดปัจจุบันทำงานถูกต้อง');
  } else if (partialIdx >= 0) {
    msg.push('❌ หาแบบ exact match ไม่เจอ! แต่หาแบบ partial เจอที่ col ' + _colLetter_(partialIdx));
    msg.push('   → ค่าจริงในเซลล์คือ: [' + JSON.stringify(hdr[partialIdx]) + ']');
    msg.push('   → ความยาว: ' + String(hdr[partialIdx]).length + ' ตัวอักษร');
    msg.push('   → นี่คือบั๊ก! exportFullRegistrationSummary กำลัง fallback ไปใช้คอลัมน์ "สถานะชำระ" (R) แทนแบบเงียบๆ');
  } else {
    msg.push('❌❌ ไม่พบคอลัมน์นี้เลยแม้แต่แบบ partial — ตรวจสอบว่าคอลัมน์ยังอยู่ไหม');
  }

  // ── ตรวจที่ 2: REGISTRATIONS มี getLastRow บวมไหม ──────────────
  msg.push('\n=== ตรวจ getLastRow() ของ REGISTRATIONS ===');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('REGISTRATIONS ยังไม่มีข้อมูล (มีแต่ header)'); return; } // [ADM2-15a]
  var col = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // col A = REG_ID
  var realLastRow = 1;
  for (var r = 0; r < col.length; r++) {
    if (String(col[r][0] || '').trim() !== '') realLastRow = r + 2;
  }
  msg.push('getLastRow() รายงาน: ' + lastRow + ' แถว');
  msg.push('แถวข้อมูลจริงสุดท้าย (สแกนจาก REG_ID): ' + realLastRow + ' แถว');
  if (lastRow === realLastRow) {
    msg.push('✅ ตรงกัน — REGISTRATIONS ไม่มีปัญหาฟอร์แมตบวม');
  } else {
    msg.push('⚠️ ไม่ตรงกัน! มีช่องว่าง ' + (lastRow - realLastRow) + ' แถว — เสี่ยงมีข้อมูลกระโดดไปไกลเหมือนที่เจอในชีทรายการแข่งขัน');
  }

  // ── ตรวจที่ 3: มี REG_ID ซ้ำกันไหม ──────────────────────────
  msg.push('\n=== ตรวจ REG_ID ซ้ำ ===');
  var allIds = {};
  var dupes = [];
  col.forEach(function(row) {
    var rid = String(row[0] || '').trim();
    if (!rid) return;
    if (allIds[rid]) dupes.push(rid);
    allIds[rid] = true;
  });
  if (dupes.length === 0) {
    msg.push('✅ ไม่พบ REG_ID ซ้ำเลย');
  } else {
    msg.push('⚠️ พบ REG_ID ซ้ำ ' + dupes.length + ' รายการ: ' + dupes.join(', '));
  }

  Logger.log(msg.join('\n'));
  ui.alert('ผลตรวจสอบ REGISTRATIONS', msg.join('\n'), ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: แสดงรายละเอียดแถวที่มี REG_ID ซ้ำ (เพื่อดูก่อนตัดสินใจลบด้วยมือ)
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน showDuplicateRegDetails แล้วกด Run
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่แสดงผล
// ================================================================

function showDuplicateRegDetails() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH_REG);
  var ui = SpreadsheetApp.getUi();

  var data = sh.getDataRange().getValues();
  var hdr = data[0];
  var iN = hdr.indexOf('ชื่อเต็ม');
  var iCard = hdr.indexOf('เลขบัตรประชาชน');

  // หา REG_ID ที่ซ้ำ
  var seen = {}, dupIds = {};
  for (var i = 1; i < data.length; i++) {
    var rid = String(data[i][0] || '').trim();
    if (!rid) continue;
    if (seen[rid]) dupIds[rid] = true;
    seen[rid] = true;
  }

  var idsToCheck = Object.keys(dupIds);
  if (idsToCheck.length === 0) {
    ui.alert('✅ ไม่พบ REG_ID ซ้ำ');
    return;
  }

  var msg = [];
  idsToCheck.forEach(function(rid) {
    msg.push('=== REG_ID: ' + rid + ' ===');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === rid) {
        msg.push('แถวที่ ' + (i + 1) + ' | ชื่อ: ' + (data[i][iN] || '') +
                  ' | บัตร: ' + (data[i][iCard] || '') + ' | วันที่: ' + data[i][1]);
      }
    }
  });

  // เช็คว่าแถวซ้ำนี้ไปกระทบชีทอื่นด้วยไหม (COMPETITIONS, PAYMENTS, MEAL_SUMMARY, ชีทรายการแข่งขัน)
  var otherSheets = [SH_COMP, SH_PAY, SH_MEAL].concat(Object.keys(COMP_SHEET).map(function(cid){ return COMP_SHEET[cid]; }));
  msg.push('\n=== เช็คชีทอื่นที่มี REG_ID นี้ซ้ำด้วย ===');
  idsToCheck.forEach(function(rid) {
    otherSheets.forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (!s || s.getLastRow() < 1) return;
      var lastCol = s.getLastColumn();
      var headerRowIdx = -1, h2 = null;
      for (var scanRow = 1; scanRow <= Math.min(3, s.getLastRow()); scanRow++) {
        var candidate = s.getRange(scanRow, 1, 1, lastCol).getValues()[0];
        if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; h2 = candidate; break; }
      }
      if (headerRowIdx < 0 || s.getLastRow() <= headerRowIdx) return;
      var iRid = h2.indexOf('REG_ID');
      var d2 = s.getRange(headerRowIdx + 1, 1, s.getLastRow() - headerRowIdx, lastCol).getValues();
      var count = d2.filter(function(r) { return String(r[iRid] || '').trim() === rid; }).length;
      if (count > 1) msg.push(name + ' → พบ ' + count + ' แถว (ซ้ำ!)');
      else if (count === 1) msg.push(name + ' → พบ 1 แถว (ปกติ)');
    });
  });

  msg.push('\n⚠️ กรุณาไปลบแถวที่ซ้ำด้วยมือ (คลิกขวาที่เลขแถว → Delete row) เหลือไว้แค่ 1 แถวต่อ REG_ID ในแต่ละชีท');

  Logger.log(msg.join('\n'));
  ui.alert('รายละเอียด REG_ID ซ้ำ', msg.join('\n'), ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: แก้ REG_ID ชนกัน — ย้าย 1 คนไป REG_ID ใหม่ที่ปลอดภัย
//  ใช้แก้กรณี: INHI26-0714-0237 ถูกใช้โดย 2 คนคนละคนกัน
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน fixDuplicateRegId แล้วกด Run
//  จะถามทีละขั้น: REG_ID เดิม → เลขบัตรของคนที่จะย้าย → ชื่อ (สำหรับจับคู่ในชีทที่ไม่มีเลขบัตร)
//  แสดง preview ก่อนแก้จริงเสมอ ไม่แตะ PAYMENTS (ให้ไปแก้เองเพราะแยกเจ้าของยาก)
// ================================================================

function generateSafeRegId_(ss, dateStr) {
  // [ADM2-01] ตรวจความไม่ซ้ำกับ REGISTRATIONS + PAYMENTS (+ MEAL_SUMMARY / MEAL_CHECKIN ถ้ามี) ก่อนคืนค่า
  var existing = {};
  [SH_REG, SH_PAY, SH_MEAL, SH_CHECKIN].forEach(function(nm) {
    var sh = ss.getSheetByName(nm);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var v = String(r[0]).trim();
      if (v) existing[v] = true;
    });
  });
  var attempt;
  var tries = 0;
  do {
    var suffix = String(Math.floor(1000 + Math.random() * 9000));
    attempt = REG_PREFIX + '-' + dateStr + '-' + suffix;
    tries++;
  } while (existing[attempt] && tries < 500);
  if (existing[attempt]) throw new Error('สร้าง REG_ID ใหม่ที่ไม่ซ้ำไม่สำเร็จ — ลองใหม่');
  return attempt;
}

function fixDuplicateRegId() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.openById(SS_ID);

  var r1 = ui.prompt('ขั้นที่ 1/3', 'ใส่ REG_ID เดิมที่ชนกัน\nเช่น INHI26-0714-0237', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var oldRegId = r1.getResponseText().trim();
  if (!oldRegId) { ui.alert('ไม่ได้ใส่ REG_ID'); return; }

  var r2 = ui.prompt('ขั้นที่ 2/3', 'ใส่เลขบัตรประชาชนของ "คนที่จะย้ายไป REG_ID ใหม่"\nเช่น 5101400059933', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var targetIdcard = r2.getResponseText().trim();
  if (!targetIdcard) { ui.alert('ไม่ได้ใส่เลขบัตร'); return; }

  var r3 = ui.prompt('ขั้นที่ 3/3', 'ใส่ชื่อ (แค่ชื่อจริงพอ ไม่ต้องนามสกุล) ของคนนี้\nเช่น ดวงพร', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var targetName = r3.getResponseText().trim();
  if (!targetName) { ui.alert('ไม่ได้ใส่ชื่อ'); return; }

  var dateStr = oldRegId.split('-')[1] || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'MMdd');
  var newRegId = generateSafeRegId_(ss, dateStr);

  var plan = []; // {sheetName, rowNum, col(1-based), lastCol}
  var preview = [];
  var regFullName = '';   // [ADM2-06] ชื่อเต็มของผู้ลงทะเบียนที่จะย้าย (ใช้จับคู่ MEAL_SUMMARY / MEAL_CHECKIN)
  var movedCids = {};     // [ADM2-07] รหัสรายการที่จับคู่ด้วยเลขบัตรสำเร็จ

  // ── 1) REGISTRATIONS — match ด้วยเลขบัตรประชาชน ──────────────
  (function() {
    var sh = ss.getSheetByName(SH_REG);
    var data = sh.getDataRange().getValues();
    var hdr = data[0];
    var iCard = hdr.indexOf('เลขบัตรประชาชน');
    var iFull = hdr.indexOf('ชื่อเต็ม');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === oldRegId && String(data[i][iCard]).trim() === targetIdcard) {
        if (iFull >= 0) regFullName = String(data[i][iFull] || '').trim();
        plan.push({ sheet: SH_REG, row: i + 1, col: 1 });
        preview.push(SH_REG + ' แถว ' + (i + 1) + ' (คอลัมน์ A)');
      }
    }
  })();

  // ── 2) MEAL_SUMMARY + MEAL_CHECKIN — match ด้วย REG_ID + ชื่อเต็ม [ADM2-06] ──
  // (MEAL_SUMMARY คอลัมน์เลขบัตรถูกเขียนเป็นค่าว่าง จึงใช้เลขบัตรจับคู่ไม่ได้)
  [SH_MEAL, SH_CHECKIN].forEach(function(mealName) {
    var sh = ss.getSheetByName(mealName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    var hdr = data[0];
    var iCard = hdr.indexOf('เลขบัตรประชาชน');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== oldRegId) continue;
      var cardCell = iCard >= 0 ? String(data[i][iCard] || '').trim() : '';
      var ok = cardCell
        ? cardCell === targetIdcard
        : (regFullName && String(data[i][1] || '').trim() === regFullName);
      if (ok) {
        plan.push({ sheet: mealName, row: i + 1, col: 1 });
        preview.push(mealName + ' แถว ' + (i + 1) + ' (คอลัมน์ A)');
      }
    }
  });

  // ── 3) ชีทรายการแข่งขันทุกอัน — match ด้วยเลขบัตรใน 4 ช่องสมาชิก ──
  Object.keys(COMP_SHEET).forEach(function(cid) {
    if (cid === 'c00') return; // online ไม่มีทีม/เลขบัตรแยก
    var name = COMP_SHEET[cid];
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0) return;
    var iRegId = hdr.indexOf('REG_ID');
    var idCardCols = [];
    for (var n = 1; n <= 4; n++) {
      var c = hdr.indexOf('เลขบัตร ' + n);
      if (c >= 0) idCardCols.push(c);
    }
    if (sh.getLastRow() <= headerRowIdx) return;   // ชีทมีแต่ title+header
    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
    data.forEach(function(r, idx) {
      if (String(r[iRegId]).trim() !== oldRegId) return;
      var hasCard = idCardCols.some(function(c) { return String(r[c] || '').trim() === targetIdcard; });
      if (hasCard) {
        movedCids[cid] = true;
        plan.push({ sheet: name, row: headerRowIdx + 1 + idx, col: iRegId + 1 });
        preview.push(name + ' แถว ' + (headerRowIdx + 1 + idx) + ' (คอลัมน์ ' + _colLetter_(iRegId) + ')');
      }
    });
  });

  // ── 4) COMPETITIONS — ไม่มีเลขบัตร ใช้ชื่อเต็ม + เฉพาะรายการที่จับคู่ได้จากชีททีม [ADM2-07] ──
  (function() {
    var sh = ss.getSheetByName(SH_COMP);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    var hdr = data[0];
    var iMembers = hdr.indexOf('รายชื่อสมาชิก');
    var iCode = hdr.indexOf('รหัสรายการ');
    if (iMembers < 0 || iCode < 0) { preview.push('⚠️ COMPETITIONS header ไม่ครบ (รายชื่อสมาชิก/รหัสรายการ) — ข้ามส่วนนี้'); return; }
    var nameKey = regFullName || targetName;
    var cands = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === oldRegId &&
          movedCids[String(data[i][iCode]).trim()] &&
          String(data[i][iMembers] || '').indexOf(nameKey) >= 0) {
        cands.push(i);
      }
    }
    // ข้ามถ้าคลุมเครือ: รหัสรายการเดียวกันเจอมากกว่า 1 แถว
    var perCid = {};
    cands.forEach(function(i) { var c = String(data[i][iCode]).trim(); perCid[c] = (perCid[c] || 0) + 1; });
    cands.forEach(function(i) {
      var c = String(data[i][iCode]).trim();
      if (perCid[c] > 1) { preview.push(SH_COMP + ' แถว ' + (i + 1) + ' — ⚠️ คลุมเครือ (รายการ ' + c + ' ตรงหลายแถว) ข้าม ตรวจมือ'); return; }
      plan.push({ sheet: SH_COMP, row: i + 1, col: 1 });
      preview.push(SH_COMP + ' แถว ' + (i + 1) + ' (คอลัมน์ A) — จับคู่จากชื่อ + รายการที่ย้าย');
    });
  })();

  // ── PAYMENTS — ไม่แตะ แค่โชว์ให้ดู ──────────────────────────
  var paymentRows = [];
  (function() {
    var sh = ss.getSheetByName(SH_PAY);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === oldRegId) {
        paymentRows.push('แถว ' + (i + 1) + ' | จำนวนเงิน: ' + data[i][2] + ' | วันที่ส่ง: ' + data[i][1] + ' | ที่อยู่ใบเสร็จ: ' + (data[i][9] || '-'));
      }
    }
  })();

  if (plan.length === 0) {
    ui.alert('ไม่พบข้อมูล', 'ไม่พบแถวที่ตรงกับเลขบัตร/ชื่อที่ระบุ ลองตรวจสอบเลขบัตรหรือชื่ออีกครั้ง', ui.ButtonSet.OK);
    return;
  }

  var msg = 'REG_ID เดิม: ' + oldRegId + '\nREG_ID ใหม่ (สุ่มปลอดภัย): ' + newRegId +
    '\n\nจะแก้ทั้งหมด ' + plan.length + ' แถว:\n' + preview.join('\n');
  if (paymentRows.length) {
    msg += '\n\n⚠️ พบใน PAYMENTS ' + paymentRows.length + ' แถว (ไม่แก้อัตโนมัติ ต้องไปดูเองว่าอันไหนเป็นของ ' + targetName + '):\n' + paymentRows.join('\n');
  }
  msg += '\n\nยืนยันดำเนินการ?';

  var confirm = ui.alert('ตรวจสอบก่อนแก้', msg, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  // [ADM2-08] ล็อก + ตรวจซ้ำว่าแถวยังเป็น oldRegId ก่อนเขียนทับ
  var fixLock = LockService.getScriptLock();
  if (!fixLock.tryLock(30000)) { ui.alert('ระบบกำลังประมวลผลอยู่ ลองใหม่อีกครั้ง'); return; }
  try {
    var staleRows = [];
    plan.forEach(function(p) {
      var sh = ss.getSheetByName(p.sheet);
      if (String(sh.getRange(p.row, p.col).getValue() || '').trim() !== oldRegId) { staleRows.push(p.sheet + ' แถว ' + p.row); return; }
      sh.getRange(p.row, p.col).setValue(newRegId);
    });
    SpreadsheetApp.flush();
    if (staleRows.length) ui.alert('⚠️ ข้ามบางแถวเพราะข้อมูลเปลี่ยนไประหว่างรอ', staleRows.join('\n'), ui.ButtonSet.OK);
  } finally {
    fixLock.releaseLock();
  }

  var doneMsg = '✅ แก้ไขเสร็จแล้ว ' + plan.length + ' แถว\n\nREG_ID ใหม่ของ ' + targetName + ': ' + newRegId;
  if (paymentRows.length) {
    doneMsg += '\n\n⚠️ อย่าลืมไปแก้ PAYMENTS ด้วยมือ ' + paymentRows.length + ' แถวที่แจ้งไว้ข้างบน';
  }
  doneMsg += '\n\nแนะนำ: รัน exportFullRegistrationSummary, exportMealListByCompetition, exportItemFoodSummary ใหม่หลังจากนี้';

  Logger.log(doneMsg);
  ui.alert('เสร็จสิ้น', doneMsg, ui.ButtonSet.OK);
}

// ================================================================
//  UPDATE: moveTeamToCompetition v2
//  เพิ่ม: (1) popup เตือนส่วนต่างราคาที่จ่ายจริงกับราคาใหม่
//        (2) อัปเดตแถวใน MEAL_SUMMARY ให้ตรงรายการใหม่ด้วย (เดิมไม่ทำ)
//
//  วิธีติดตั้ง: ลบฟังก์ชัน moveTeamToCompetition, _recalcRegTotals,
//  และ generateSafeRegId_2_dummy เวอร์ชันเก่าออกทั้งหมด แล้ววางไฟล์นี้แทนที่
//  (ถ้าเพิ่งติดตั้งเวอร์ชันแรกไปแล้ว ให้ลบของเดิมทิ้งก่อน อย่าวางซ้อนกัน)
// ================================================================

function moveTeamToCompetition() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.openById(SS_ID);

  var r1 = ui.prompt('ย้ายทีมไปรายการแข่งขันอื่น (1/2)', 'ใส่ REG_ID ของทีมที่ต้องการย้าย', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var regId = r1.getResponseText().trim();
  if (!regId) { ui.alert('ไม่ได้ใส่ REG_ID'); return; }

  var foundIn = [];
  Object.keys(COMP_SHEET).forEach(function(cid) {
    if (cid === 'c00') return;
    var name = COMP_SHEET[cid];
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0) return;
    var iRegId = hdr.indexOf('REG_ID');
    if (sh.getLastRow() <= headerRowIdx) return;   // ชีทมีแต่ title+header
    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
    data.forEach(function(r, idx) {
      if (String(r[iRegId]).trim() === regId) {
        foundIn.push({ cid: cid, name: name, sh: sh, headerRowIdx: headerRowIdx, hdr: hdr, rowNum: headerRowIdx + 1 + idx, rowData: r });
      }
    });
  });

  if (foundIn.length === 0) {
    ui.alert('ไม่พบ REG_ID นี้ในรายการแข่งขันใดๆ (ไม่รวมประชุมออนไลน์)');
    return;
  }

  var srcPick;
  if (foundIn.length === 1) {
    srcPick = foundIn[0];
  } else {
    var listMsg = foundIn.map(function(f, i) { return (i + 1) + '. ' + COMP[f.cid].nm + ' (' + f.name + ')'; }).join('\n');
    var r2 = ui.prompt('พบหลายรายการ — เลือกอันที่จะย้าย', listMsg + '\n\nใส่หมายเลข (1-' + foundIn.length + ')', ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    var pickIdx = parseInt(r2.getResponseText().trim(), 10) - 1;
    if (isNaN(pickIdx) || pickIdx < 0 || pickIdx >= foundIn.length) { ui.alert('เลือกไม่ถูกต้อง'); return; }
    srcPick = foundIn[pickIdx];
  }

  var srcCfg = COMP[srcPick.cid];

  var destList = Object.keys(COMP_SHEET).filter(function(cid) { return cid !== 'c00' && cid !== srcPick.cid; });
  var destMsg = destList.map(function(cid, i) { return (i + 1) + '. ' + COMP[cid].nm; }).join('\n');
  var r3 = ui.prompt('ย้ายไปรายการไหน (2/2)', 'จากรายการ: ' + srcCfg.nm + '\n\n' + destMsg + '\n\nใส่หมายเลข', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var destIdx = parseInt(r3.getResponseText().trim(), 10) - 1;
  if (isNaN(destIdx) || destIdx < 0 || destIdx >= destList.length) { ui.alert('เลือกไม่ถูกต้อง'); return; }
  var destCid = destList[destIdx];
  var destCfg = COMP[destCid];
  var destSheetName = COMP_SHEET[destCid];

  var srcHdr = srcPick.hdr;
  var slotCols = [];
  for (var n = 1; n <= 4; n++) {
    slotCols.push({
      fn: srcHdr.indexOf('ชื่อ ' + n),
      combined: srcHdr.indexOf('ชื่อ สมาชิก ' + n)
    });
  }
  var filledCount = 0;
  slotCols.forEach(function(col) {
    var nm = col.fn >= 0 ? srcPick.rowData[col.fn] : (col.combined >= 0 ? srcPick.rowData[col.combined] : '');
    if (String(nm || '').trim()) filledCount++;
  });

  var warnMsg = '';
  if (filledCount > destCfg.team) {
    warnMsg = '\n\n⚠️ คำเตือน: ทีมนี้กรอกสมาชิกไว้ ' + filledCount + ' คน แต่ "' + destCfg.nm + '" รับแค่ ' + destCfg.team + ' คน/ทีม';
  }

  var daysColIdx = srcHdr.indexOf('จำนวนวัน');
  var days = daysColIdx >= 0 ? (Number(srcPick.rowData[daysColIdx]) || 1) : 1;
  var newFee = (days === 2 && destCfg.d2) ? destCfg.d2 : destCfg.d1;

  // ── เช็คราคาที่เคยจ่ายจริงจาก PAYMENTS (เอาแถวล่าสุดของ REG_ID นี้) ──
  var paidAmount = null;
  (function() {
    var psh = ss.getSheetByName(SH_PAY);
    if (!psh || psh.getLastRow() < 2) return;
    var pData = psh.getDataRange().getValues();
    for (var i = pData.length - 1; i >= 1; i--) {
      if (String(pData[i][0]).trim() === regId) { paidAmount = Number(pData[i][2]) || 0; break; }
    }
  })();

  var paidMsg = paidAmount === null
    ? '\nยังไม่พบประวัติการโอนเงินใน PAYMENTS สำหรับ REG_ID นี้'
    : '\nเคยโอนมาแล้ว: ' + paidAmount.toLocaleString() + ' บาท (รายการนี้อย่างเดียวราคาใหม่ ' + newFee.toLocaleString() + ' บาท)';

  var confirm = ui.alert(
    'ยืนยันการย้ายทีม',
    'REG_ID: ' + regId +
    '\nจาก: ' + srcCfg.nm +
    '\nไป: ' + destCfg.nm +
    '\nจำนวนวันเดิม: ' + days + ' วัน → ค่าธรรมเนียมรายการนี้ใหม่: ' + newFee.toLocaleString() + ' บาท' +
    paidMsg + warnMsg +
    '\n\n⚠️ จะย้ายแถวข้ามชีท + อัปเดต COMPETITIONS + MEAL_SUMMARY + คำนวณยอดรวม/สิทธิ์อาหารใน REGISTRATIONS ใหม่ทั้งหมด\n\nยืนยันดำเนินการ?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // [ADM2-03] ล็อก — ตรวจทุกอย่างก่อน (validate) แล้วค่อยเขียนตามลำดับ: append ปลายทาง -> ลบต้นทาง -> อัปเดตชีทอื่น
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { ui.alert('ระบบกำลังประมวลผลอยู่ ลองใหม่อีกครั้ง'); return; }
  var done = [];
  var mealUpdated = false;
  try {
    // ── VALIDATE (ยังไม่เขียนข้อมูลใดๆ) ──
    // [ADM2-02] ชีทปลายทางอาจยังไม่ถูกสร้าง — สร้างด้วย _compSheet (วิธีเดียวกับตอนลงทะเบียน)
    var destSh = ss.getSheetByName(destSheetName);
    if (!destSh) {
      destSh = _compSheet(ss, destSheetName, destCfg.nm, destCfg.nm_en);
    }
    if (!destSh) { ui.alert('ไม่พบ/สร้างชีทปลายทาง "' + destSheetName + '" ไม่ได้ — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return; }

    var destLastCol = destSh.getLastColumn();
    var destHeaderRowIdx = -1, destHdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, destSh.getLastRow() || 1); scanRow++) {
      var candidate = destSh.getRange(scanRow, 1, 1, destLastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { destHeaderRowIdx = scanRow; destHdr = candidate; break; }
    }
    if (destHeaderRowIdx < 0) { ui.alert('หา header ของชีทปลายทางไม่เจอ — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return; }

    // [ADM2-04b] REG_ID ต้องยังไม่มีในปลายทาง
    var iDR = destHdr.indexOf('REG_ID');
    if (destSh.getLastRow() > destHeaderRowIdx) {
      var destIds = destSh.getRange(destHeaderRowIdx + 1, iDR + 1, destSh.getLastRow() - destHeaderRowIdx, 1).getValues();
      var dupDest = destIds.some(function(r) { return String(r[0]).trim() === regId; });
      if (dupDest) { ui.alert('REG_ID นี้มีอยู่ในชีทปลายทางแล้ว — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return; }
    }

    // header ของ COMPETITIONS / REGISTRATIONS ต้องครบก่อนเริ่มเขียน
    var compSh = ss.getSheetByName(SH_COMP);
    var cData = null, cHdr = null, iCRid = -1, iCCode = -1, iCName = -1, iCQuota = -1;
    if (compSh && compSh.getLastRow() > 1) {
      cData = compSh.getDataRange().getValues();
      cHdr = cData[0];
      iCRid = cHdr.indexOf('REG_ID'); iCCode = cHdr.indexOf('รหัสรายการ');
      iCName = cHdr.indexOf('ชื่อรายการ'); iCQuota = cHdr.indexOf('โควต้าทีม');
      if (iCRid < 0 || iCCode < 0 || iCName < 0 || iCQuota < 0) {
        ui.alert('header ของ COMPETITIONS ไม่ครบ — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return;
      }
    }
    var regChk = ss.getSheetByName(SH_REG);
    if (!regChk || regChk.getLastRow() < 2) { ui.alert('ไม่พบข้อมูล REGISTRATIONS — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return; }
    var regChkHdr = regChk.getRange(1, 1, 1, regChk.getLastColumn()).getValues()[0];
    if (['ยอดรวม', 'อาหาร_27พค', 'อาหาร_28พค', 'คนได้ข้าว_รวม'].some(function(h) { return regChkHdr.indexOf(h) < 0; })) {
      ui.alert('header ของ REGISTRATIONS ไม่ครบ (ยอดรวม/อาหาร) — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return;
    }

    // [ADM2-03] หาแถวต้นทางใหม่ด้วย REG_ID (ไม่ใช้เลขแถวที่จำไว้ก่อน dialog)
    var srcRegCol = srcPick.hdr.indexOf('REG_ID') + 1;
    var srcRow = -1;
    if (srcPick.sh.getLastRow() > srcPick.headerRowIdx) {
      var srcIds = srcPick.sh.getRange(srcPick.headerRowIdx + 1, srcRegCol, srcPick.sh.getLastRow() - srcPick.headerRowIdx, 1).getValues();
      var hits = [];
      srcIds.forEach(function(r, ix) { if (String(r[0]).trim() === regId) hits.push(srcPick.headerRowIdx + 1 + ix); });
      if (hits.length === 1) srcRow = hits[0];
      else if (hits.length > 1) { ui.alert('พบ REG_ID นี้ ' + hits.length + ' แถวในชีทต้นทาง — ยกเลิก (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return; }
    }
    if (srcRow < 0) { ui.alert('แถวต้นทางหายไประหว่างรอยืนยัน — ยกเลิก กรุณารันใหม่ (ยังไม่มีการเปลี่ยนแปลงข้อมูล)'); return; }
    var srcRowData = srcPick.sh.getRange(srcRow, 1, 1, srcPick.hdr.length).getValues()[0];

    var newRow = new Array(destLastCol).fill('');
    destHdr.forEach(function(colName, di) {
      var si = srcHdr.indexOf(colName);
      if (si >= 0) newRow[di] = srcRowData[si];
    });
    var iFeeDest = destHdr.indexOf('ค่าธรรมเนียม (บาท)');
    if (iFeeDest >= 0) newRow[iFeeDest] = newFee;
    var iDateDest = destHdr.indexOf('วันที่บันทึก');
    if (iDateDest >= 0) newRow[iDateDest] = new Date();
    // [ADM2-04a] ลำดับทีมใหม่ในปลายทาง (ตามวิธีเดียวกับ _writeCompSheets: getLastRow ก่อน append)
    var iSeq = destHdr.indexOf('ลำดับทีม');
    if (iSeq >= 0) newRow[iSeq] = destSh.getLastRow();

    // ── WRITE: append ปลายทาง -> ลบต้นทาง -> อัปเดตชีทอื่น ──
    destSh.appendRow(newRow);
    done.push('เพิ่มแถวในชีทปลายทาง');
    srcPick.sh.deleteRow(srcRow);
    done.push('ลบแถวในชีทต้นทาง');

    // ── อัปเดต COMPETITIONS ──
    if (cData) {
      for (var ci = 1; ci < cData.length; ci++) {
        if (String(cData[ci][iCRid]).trim() === regId && String(cData[ci][iCCode]).trim() === srcPick.cid) {
          compSh.getRange(ci + 1, iCCode + 1).setValue(destCid);
          compSh.getRange(ci + 1, iCName + 1).setValue(destCfg.nm);
          compSh.getRange(ci + 1, iCQuota + 1).setValue(destCfg.team);
          break;
        }
      }
      done.push('อัปเดต COMPETITIONS');
    }

    // ── อัปเดต MEAL_SUMMARY ให้ตรงรายการใหม่ ──
    mealUpdated = _updateMealSummaryRow(ss, regId, srcCfg.nm, destCfg, days);
    done.push('อัปเดต MEAL_SUMMARY');

    // ── คำนวณค่าธรรมเนียมรวม + สิทธิ์อาหารใหม่ทั้งหมดของ REG_ID นี้ ──
    _recalcRegTotals(ss, regId);
    done.push('คำนวณยอดรวม/สิทธิ์อาหารใน REGISTRATIONS');

    SpreadsheetApp.flush();
  } catch (moveErr) {
    Logger.log('moveTeamToCompetition error: ' + moveErr.message + ' | done: ' + done.join(', '));
    ui.alert(
      '⚠️ ย้ายทีมไม่สมบูรณ์',
      'เกิดข้อผิดพลาด: ' + moveErr.message +
      '\n\nขั้นที่ทำสำเร็จแล้ว: ' + (done.length ? done.join(' → ') : '(ไม่มี)') +
      '\n\nกรุณาตรวจสอบ REG_ID ' + regId + ' ด้วยมือก่อนรันซ้ำ',
      ui.ButtonSet.OK
    );
    return;
  } finally {
    lock.releaseLock();
  }

  // ── Auto-export: อัปเดต 4 ชีทสรุป/รายชื่ออาหาร ให้ตรงหลังย้ายทีม ──
  try{ _scheduleAutoExports(); }catch(autoErr){ Logger.log('schedule auto export err: '+autoErr.message); }

  var diffMsg = '';
  if (paidAmount !== null) {
    var diff = newFee - paidAmount;
    if (diff === 0) diffMsg = '\n\n✅ ราคาเท่ากับที่เคยจ่ายไว้พอดี ไม่ต้องเก็บ/คืนเงินเพิ่ม';
    else if (diff > 0) diffMsg = '\n\nราคาใหม่แพงกว่าที่เคยจ่าย ' + diff.toLocaleString() + ' บาท — ต้องเก็บเงินเพิ่ม';
    else diffMsg = '\n\nราคาใหม่ถูกกว่าที่เคยจ่าย ' + Math.abs(diff).toLocaleString() + ' บาท — ควรคืนเงินส่วนต่าง';
  } else {
    diffMsg = '\n\n(ยังไม่เคยพบประวัติโอนเงิน — ไม่ต้องกังวลเรื่องส่วนต่าง)';
  }

  ui.alert(
    '✅ ย้ายทีมเสร็จแล้ว',
    'ย้าย REG_ID ' + regId + ' จาก "' + srcCfg.nm + '" ไป "' + destCfg.nm + '" เรียบร้อย' +
    '\n\nMEAL_SUMMARY: ' + (mealUpdated ? 'อัปเดตแล้ว ✓' : 'ไม่พบแถวเดิมให้อัปเดต (ข้ามไป)') +
   diffMsg +
    '\n\n✅ อัปเดตชีทสรุปทั้ง 4 ให้อัตโนมัติแล้ว (จะเสร็จภายใน ~20 วิ)',
    ui.ButtonSet.OK
  );
}

// helper: อัปเดตแถวใน MEAL_SUMMARY ให้ตรงรายการ/สิทธิ์อาหารใหม่
function _updateMealSummaryRow(ss, regId, oldCompName, destCfg, days) {
  var sh = ss.getSheetByName(SH_MEAL);
  if (!sh || sh.getLastRow() < 2) return false;
  var data = sh.getDataRange().getValues();
  var hdr = data[0];
  var iRid = hdr.indexOf('REG_ID'), iSrc = hdr.indexOf('รายการ/ที่มา'),
      iD17 = hdr.indexOf('อาหาร_27พค'), iD18 = hdr.indexOf('อาหาร_28พค');

  var getsD17 = destCfg.mealDay === 'both' || destCfg.mealDay === 'd17' || (destCfg.mealDay === 'd18' && days === 2);
  var getsD18 = destCfg.mealDay === 'both' || destCfg.mealDay === 'd18' || (destCfg.mealDay === 'd17' && days === 2);
  var q = destCfg.team || 1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iRid]).trim() === regId && String(data[i][iSrc] || '').indexOf(oldCompName) >= 0) {
      sh.getRange(i + 1, iSrc + 1).setValue(destCfg.nm);
      sh.getRange(i + 1, iD17 + 1).setValue(getsD17 ? q : 0);
      sh.getRange(i + 1, iD18 + 1).setValue(getsD18 ? q : 0);
      return true;
    }
  }
  return false;
}

// helper: คำนวณยอดรวม + สิทธิ์อาหารใหม่ทั้งหมดของ REG_ID จากข้อมูลใน COMPETITIONS ปัจจุบัน
function _recalcRegTotals(ss, regId) {
  var compSh = ss.getSheetByName(SH_COMP);
  var comps = [];
  if (compSh && compSh.getLastRow() > 1) {
    var cData = compSh.getDataRange().getValues();
    var cHdr = cData[0];
    var iCRid = cHdr.indexOf('REG_ID'), iCCode = cHdr.indexOf('รหัสรายการ'), iCDays = cHdr.indexOf('จำนวนวัน');
    for (var i = 1; i < cData.length; i++) {
      if (String(cData[i][iCRid]).trim() === regId) {
        comps.push({ cid: String(cData[i][iCCode]).trim(), days: Number(cData[i][iCDays]) || 1 });
      }
    }
  }

  // [ADM2-05] ไม่พบรายการใน COMPETITIONS -> ห้ามเขียนทับยอดด้วย 0
  if (!comps.length) { Logger.log('_recalcRegTotals skipped: no COMPETITIONS rows for ' + regId); return; }

  var fee = 0, d17 = 0, d18 = 0;
  comps.forEach(function(c) {
    var cfg = COMP[c.cid];
    if (!cfg) { Logger.log('_recalcRegTotals: unknown cid "' + c.cid + '" for ' + regId); return; }
    fee += (c.days === 2 && cfg.d2) ? cfg.d2 : cfg.d1;
    var q = cfg.team || 1;
    var getsD17 = cfg.mealDay === 'both' || cfg.mealDay === 'd17' || (cfg.mealDay === 'd18' && c.days === 2);
    var getsD18 = cfg.mealDay === 'both' || cfg.mealDay === 'd18' || (cfg.mealDay === 'd17' && c.days === 2);
    if (getsD17) d17 += q;
    if (getsD18) d18 += q;
  });

  var regSh = ss.getSheetByName(SH_REG);
  var rData = regSh.getDataRange().getValues();
  var rHdr = rData[0];
  var iFee = rHdr.indexOf('ยอดรวม'), iD17 = rHdr.indexOf('อาหาร_27พค'),
      iD18 = rHdr.indexOf('อาหาร_28พค'), iTotal = rHdr.indexOf('คนได้ข้าว_รวม');
  if ([iFee, iD17, iD18, iTotal].some(function(x) { return x < 0; })) throw new Error('REGISTRATIONS header mismatch');
  for (var r = 1; r < rData.length; r++) {
    if (String(rData[r][0]).trim() === regId) {
      regSh.getRange(r + 1, iFee + 1).setValue(fee);
      regSh.getRange(r + 1, iD17 + 1).setValue(d17);
      regSh.getRange(r + 1, iD18 + 1).setValue(d18);
      regSh.getRange(r + 1, iTotal + 1).setValue(d17 + d18);
      break;
    }
  }
}

// ================================================================
//  ADD-ON: หารายการที่ลงทะเบียนไว้ (มีใน COMPETITIONS) แต่ไม่มีแถวจริง
//          อยู่ในชีทรายการแข่งขันที่ควรจะอยู่ — เช็คทั้ง 2 ทิศทาง
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน findMissingCompSheetRows แล้วกด Run
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่แสดงผล
// ================================================================

function findMissingCompSheetRows() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  // ── เตรียมข้อมูล REG_ID ที่มีจริงในแต่ละชีทรายการแข่งขัน ──────
  var sheetRegIds = {}; // { sheetName: {regId: true, ...} }
  Object.keys(COMP_SHEET).forEach(function(cid) {
    var name = COMP_SHEET[cid];
    var sh = ss.getSheetByName(name);
    sheetRegIds[name] = {};
    if (!sh || sh.getLastRow() < 2) return;

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0 || sh.getLastRow() <= headerRowIdx) return;   // ชีทมีแต่ title+header ยังไม่มีข้อมูล
    var iRegId = hdr.indexOf('REG_ID');
    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
    data.forEach(function(r) {
      var rid = String(r[iRegId] || '').trim();
      if (rid) sheetRegIds[name][rid] = (sheetRegIds[name][rid] || 0) + 1;
    });
  });

  // ── เช็คทิศทางที่ 1: มีใน COMPETITIONS แต่ไม่มีแถวจริงในชีทรายการ ──
  var missing = [];
  var compSh = ss.getSheetByName(SH_COMP);
  if (compSh && compSh.getLastRow() > 1) {
    var cData = compSh.getDataRange().getValues();
    var cHdr = cData[0];
    var iCRid = cHdr.indexOf('REG_ID'), iCCode = cHdr.indexOf('รหัสรายการ'), iCName = cHdr.indexOf('ชื่อรายการ');
    for (var i = 1; i < cData.length; i++) {
      var rid = String(cData[i][iCRid] || '').trim();
      var cid = String(cData[i][iCCode] || '').trim();
      if (!rid || !cid) continue;
      var targetSheet = COMP_SHEET[cid];
      if (!targetSheet) { missing.push(rid + ' | รหัสรายการ "' + cid + '" ไม่มีอยู่จริงใน COMP_SHEET (พิมพ์ผิด?)'); continue; }
      var count = (sheetRegIds[targetSheet] || {})[rid] || 0;
      if (count === 0) {
        missing.push(rid + ' | ' + (cData[i][iCName] || cid) + ' → ไม่พบแถวจริงในชีท "' + targetSheet + '" เลย (แถว COMPETITIONS ที่ ' + (i + 1) + ')');
      }
    }
  }

  // ── เช็คทิศทางที่ 2: มีแถวในชีทรายการ แต่ REG_ID ไม่มีใน REGISTRATIONS เลย ──
  var orphans = [];
  var regSh = ss.getSheetByName(SH_REG);
  var validRegIds = {};
  if (regSh && regSh.getLastRow() > 1) {
    var rData = regSh.getRange(2, 1, regSh.getLastRow() - 1, 1).getValues();
    rData.forEach(function(r) { var rid = String(r[0] || '').trim(); if (rid) validRegIds[rid] = true; });
  }
  Object.keys(sheetRegIds).forEach(function(sheetName) {
    Object.keys(sheetRegIds[sheetName]).forEach(function(rid) {
      if (!validRegIds[rid]) {
        orphans.push(rid + ' | อยู่ในชีท "' + sheetName + '" แต่ไม่มีใน REGISTRATIONS เลย (อาจเป็นแถวค้าง/ทดสอบเก่า)');
      }
    });
  });

  var msg = [];
  msg.push('=== รายการที่ลงทะเบียนไว้ แต่หาแถวจริงในชีทรายการแข่งขันไม่เจอ ===');
  msg.push(missing.length ? missing.join('\n') : '✅ ไม่พบ — ทุกรายการมีแถวจริงครบ');

  msg.push('\n=== รายการที่มีแถวในชีทรายการแข่งขัน แต่ไม่มีใน REGISTRATIONS ===');
  msg.push(orphans.length ? orphans.join('\n') : '✅ ไม่พบ — ไม่มีแถวลอยค้าง');

  Logger.log(msg.join('\n'));
  ui.alert(
    'ผลตรวจสอบความสอดคล้องของข้อมูล',
    msg.join('\n') + '\n\nรวมพบปัญหาทั้งหมด: ' + (missing.length + orphans.length) + ' รายการ',
    ui.ButtonSet.OK
  );
}

// ================================================================
//  ADD-ON: ดีบักการอ่านข้อมูลของ exportMealListByCompetition แบบละเอียด
//  โชว์ทีละแถวว่าอ่าน REG_ID / ชื่อทีม / สมาชิกแต่ละคนได้อะไรบ้าง
//  พร้อมบอกว่า header คอลัมน์ที่เจอคือแบบไหน (แยก/รวม/หาไม่เจอเลย)
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน debugMealListExtraction แล้วกด Run
//           จะถามชื่อชีท เช่น SIM_MAN (พิมพ์ให้ตรงเป๊ะ)
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่แสดงผลผ่าน Logger + popup
// ================================================================

function debugMealListExtraction() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.openById(SS_ID);

  var r1 = ui.prompt('ดีบักชีทไหน', 'พิมพ์ชื่อชีทให้ตรงเป๊ะ เช่น SIM_MAN หรือ ตอบคำถาม_พยาบาล', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var shName = r1.getResponseText().trim();

  var sh = ss.getSheetByName(shName);
  if (!sh) { ui.alert('ไม่พบชีทชื่อนี้ พิมพ์ชื่อผิดหรือเปล่า'); return; }
  if (sh.getLastRow() < 2) { ui.alert('ชีทนี้ไม่มีข้อมูล'); return; }

  var lastCol = sh.getLastColumn();
  var headerRowIdx = -1, hdr = null;
  for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
    var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
    if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
  }
  if (headerRowIdx < 0) { ui.alert('หา header (REG_ID) ไม่เจอในชีทนี้เลย 3 แถวแรก'); return; }
  if (sh.getLastRow() <= headerRowIdx) { ui.alert('ชีทนี้มีแต่ header ยังไม่มีข้อมูล'); return; }

  var msg = [];
  msg.push('=== Header พบที่แถว ' + headerRowIdx + ' | ทั้งหมด ' + lastCol + ' คอลัมน์ ===');

  var iRegId = hdr.indexOf('REG_ID');
  var iTeam = hdr.indexOf('ชื่อทีม');
  var iDays = hdr.indexOf('จำนวนวัน');
  msg.push('REG_ID พบที่ col ' + (iRegId >= 0 ? _colLetter_(iRegId) : '❌ ไม่พบ'));
  msg.push('ชื่อทีม พบที่ col ' + (iTeam >= 0 ? _colLetter_(iTeam) : '❌ ไม่พบ'));
  msg.push('จำนวนวัน พบที่ col ' + (iDays >= 0 ? _colLetter_(iDays) : '❌ ไม่พบ'));

  var slotCols = [];
  for (var n = 1; n <= 4; n++) {
    var pre = hdr.indexOf('คำนำหน้า ' + n);
    var fn = hdr.indexOf('ชื่อ ' + n);
    var ln = hdr.indexOf('นามสกุล ' + n);
    var combined = hdr.indexOf('ชื่อ สมาชิก ' + n);
    slotCols.push({ pre: pre, fn: fn, ln: ln, combined: combined });
    var mode = fn >= 0 ? 'แยกคอลัมน์ (ชื่อ ' + n + '=col ' + _colLetter_(fn) + ', นามสกุล ' + n + '=col ' + (ln >= 0 ? _colLetter_(ln) : '❌') + ')'
             : combined >= 0 ? 'รวมคอลัมน์เดียว (ชื่อ สมาชิก ' + n + '=col ' + _colLetter_(combined) + ')'
             : '❌ หาคอลัมน์สมาชิกคนที่ ' + n + ' ไม่เจอเลย!';
    msg.push('สมาชิกที่ ' + n + ' → ' + mode);
  }

  msg.push('\n=== อ่านข้อมูลแต่ละแถวได้อะไรบ้าง ===');
  var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
  data.forEach(function(r, idx) {
    var regId = iRegId >= 0 ? r[iRegId] : '';
    if (!String(regId || '').trim()) return;
    var teamName = iTeam >= 0 ? r[iTeam] : '';
    var lineOut = 'แถว ' + (headerRowIdx + 1 + idx) + ' | ' + regId + ' | ทีม: ' + teamName + ' → ';
    var names = [];
    slotCols.forEach(function(col, i) {
      var fullName = '';
      if (col.fn >= 0) {
        var pre = col.pre >= 0 ? r[col.pre] : '';
        var ln = col.ln >= 0 ? r[col.ln] : '';
        fullName = (String(pre || '') + ' ' + String(r[col.fn] || '') + ' ' + String(ln || '')).trim();
      } else if (col.combined >= 0) {
        fullName = String(r[col.combined] || '').trim();
      }
      names.push(fullName ? ('[' + (i + 1) + ']' + fullName) : ('[' + (i + 1) + ']ว่าง'));
    });
    lineOut += names.join(' , ');
    msg.push(lineOut);
  });

  Logger.log(msg.join('\n'));
  // popup จำกัดความยาว ถ้าเยอะไปจะโชว์แค่บางส่วน เต็มๆ ดูใน View > Logs
  var full = msg.join('\n');
  var shown = full.length > 3800 ? full.slice(0, 3800) + '\n\n...(ยาวเกิน ดูฉบับเต็มที่ View > Logs)' : full;
  ui.alert('ผลดีบัก: ' + shName, shown, ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: เมนูกดง่ายๆ ในสเปรดชีต สำหรับอัปเดตชีทสรุปการลงทะเบียน/อาหาร
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก แล้ว Save
//  จากนั้นรีเฟรชหน้าสเปรดชีต (F5) จะเห็นเมนู "รายงานสรุป" โผล่ขึ้นมาบนแถบเมนู
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่เพิ่มเมนูเรียกฟังก์ชันที่มีอยู่แล้ว
// ================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('รายงานสรุป')
    .addItem('1) อัปเดตรายชื่ออาหารรายบุคคล (แยกตามรายการแข่งขัน)', 'exportMealListByCompetition')
    .addItem('2) อัปเดตสรุปจำนวนรายการ+อาหารแต่ละวัน', 'exportItemFoodSummary')
    .addSeparator()
    .addItem('3) อัปเดตสรุปการลงทะเบียนทั้งหมด', 'exportFullRegistrationSummary')
    .addSeparator()
    .addItem('▶ อัปเดตทั้งหมดในคลิกเดียว', 'runAllSummaryReports')
    .addSeparator()
    .addItem('✉️ ส่งอีเมลยืนยันซ้ำ (REG ID / ชื่อ)', 'resendConfirmationEmail')
    .addItem('✉️ ส่งอีเมลค้างในคิวตอนนี้', 'sendPendingEmailsNow')
    .addToUi();
}

// รันรายงานสรุปทั้ง 3 ชีทพร้อมกัน แล้วโชว์สรุปผลครั้งเดียว
function runAllSummaryReports() {
  var ui = SpreadsheetApp.getUi();
  // [ADM2-18] แยก try/catch ต่อรายงาน และแจ้งว่าอันไหนสำเร็จ/ล้มเหลว
  var steps = [
    ['รายชื่ออาหารรายบุคคล', function() { exportMealListByCompetition(true); }],   // true = silent
    ['สรุปจำนวนรายการและอาหาร', function() { exportItemFoodSummary(true); }],
    ['สรุปการลงทะเบียนทั้งหมด', function() { exportFullRegistrationSummary(true); }]
  ];
  var okList = [], failList = [];
  steps.forEach(function(st) {
    try { st[1](); okList.push('• ' + st[0]); }
    catch (e) { failList.push('• ' + st[0] + ' ✗ ' + e.message); }
  });
  ui.alert(
    failList.length ? 'อัปเดตรายงานสรุปไม่ครบ ⚠️' : 'อัปเดตรายงานสรุปเสร็จแล้ว ✅',
    (okList.length ? 'สำเร็จ:\n' + okList.join('\n') : '') +
    (failList.length ? '\n\nล้มเหลว:\n' + failList.join('\n') : ''),
    ui.ButtonSet.OK
  );
}

// ================================================================
//  ADD-ON: เช็คว่าชีทรายการแข่งขันไหน "จำนวนทีม" ไม่ตรงกับ
//          ชีท "รายชื่ออาหารรายบุคคล" บ้าง
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน checkTeamCountVsMealList แล้วกด Run
//           (หรือกดผ่านเมนู รายงานสรุป ถ้าเพิ่มเมนูไว้แล้ว — ดูด้านล่าง)
//  ไม่แก้ไข/ไม่ลบข้อมูลใดๆ ทั้งสิ้น แค่แสดงผลผ่าน Logger + popup
// ================================================================

function checkTeamCountVsMealList() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  // ── 1) นับทีมจริงในแต่ละชีทรายการแข่งขัน (นับจาก REG_ID ที่ไม่ว่าง) ──────
  var compTeams = {}; // { cid: { nm, mealDay, regIds:{regId:true} } }
  Object.keys(COMP_SHEET).forEach(function (cid) {
    var cfg = COMP[cid];
    if (!cfg) return;
    var shName = COMP_SHEET[cid];
    var sh = ss.getSheetByName(shName);
    var regIds = {};

    if (sh && sh.getLastRow() >= 2) {
      var lastCol = sh.getLastColumn();
      var headerRowIdx = -1, hdr = null;
      for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
        var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
        if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
      }
      if (headerRowIdx >= 0 && sh.getLastRow() > headerRowIdx) {
        var iRegId = hdr.indexOf('REG_ID');
        var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
        data.forEach(function (r) {
          var rid = String((iRegId >= 0 ? r[iRegId] : '') || '').trim();
          if (rid) regIds[rid] = true;
        });
      }
    }
    compTeams[cid] = { nm: cfg.nm, mealDay: cfg.mealDay, regIds: regIds };
  });

  // ── 2) นับทีมที่ปรากฏจริงในชีท "รายชื่ออาหารรายบุคคล" แยกตามชื่อรายการ ──
  //     (กรองด้วย REG_PREFIX เพื่อกันแถวตาราง "สรุปจำนวนอาหาร" / แถวเตือนท้ายชีท หลุดเข้ามา)
  var mealTeams = {}; // { ชื่อรายการ: {regId:true} }
  var mealSh = ss.getSheetByName(SH_MEALLIST);
  if (mealSh && mealSh.getLastRow() >= 2) {
    var mData = mealSh.getRange(2, 1, mealSh.getLastRow() - 1, 7).getValues();
    mData.forEach(function (r) {
      var nm = String(r[0] || '').trim();
      var regId = String(r[2] || '').trim();
      if (!nm || !regId) return;
      if (regId.indexOf(REG_PREFIX) !== 0) return; // ไม่ใช่แถวข้อมูลจริง ข้าม
      if (!mealTeams[nm]) mealTeams[nm] = {};
      mealTeams[nm][regId] = true;
    });
  } else {
    ui.alert('ยังไม่มีข้อมูลในชีท "' + SH_MEALLIST + '"', 'กรุณารัน exportMealListByCompetition ก่อน แล้วค่อยเช็คใหม่', ui.ButtonSet.OK);
    return;
  }

  // ── 3) เทียบทีละรายการแข่งขัน (ข้าม mealDay='none' เช่น ประชุม Online เพราะไม่มีอาหารเลย) ──
  var mismatches = [];
  var okList = [];
  Object.keys(compTeams).forEach(function (cid) {
    var info = compTeams[cid];
    if (info.mealDay === 'none') return;

    var compRegIds = Object.keys(info.regIds);
    var mealRegIds = Object.keys(mealTeams[info.nm] || {});

    var missingFromMealList = compRegIds.filter(function (id) { return mealRegIds.indexOf(id) < 0; });
    var extraInMealList     = mealRegIds.filter(function (id) { return compRegIds.indexOf(id) < 0; });

    if (missingFromMealList.length || extraInMealList.length) {
      mismatches.push({
        nm: info.nm,
        compCount: compRegIds.length,
        mealCount: mealRegIds.length,
        missing: missingFromMealList,
        extra: extraInMealList
      });
    } else {
      okList.push(info.nm + ' (' + compRegIds.length + ' ทีม)');
    }
  });

  // ── 4) แสดงผล ──────────────────────────────────────────────
  var msg = [];
  msg.push('=== รายการที่จำนวนทีมไม่ตรงกัน ===');
  if (!mismatches.length) {
    msg.push('✅ ไม่พบ — ทุกรายการจำนวนทีมตรงกันหมด');
  } else {
    mismatches.forEach(function (m) {
      msg.push('');
      msg.push('' + m.nm);
      msg.push('   ชีทรายการแข่งขัน: ' + m.compCount + ' ทีม  |  รายชื่ออาหารรายบุคคล: ' + m.mealCount + ' ทีม');
      if (m.missing.length) {
        msg.push('   ⚠️ มีในชีทรายการ แต่ไม่โผล่ในรายชื่ออาหารเลย (' + m.missing.length + ' ทีม) — เช็คว่ากรอกชื่อสมาชิกไว้ 0 คนหรือเปล่า:');
        msg.push('      ' + m.missing.join(', '));
      }
      if (m.extra.length) {
        msg.push('   ⚠️ มีในรายชื่ออาหาร แต่หาแถวจริงในชีทรายการไม่เจอ (' + m.extra.length + ' ทีม) — อาจเป็นแถวลอย/REG_ID พิมพ์ผิด:');
        msg.push('      ' + m.extra.join(', '));
      }
    });
  }

  msg.push('');
  msg.push('=== รายการที่จำนวนทีมตรงกันปกติ ===');
  msg.push(okList.length ? okList.join('\n') : '(ไม่มี)');

  var full = msg.join('\n');
  Logger.log(full);

  var shown = full.length > 3800 ? full.slice(0, 3800) + '\n\n...(ยาวเกิน ดูฉบับเต็มที่ View > Logs)' : full;
  ui.alert(
    'ผลตรวจสอบจำนวนทีม: ชีทรายการแข่งขัน ↔ รายชื่ออาหารรายบุคคล',
    shown + '\n\nพบรายการที่มีปัญหาทั้งหมด: ' + mismatches.length + ' รายการ',
    ui.ButtonSet.OK
  );
}

// ================================================================
//  ADD-ON: แยกกรรมการตัดสินการแข่งขันเป็นก้อนตามห้อง + เช็คว่าลงทะเบียนหรือยัง
//  ต้องมีชีท "MASTER_กรรมการตัดสิน" ก่อน (นำเข้าจากไฟล์รายชื่อกรรมการเจ้าภาพร่วม)
//  คอลัมน์ที่ต้องมี: รหัสรายการ(cid) | รายการแข่งขัน | ห้องย่อย | ชื่อ-สกุล | สังกัด/สถาบัน
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  วิธีรัน: เลือกฟังก์ชัน exportJudgesByRoom แล้วกด Run
//           (หรือกดผ่านเมนู รายงานสรุป ถ้าเพิ่มบรรทัดเมนูตามด้านล่างไว้แล้ว)
//  สร้าง/อัปเดตเฉพาะชีท "กรรมการตัดสินแยกห้อง" เท่านั้น ไม่แตะ/ไม่ลบข้อมูลชีทอื่นใดๆ
// ================================================================

var SH_JUDGES_MASTER = 'MASTER_กรรมการตัดสิน';
var SH_JUDGES_BYROOM  = 'กรรมการตัดสินแยกห้อง';

// ตัดคำนำหน้า (อาจารย์/ดร./ผศ./รศ./ศ./พว./นาย/นาง/นางสาว ฯลฯ) + ช่องว่างออก
// เพื่อเทียบชื่อระหว่างไฟล์รายชื่อกรรมการ กับชื่อเต็มใน REGISTRATIONS ให้แม่นขึ้น
function _normalizeJudgeName(s) {
  s = String(s || '').replace(/[\r\n]+/g, ' ');
  var prefixPattern = /^\s*(ศาสตราจารย์|รองศาสตราจารย์|ผู้ช่วยศาสตราจารย์|ผศ\.|รศ\.|ศ\.|ดร\.|อาจารย์|อ\.|พว\.|นางสาว|นาง|นาย)\s*/;
  var changed = true;
  while (changed) {
    var before = s;
    s = s.replace(prefixPattern, '');
    changed = (s !== before);
  }
  return s.replace(/\s+/g, '').trim();
}

function exportJudgesByRoom() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  var masterSh = ss.getSheetByName(SH_JUDGES_MASTER);
  if (!masterSh || masterSh.getLastRow() < 2) {
    ui.alert(
      'ไม่พบข้อมูลในชีท "' + SH_JUDGES_MASTER + '"',
      'กรุณาสร้างชีทนี้และวางข้อมูลกรรมการก่อน\nคอลัมน์ที่ต้องมี: รหัสรายการ(cid) | รายการแข่งขัน | ห้องย่อย | ชื่อ-สกุล | สังกัด/สถาบัน',
      ui.ButtonSet.OK
    );
    return;
  }

  var mData = masterSh.getRange(2, 1, masterSh.getLastRow() - 1, 5).getValues();

  // ── เตรียมชื่อจาก REGISTRATIONS (normalize ไว้ล่วงหน้าเพื่อเทียบ) ──────
  var regNorm = []; // [{norm, regId, fullName, type}]
  var reg = ss.getSheetByName(SH_REG);
  if (reg && reg.getLastRow() > 1) {
    var rdata = reg.getDataRange().getValues();
    var rhdr = rdata[0];
    var iName = rhdr.indexOf('ชื่อเต็ม');
    var iType = rhdr.indexOf('ประเภทผู้ลงทะเบียน');
    for (var i = 1; i < rdata.length; i++) {
      var fullName = String(rdata[i][iName] || '').trim();
      if (!fullName) continue;
      var normFull = _normalizeJudgeName(fullName);
      if (!normFull || normFull.length < 2) continue; // [ADM2-12] ชื่อที่ normalize แล้วว่าง/สั้นเกิน ห้ามใช้เทียบ
      regNorm.push({
        norm: normFull,
        regId: rdata[i][0],
        fullName: fullName,
        type: String(rdata[i][iType] || '')
      });
    }
  }

  function findRegMatch(judgeName) {
    var jn = _normalizeJudgeName(judgeName);
    if (!jn || jn.length < 2) return null;
    for (var k = 0; k < regNorm.length; k++) {
      if (regNorm[k].norm === jn) return regNorm[k];
    }
    // เผื่อสะกด/เว้นวรรค/คำนำหน้าต่างกันเล็กน้อย — ลองแบบ contains (ชื่อยาวพอสมควรเท่านั้น กันจับคู่มั่ว)
    for (var m = 0; m < regNorm.length; m++) {
      var rn = regNorm[m].norm;
      if (jn.length >= 4 && rn.length >= 4 && (rn.indexOf(jn) >= 0 || jn.indexOf(rn) >= 0)) {
        return regNorm[m];
      }
    }
    return null;
  }

  // ── จัดกลุ่มตาม cid + ห้องย่อย ──────────────────────────────────
  var groups = {};
  var groupOrder = [];
  mData.forEach(function(r) {
    var cid = String(r[0] || '').trim();
    var nm = String(r[1] || '').trim();
    var subroom = String(r[2] || '').trim();
    var name = String(r[3] || '').trim();
    var org = String(r[4] || '').trim();
    if (!name) return;
    var key = cid + '|' + subroom;
    if (!groups[key]) {
      groups[key] = { nm: nm, subroom: subroom, rows: [] };
      groupOrder.push(key);
    }
    var match = findRegMatch(name);
    groups[key].rows.push([
      name, org,
      match ? 'ลงทะเบียนแล้ว ✅' : 'ยังไม่ลงทะเบียน ⚠️',
      match ? match.regId : '',
      match ? match.type : ''
    ]);
  });

  // ── เขียนลงชีท "กรรมการตัดสินแยกห้อง" เป็นก้อนๆ ต่อกัน ──────────
  var sh = ss.getSheetByName(SH_JUDGES_BYROOM);
  if (!sh) { sh = ss.insertSheet(SH_JUDGES_BYROOM); } else {
    // [ADM2-11] ยกเลิก merge เก่าก่อนล้าง
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.clearContents(); sh.clearFormats();
  }

  var totalJudges = 0, totalRegistered = 0, totalMissing = 0;
  var missingList = [];
  // [ADM2-11] เก็บทุกอย่างเป็น array แล้วเขียนทีเดียว
  var out = [], bgs = [], titleRows = [], hdrRows = [];
  var hdrLabels = ['ชื่อ-สกุล', 'สังกัด/สถาบัน', 'สถานะลงทะเบียน', 'REG_ID', 'ประเภทที่ลงทะเบียนไว้'];
  var blank5 = function() { return ['', '', '', '', '']; };
  var bgRow = function(c) { return [c, c, c, c, c]; };

  groupOrder.forEach(function(key) {
    var g = groups[key];
    var title = g.nm + (g.subroom ? ' (' + g.subroom + ')' : '');

    titleRows.push(out.length + 1);
    out.push([title, '', '', '', '']); bgs.push(bgRow('#0f1923'));
    hdrRows.push(out.length + 1);
    out.push(hdrLabels.slice()); bgs.push(bgRow('#162544'));

    g.rows.forEach(function(r) {
      out.push(r);
      if (String(r[2]).indexOf('ยังไม่') === 0) {
        bgs.push(bgRow('#fffbeb'));
        missingList.push(title + ' — ' + r[0] + ' (' + r[1] + ')');
        totalMissing++;
      } else {
        bgs.push(bgRow('#f0fdf4'));
        totalRegistered++;
      }
      totalJudges++;
    });

    // เว้น 2 แถวก่อนก้อนถัดไป
    out.push(blank5()); bgs.push(bgRow('#ffffff'));
    out.push(blank5()); bgs.push(bgRow('#ffffff'));
  });

  if (out.length) {
    sh.getRange(1, 1, out.length, 5).setValues(out);
    sh.getRange(1, 1, bgs.length, 5).setBackgrounds(bgs);
    titleRows.forEach(function(rw) {
      var tr = sh.getRange(rw, 1, 1, 5);
      tr.merge();
      sh.getRange(rw, 1).setFontWeight('bold').setFontSize(12).setFontColor('#fff');
    });
    hdrRows.forEach(function(rw) {
      sh.getRange(rw, 1, 1, 5).setFontColor('#fff').setFontWeight('bold');
    });
  }

  sh.autoResizeColumns(1, 5);

  Logger.log('=== กรรมการที่ยังไม่ลงทะเบียน (' + totalMissing + ' คน) ===\n' + missingList.join('\n'));

  ui.alert(
    'อัปเดตชีท "' + SH_JUDGES_BYROOM + '" เสร็จแล้ว',
    'กรรมการตัดสินทั้งหมด: ' + totalJudges + ' คน\n' +
    '✅ ลงทะเบียนแล้ว: ' + totalRegistered + ' คน\n' +
    '⚠️ ยังไม่ลงทะเบียน: ' + totalMissing + ' คน\n\n' +
    (totalMissing
      ? 'ดูรายชื่อที่ยังไม่ลงทะเบียนแบบเต็มได้ที่ View > Logs\n(สีเหลืองในชีทคือคนที่ยังไม่ลงทะเบียน)'
      : 'ลงทะเบียนครบทุกคนแล้ว'),
    ui.ButtonSet.OK
  );
}

// ================================================================
//  ADD-ON: รวม REGISTRATIONS + MASTER_กรรมการตัดสิน + MASTER_STAFF
//          เข้า "รายชื่ออาหารรายบุคคล" แบบ 2 ชั้น (exact match อัตโนมัติ +
//          fuzzy match ให้ตรวจก่อน) พร้อมชีท MERGE_OVERRIDE ให้แก้ไขเอง
//
//  ⚠️ ต้องมีมาก่อน: ชีท MASTER_กรรมการตัดสิน, MASTER_STAFF,
//                    ฟังก์ชัน _normalizeJudgeName, ตัวแปร SH_JUDGES_MASTER, SH_STAFF_MASTER,
//                    ฟังก์ชัน _rebuildMealListSheet (ทั้งหมดจากแอดออนก่อนหน้า — ไฟล์
//                    staff_judges_pipeline.gs ที่วางไปแล้ว ต้องยังอยู่ในไฟล์ .gs หลัก)
//
//  ลำดับการใช้งาน:
//   1) exportMealListMergedAllSources    → รวม 3 แหล่ง ด้วย exact match เท่านั้น
//   2) findFuzzyDuplicatesInMealList     → หาคู่ชื่อ "คล้ายกันมาก" มาให้ตรวจ (ไม่ตัดอะไรเอง)
//        → ไปดูชีท "รายชื่อที่อาจซ้ำ - ต้องตรวจสอบ" แล้วคัดลอกคู่ที่ยืนยันว่าใช่คนเดียวกัน
//          ไปกรอกในชีท MERGE_OVERRIDE (คอลัมน์ คำตัดสิน = "ซ้ำ" หรือ "ไม่ซ้ำ")
//   3) applyMergeOverridesAndFinalize    → รวมคู่ที่ยืนยันว่า "ซ้ำ" เป็นคนเดียว คำนวณยอดสุดท้าย
//        → คนที่มี REG_ID จริง (ลงทะเบียนแล้ว) จะไม่ถูกรวมกันเองอัตโนมัติ
//          แค่เพราะชื่อ normalize ตรงกัน — ป้องกันกรณีคนเดียวกันลงทะเบียนคนละห้อง/
//          คนละทีม (เช่น ลง SIM MAN + แนวปฏิบัติ) ถูกตัดออกผิดๆ จะรวมกลุ่มก็ต่อเมื่อ
//          มีคนยืนยัน "ซ้ำ" ใน MERGE_OVERRIDE เองเท่านั้น ส่วน MASTER (กรรมการ/STAFF
//          ที่ไม่มี REG_ID) ยังคงรวมกลุ่มอัตโนมัติเหมือนเดิม (ยกเว้นมี override "ไม่ซ้ำ")
//   4) listCrossRoomSameNamePeople        → ดูรายชื่อที่ซ้ำข้ามห้อง (ค่าเริ่มต้น = ไม่ตัด)
// ================================================================

var SH_FUZZY_REVIEW   = 'รายชื่อที่อาจซ้ำ - ต้องตรวจสอบ';
var SH_MERGE_OVERRIDE = 'MERGE_OVERRIDE';

// ── Levenshtein edit distance (ปกติ, ไม่ weight พิเศษ) ──────────────
function _editDistance(a, b) {
  var m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  var prev = new Array(n + 1);
  var cur = new Array(n + 1);
  for (var j = 0; j <= n; j++) prev[j] = j;
  for (var i = 1; i <= m; i++) {
    cur[0] = i;
    for (var k = 1; k <= n; k++) {
      var cost = (a.charAt(i - 1) === b.charAt(k - 1)) ? 0 : 1;
      cur[k] = Math.min(prev[k] + 1, cur[k - 1] + 1, prev[k - 1] + cost);
    }
    var tmp = prev; prev = cur; cur = tmp;
  }
  return prev[n];
}

// priority: 1 = ลงทะเบียนจริง (มี REG_ID) , 2 = กรรมการ (จาก MASTER, ยังไม่ลงทะเบียน) , 3 = STAFF
function _rowPriority(regId, category) {
  if (String(regId || '').indexOf(REG_PREFIX) === 0) return 1;
  if (String(category || '').indexOf('STAFF') >= 0) return 3;
  return 2; // กรรมการที่ยังไม่ลงทะเบียน หรือหมวดอื่นที่ไม่มี REG_ID
}


// ────────────────────────────────────────────────────────────────
// ขั้นที่ 1: รวม 3 แหล่ง ด้วย exact-match เท่านั้น (ปลอดภัย ตัดอัตโนมัติได้)
// ────────────────────────────────────────────────────────────────
function exportMealListMergedAllSources() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  // รีเฟรชฐาน (คนลงทะเบียนจริง) ก่อนเสมอ
  exportMealListByCompetition(true);

  var mealSh = ss.getSheetByName(SH_MEALLIST);
  if (!mealSh || mealSh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลใน "' + SH_MEALLIST + '" หลังรีเฟรช'); return; }

var keepBothPairs = _loadKeepBothOverrides(); // pairKey (ชื่อดิบ 2 ฝั่ง) → true ถ้าแอดมินยืนยันว่าไม่ใช่คนเดียวกัน

  var raw = mealSh.getRange(2, 1, mealSh.getLastRow() - 1, 7).getValues();
  var baseRows = [];
  var normSet = {};
  raw.forEach(function (r) {
    var regId = String(r[2] || '').trim();
    if (regId.indexOf(REG_PREFIX) !== 0) return;
    var dayLbl = String(r[5] || '');
    baseRows.push([r[0], r[1], r[2], r[3], r[4], r[5], r[6], dayLbl.indexOf('27') >= 0, dayLbl.indexOf('28') >= 0]);
    var norm = _normalizeJudgeName(r[4]);
    if (norm) normSet[norm] = true;
  });

var addedRows = [];
  var addedJudges = 0, addedStaff = 0;
  var skippedLog = []; // [ประเภท, ชื่อที่ถูกข้าม, ซ้ำกับชื่อ(ที่มีอยู่แล้ว), หมวดของคนที่มีอยู่แล้ว]
  var normToDisplay = {}; // norm → {name, category} ของคนที่ถูกเก็บไว้แล้วใน baseRows

  baseRows.forEach(function (r) {
    var n = _normalizeJudgeName(r[4]);
    if (n && !normToDisplay[n]) normToDisplay[n] = { name: r[4], category: r[0] };
  });

  // ── กรรมการจาก MASTER (priority 2) ─────────────────────────────
  var judgeSh = ss.getSheetByName(SH_JUDGES_MASTER);
  if (judgeSh && judgeSh.getLastRow() >= 2) {
    var jData = judgeSh.getRange(2, 1, judgeSh.getLastRow() - 1, 5).getValues();
    jData.forEach(function (r, idx) {
      var roomNm = String(r[1] || '').trim();
      var subroom = String(r[2] || '').trim();
      var name = String(r[3] || '').trim();
      var org = String(r[4] || '').trim();
      if (!name) return;
      var norm = _normalizeJudgeName(name);
      if (!norm) return;
 if (normSet[norm]) {
        var existing = normToDisplay[norm];
        var pKey = existing ? [name.trim(), existing.name.trim()].sort().join('||') : null;
        if (pKey && keepBothPairs[pKey]) {
          // แอดมินยืนยันว่าไม่ใช่คนเดียวกัน → เก็บไว้ทั้งคู่ ไม่ตัด
          addedJudges++;
          var teamNoteKB = (roomNm + (subroom ? ' (' + subroom + ')' : '')) + ' | ' + org + ' [ยืนยันคนละคนกับ ' + existing.name + ']';
          addedRows.push(['กรรมการตัดสินการแข่งขัน (ยังไม่ลงทะเบียนในระบบ)', teamNoteKB, '', idx + 1, name, '27 พ.ค. + 28 พ.ค.', '', true, true]);
          return;
        }
        skippedLog.push(['กรรมการ (MASTER)', name, existing ? existing.name : '(ไม่ทราบ)', existing ? existing.category : '']);
        return;
      }
      normSet[norm] = true;
      normToDisplay[norm] = { name: name, category: 'กรรมการตัดสินการแข่งขัน (ยังไม่ลงทะเบียนในระบบ)' };
      addedJudges++;

      var teamNote = (roomNm + (subroom ? ' (' + subroom + ')' : '')) + ' | ' + org;
      addedRows.push([
        'กรรมการตัดสินการแข่งขัน (ยังไม่ลงทะเบียนในระบบ)',
        teamNote, '', idx + 1, name, '27 พ.ค. + 28 พ.ค.', '', true, true
      ]);
    });
  }

// ── STAFF จาก MASTER (priority 3) ───────────────────────────────
  var staffSh = ss.getSheetByName(SH_STAFF_MASTER);
  if (staffSh && staffSh.getLastRow() >= 2) {
    var sData = staffSh.getRange(2, 1, staffSh.getLastRow() - 1, 3).getValues();
    sData.forEach(function (r, idx) {
      var name = String(r[0] || '').trim();
      var pos = String(r[1] || '').trim();
      if (!name) return;
      var norm = _normalizeJudgeName(name);
      if (!norm) return;
  if (normSet[norm]) {
        var existing2 = normToDisplay[norm];
        var pKey2 = existing2 ? [name.trim(), existing2.name.trim()].sort().join('||') : null;
        if (pKey2 && keepBothPairs[pKey2]) {
          addedStaff++;
          addedRows.push(['อาจารย์และเจ้าหน้าที่ (STAFF)', pos + ' [ยืนยันคนละคนกับ ' + existing2.name + ']', '', idx + 1, name, '27 พ.ค. + 28 พ.ค.', '', true, true]);
          return;
        }
        skippedLog.push(['STAFF (MASTER)', name, existing2 ? existing2.name : '(ไม่ทราบ)', existing2 ? existing2.category : '']);
        return;
      }
      normSet[norm] = true;
      normToDisplay[norm] = { name: name, category: 'อาจารย์และเจ้าหน้าที่ (STAFF)' };
      addedStaff++;
      addedRows.push([
        'อาจารย์และเจ้าหน้าที่ (STAFF)', pos, '', idx + 1, name, '27 พ.ค. + 28 พ.ค.', '', true, true
      ]);
    });
  }

var allRows = baseRows.concat(addedRows);
  _rebuildMealListSheet(mealSh, allRows, 'สรุปจำนวนอาหารตามรายการ (รวม 3 แหล่ง — exact match แล้ว)');

  // ── บันทึกรายชื่อที่ถูกตัดออกเพราะซ้ำ (exact match) ───────────
  _appendDedupeLog('ขั้นที่ 1 (exact match)', skippedLog.map(function (s) {
    return [s[0], s[1], 'ซ้ำกับ: ' + s[2] + ' (หมวด: ' + s[3] + ')'];
  }));

  ui.alert(
    'ขั้นที่ 1 เสร็จแล้ว: รวม 3 แหล่งด้วย exact match',
    'คนลงทะเบียนจริง: ' + baseRows.length + ' รายการ\n' +
    'เพิ่มกรรมการ (ยังไม่ลงทะเบียน): ' + addedJudges + ' คน\n' +
    'เพิ่ม STAFF: ' + addedStaff + ' คน\n\n' +
    'ขั้นต่อไป: รัน findFuzzyDuplicatesInMealList เพื่อหาคู่ชื่อที่อาจซ้ำแบบสะกดเพี้ยน',
    ui.ButtonSet.OK
  );
}


// ────────────────────────────────────────────────────────────────
// ขั้นที่ 2: หาคู่ชื่อ "คล้ายกันมาก" (fuzzy) มาให้ตรวจ — ไม่ตัดอะไรเองทั้งสิ้น
// ────────────────────────────────────────────────────────────────
function findFuzzyDuplicatesInMealList() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  var mealSh = ss.getSheetByName(SH_MEALLIST);
  if (!mealSh || mealSh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลใน "' + SH_MEALLIST + '" กรุณารันขั้นที่ 1 ก่อน'); return; }

  // ── เตรียม MERGE_OVERRIDE (สร้างถ้ายังไม่มี) + เก็บคู่ที่เคยตรวจแล้วไว้กันโชว์ซ้ำ ──
  var ovSh = ss.getSheetByName(SH_MERGE_OVERRIDE);
  if (!ovSh) {
    ovSh = ss.insertSheet(SH_MERGE_OVERRIDE);
    ovSh.appendRow(['ชื่อ 1', 'ชื่อ 2', 'คำตัดสิน (พิมพ์ "ซ้ำ" หรือ "ไม่ซ้ำ")', 'หมายเหตุ']);
    ovSh.getRange(1, 1, 1, 4).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
    ovSh.setFrozenRows(1);
    ovSh.autoResizeColumns(1, 4);
  }
  var reviewedPairs = {};
  if (ovSh.getLastRow() > 1) {
    ovSh.getRange(2, 1, ovSh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var n1 = _normalizeJudgeName(r[0]), n2 = _normalizeJudgeName(r[1]);
      if (n1 && n2) reviewedPairs[[n1, n2].sort().join('||')] = true;
    });
  }

  // ── อ่านรายชื่อปัจจุบันทั้งหมด ────────────────────────────────
  var raw = mealSh.getRange(2, 1, mealSh.getLastRow() - 1, 7).getValues();
  var people = [];
  raw.forEach(function (r) {
    var regId = String(r[2] || '').trim();
    if (regId.indexOf(REG_PREFIX) !== 0 && regId !== '') return;
    var name = String(r[4] || '').trim();
    var norm = _normalizeJudgeName(name);
    if (!norm) return;
    people.push({ name: name, norm: norm, category: String(r[0] || '') });
  });

  // ── bucket ตามตัวอักษรแรกของชื่อ normalize แล้ว เพื่อลดจำนวนคู่ที่ต้องเทียบ ──
  var buckets = {};
  people.forEach(function (p) {
    var k = p.norm.charAt(0);
    if (!buckets[k]) buckets[k] = [];
    buckets[k].push(p);
  });

  var candidates = [];
  var seenPairKey = {};
  Object.keys(buckets).forEach(function (k) {
    var list = buckets[k];
    for (var i = 0; i < list.length; i++) {
      for (var j = i + 1; j < list.length; j++) {
        var a = list[i], b = list[j];
        if (a.norm === b.norm) continue; // ซ้ำเป๊ะ ควรถูกตัดไปแล้วในขั้น 1 ข้าม
        var pairKey = [a.norm, b.norm].sort().join('||');
        if (seenPairKey[pairKey] || reviewedPairs[pairKey]) continue;
        seenPairKey[pairKey] = true;

        var maxLen = Math.max(a.norm.length, b.norm.length);
        if (maxLen < 4) continue; // ชื่อสั้นเกินไป เทียบแล้วไม่น่าเชื่อถือ ข้าม
        var dist = _editDistance(a.norm, b.norm);
        var threshold = maxLen >= 10 ? 3 : 2; // ชื่อยาวยอมต่างได้มากกว่านิดหน่อย
        if (dist <= threshold) {
          candidates.push({ a: a, b: b, dist: dist });
        }
      }
    }
  });

  candidates.sort(function (x, y) { return x.dist - y.dist; });

  // ── เขียนชีท "รายชื่อที่อาจซ้ำ - ต้องตรวจสอบ" ────────────────
  var sh = ss.getSheetByName(SH_FUZZY_REVIEW);
  if (!sh) { sh = ss.insertSheet(SH_FUZZY_REVIEW); } else { sh.clearContents(); sh.clearFormats(); }

  var hdr = ['ชื่อ 1', 'หมวด 1', 'ชื่อ 2', 'หมวด 2', 'ระยะห่างตัวอักษร (ยิ่งน้อยยิ่งคล้าย)', 'คำแนะนำ'];
  sh.appendRow(hdr);
  sh.getRange(1, 1, 1, hdr.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);

  var rows = candidates.map(function (c) {
    return [c.a.name, c.a.category, c.b.name, c.b.category, c.dist,
      'ถ้าใช่คนเดียวกัน → ไปกรอกในชีท "' + SH_MERGE_OVERRIDE + '" คอลัมน์คำตัดสิน = "ซ้ำ"'];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
  sh.autoResizeColumns(1, hdr.length);

  ui.alert(
    'ขั้นที่ 2 เสร็จแล้ว: หาคู่ชื่อที่อาจซ้ำ',
    'พบคู่ชื่อที่คล้ายกันมาก (และยังไม่เคยตรวจ): ' + rows.length + ' คู่\n\n' +
    (rows.length
      ? 'ไปดูชีท "' + SH_FUZZY_REVIEW + '" แล้วคัดลอกคู่ที่ยืนยันว่าใช่คนเดียวกันจริง\nไปกรอกในชีท "' + SH_MERGE_OVERRIDE + '" (คำตัดสิน = "ซ้ำ" หรือ "ไม่ซ้ำ")\nจากนั้นรัน applyMergeOverridesAndFinalize'
      : '✅ ไม่พบคู่ชื่อที่น่าสงสัยเลย ไปรัน applyMergeOverridesAndFinalize ได้เลย'),
    ui.ButtonSet.OK
  );
}


// ────────────────────────────────────────────────────────────────
// ขั้นที่ 3: รวมคู่ที่ยืนยันว่า "ซ้ำ" ใน MERGE_OVERRIDE เป็นคนเดียว แล้วสรุปยอดสุดท้าย
//   ⚠️ เวอร์ชันนี้แก้แล้ว: คนที่มี REG_ID จริง (ลงทะเบียนแล้ว) จะไม่ถูกรวมกลุ่มกันเอง
//      อัตโนมัติอีกต่อไปแค่เพราะชื่อ normalize ตรงกัน — ป้องกันกรณีคนเดียวกันลงทะเบียน
//      คนละห้อง/คนละทีม (เช่น ลง SIM MAN + แนวปฏิบัติ) ถูกตัดออกผิดๆ
//      จะรวมกลุ่มก็ต่อเมื่อมีคนยืนยัน "ซ้ำ" ใน MERGE_OVERRIDE เองเท่านั้น
//      ส่วน MASTER (กรรมการ/STAFF ที่ไม่มี REG_ID) ยังคงรวมกลุ่มอัตโนมัติเหมือนเดิม
//      (ยกเว้นมี override "ไม่ซ้ำ" ยืนยันว่าคนละคน)
// ────────────────────────────────────────────────────────────────
function applyMergeOverridesAndFinalize() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();

  var mealSh = ss.getSheetByName(SH_MEALLIST);
  var ovSh = ss.getSheetByName(SH_MERGE_OVERRIDE);
  if (!mealSh || mealSh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลใน "' + SH_MEALLIST + '"'); return; }
  if (!ovSh || ovSh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลใน "' + SH_MERGE_OVERRIDE + '" — ถ้าไม่มีคู่ที่ต้องรวม ข้ามไปดูยอดในชีทได้เลย'); return; }

  // ── สร้างกลุ่ม (union-find อย่างง่าย) จากคู่ที่คำตัดสิน = "ซ้ำ" ──────
  var parent = {};
  function find(x) { while (parent[x] && parent[x] !== x) x = parent[x]; return x; }
  function union(a, b) {
    var ra = find(a), rb = find(b);
    if (ra !== rb) parent[ra] = rb;
  }

  var mergeCount = 0;
  ovSh.getRange(2, 1, ovSh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var n1 = _normalizeJudgeName(r[0]), n2 = _normalizeJudgeName(r[1]);
    var verdict = String(r[2] || '').trim();
    // ต้องขึ้นต้นด้วยคำว่า "ซ้ำ" เป๊ะเท่านั้น — "ไม่ซ้ำ" จะมี "ซ้ำ" อยู่ตำแหน่งที่ 1 ไม่ใช่ 0 เลยไม่ผ่านเงื่อนไขนี้ (ถูกข้ามไปพอดี)
    if (!n1 || !n2 || verdict.indexOf('ซ้ำ') !== 0) return;
    if (!parent[n1]) parent[n1] = n1;
    if (!parent[n2]) parent[n2] = n2;
    union(n1, n2);
    mergeCount++;
  });

  // ── อ่านรายชื่อปัจจุบัน + ใส่ priority ─────────────────────────
  var raw = mealSh.getRange(2, 1, mealSh.getLastRow() - 1, 7).getValues();
  var people = [];
  raw.forEach(function (r) {
    var regId = String(r[2] || '').trim();
    if (regId.indexOf(REG_PREFIX) !== 0 && regId !== '') return;
    var name = String(r[4] || '').trim();
    var norm = _normalizeJudgeName(name);
    if (!norm) return;
    var dayLbl = String(r[5] || '');
    people.push({
      row: [r[0], r[1], r[2], r[3], r[4], r[5], r[6]],
      norm: norm,
      priority: _rowPriority(regId, String(r[0] || '')),
      d17: dayLbl.indexOf('27') >= 0,
      d18: dayLbl.indexOf('28') >= 0
    });
  });

  var keepBothPairs2 = _loadKeepBothOverrides();

  // ── จัดกลุ่มตาม root ──────────────────────────────────────────
  // กติกา:
  //  1. ยืนยัน "ซ้ำ" ใน MERGE_OVERRIDE ไว้ (parent map) → รวมกลุ่มเสมอ ไม่ว่ามี REG_ID หรือไม่
  //  2. มี REG_ID จริง (ลงทะเบียนแล้ว) → แยกกลุ่มตาม REG_ID เสมอ ไม่รวมกันเองอัตโนมัติ
  //     (กันกรณีคนเดียวกันลงคนละห้อง/คนละทีม ถูกตัดออกผิดๆ)
  //  3. ไม่มี REG_ID (มาจาก MASTER กรรมการ/STAFF) → รวมกลุ่มกันเองอัตโนมัติถ้าชื่อตรงกันเป๊ะ เหมือนเดิม
  //     ยกเว้นมี override "ไม่ซ้ำ" ยืนยันไว้ว่าคนละคน
  var groups = {};
  people.forEach(function (p) {
    var hasRegId = String(p.row[2] || '').indexOf(REG_PREFIX) === 0;
    var root;
    if (parent[p.norm]) {
      root = find(p.norm);
    } else if (hasRegId) {
      root = p.norm + '#reg#' + p.row[2];
    } else {
      root = p.norm;
    }
    if (groups[root] && groups[root].length && !parent[p.norm]) {
      var other = groups[root][0];
      var pKey = [p.row[4].trim(), other.row[4].trim()].sort().join('||');
      if (keepBothPairs2[pKey]) {
        root = root + '#keepapart#' + Math.random().toString(36).slice(2, 7);
      }
    }
    if (!groups[root]) groups[root] = [];
    groups[root].push(p);
  });

  var finalRows = [];
  var mergedNames = [];
  Object.keys(groups).forEach(function (root) {
    var g = groups[root];
    if (g.length === 1) {
      var p = g[0];
      finalRows.push(p.row.concat([p.d17, p.d18]));
      return;
    }
    // มากกว่า 1 คนในกลุ่ม → รวมเป็นคนเดียว เอาแถวจาก priority สูงสุด (เลขน้อยสุด) เป็นตัวหลัก
    g.sort(function (x, y) { return x.priority - y.priority; });
    var main = g[0];
    var d17 = g.some(function (p) { return p.d17; });
    var d18 = g.some(function (p) { return p.d18; });
    finalRows.push(main.row.concat([d17, d18]));
    mergedNames.push(g.map(function (p) { return p.row[4]; }).join(' = '));
  });

  _rebuildMealListSheet(mealSh, finalRows, 'สรุปจำนวนอาหารตามรายการ (สุดท้าย — รวมชื่อซ้ำตาม MERGE_OVERRIDE แล้ว)');

  // ── บันทึกรายชื่อที่ถูกรวม/ตัดออกจาก MERGE_OVERRIDE ───────────
  _appendDedupeLog('ขั้นที่ 3 (MERGE_OVERRIDE ยืนยันซ้ำ)', mergedNames.map(function (m) {
    return ['รวมชื่อซ้ำ', m, 'เก็บไว้แถวเดียว (priority: ลงทะเบียนจริง > กรรมการ > STAFF)'];
  }));

  ui.alert(
    'ขั้นที่ 3 เสร็จแล้ว: สรุปยอดสุดท้าย',
    'คู่ที่ถูกยืนยันว่า "ซ้ำ" และรวมเป็นคนเดียว: ' + mergeCount + ' คู่\n\n' +
    (mergedNames.length ? 'รายชื่อที่รวมกัน:\n' + mergedNames.join('\n') : '(ไม่มีการรวมชื่อ)') +
    '\n\n✅ จำนวนคนสุดท้ายทั้งหมด: ' + finalRows.length + ' คน\n\n' +
    'หมายเหตุ: คนที่มี REG_ID ลงทะเบียนคนละห้อง/คนละทีม จะไม่ถูกรวมอัตโนมัติแล้ว (นับแยกตามห้องตรงกับจำนวนทีมจริง)',
    ui.ButtonSet.OK
  );
}


// ────────────────────────────────────────────────────────────────
// ขั้นที่ 4: หา "ชื่อซ้ำข้ามห้อง" — คนเดียวกันที่ลงทะเบียนไว้คนละห้อง/
//           คนละรายการแข่งขัน (เช่น ลง SIM MAN ด้วย และแนวปฏิบัติด้วย)
//   ค่าเริ่มต้น = ไม่ตัด (นับแยกตามห้อง ตรงกับจำนวนทีมที่ส่งจริงในแต่ละห้อง)
//   ถ้าอยากรวมคู่ไหนจริงๆ (เช่น ลงทะเบียนซ้ำโดยไม่ตั้งใจ ไม่ใช่ลงแข่ง 2 รายการจริง)
//   ให้ไปกรอกคู่นั้นใน MERGE_OVERRIDE เป็น "ซ้ำ" เอง แล้วรัน applyMergeOverridesAndFinalize
//   สร้าง/อัปเดตเฉพาะชีท "ชื่อซ้ำข้ามห้อง - ตรวจสอบ" เท่านั้น
// ────────────────────────────────────────────────────────────────
function listCrossRoomSameNamePeople() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();
  var mealSh = ss.getSheetByName(SH_MEALLIST);
  if (!mealSh || mealSh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลใน "' + SH_MEALLIST + '"'); return; }

  var raw = mealSh.getRange(2, 1, mealSh.getLastRow() - 1, 7).getValues();
  var byNorm = {}; // norm -> [{category, team, regId, name}]
  raw.forEach(function (r) {
    var regId = String(r[2] || '').trim();
    if (regId.indexOf(REG_PREFIX) !== 0) return; // เอาเฉพาะคนที่ลงทะเบียนจริง มี REG_ID เท่านั้น
    var name = String(r[4] || '').trim();
    var norm = _normalizeJudgeName(name);
    if (!norm) return;
    if (!byNorm[norm]) byNorm[norm] = [];
    byNorm[norm].push({ category: String(r[0] || ''), team: String(r[1] || ''), regId: regId, name: name });
  });

  var crossRoom = [];
  Object.keys(byNorm).forEach(function (norm) {
    var list = byNorm[norm];
    if (list.length < 2) return;
    var distinctCats = {};
    list.forEach(function (l) { distinctCats[l.category] = true; });
    if (Object.keys(distinctCats).length < 2) return; // ห้องเดียวกันหมด (อาจเป็นซ้ำจริงในห้องเดียวกัน) ข้ามไปก่อน
    crossRoom.push(list);
  });

  var sh = ss.getSheetByName('ชื่อซ้ำข้ามห้อง - ตรวจสอบ');
  if (!sh) { sh = ss.insertSheet('ชื่อซ้ำข้ามห้อง - ตรวจสอบ'); } else { sh.clearContents(); sh.clearFormats(); }
  var hdr = ['ชื่อ', 'ห้อง/รายการ', 'ชื่อทีม/สังกัด', 'REG_ID', 'สถานะ (ค่าเริ่มต้น)'];
  sh.appendRow(hdr);
  sh.getRange(1, 1, 1, hdr.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);

  var curRow = 2;
  crossRoom.forEach(function (list) {
    list.forEach(function (l) {
      sh.getRange(curRow, 1, 1, 5).setValues([[l.name, l.category, l.team, l.regId, 'ไม่ตัด (ค่าเริ่มต้น — นับแยกตามห้อง)']]);
      sh.getRange(curRow, 1, 1, 5).setBackground('#f0fdf4');
      curRow++;
    });
    curRow++; // เว้นบรรทัดคั่นแต่ละกลุ่ม
  });
  sh.autoResizeColumns(1, 5);

  ui.alert(
    'พบชื่อซ้ำข้ามห้อง ' + crossRoom.length + ' คน',
    'ค่าเริ่มต้นคือ "ไม่ตัด" ทุกคน (นับแยกตามห้องเหมือนเดิม ตรงกับจำนวนทีมที่ส่งจริง — ' +
    'ระบบตอนนี้จะไม่รวมคนที่มี REG_ID กันเองอัตโนมัติแล้ว)\n\n' +
    (crossRoom.length
      ? 'ถ้าคนไหนจริงๆ คือลงทะเบียนซ้ำโดยไม่ตั้งใจ (ไม่ใช่ลงแข่ง 2 รายการจริง) ให้ไปกรอกคู่นั้นใน MERGE_OVERRIDE เป็น "ซ้ำ" เอง แล้วรัน applyMergeOverridesAndFinalize'
      : '✅ ไม่พบคนที่ชื่อซ้ำข้ามห้องเลย'),
    ui.ButtonSet.OK
  );
}

// ================================================================
//  FIX: 2 จุดที่ขาดหายไป — ทำให้ exportMealListMergedAllSources
//       และ applyMergeOverridesAndFinalize รันไม่ผ่าน
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก (ที่ไหนก็ได้ในไฟล์) แล้ว Save
// ================================================================

// 1) ตัวแปรที่ขาดหาย — ชื่อชีทข้อมูล STAFF
var SH_STAFF_MASTER = 'MASTER_STAFF';

// 2) ฟังก์ชันที่ขาดหาย — เขียนชีท "รายชื่ออาหารรายบุคคล" ใหม่ทั้งหมดจาก allRows ที่ส่งเข้ามา
// allRows แต่ละแถว = [nm, teamName, regId, idx, fullName, dayLbl, allergy, getsD17, getsD18]
function _rebuildMealListSheet(mealSh, allRows, summaryTitle) {
  var headers = ['รายการแข่งขัน/ประเภท', 'ชื่อทีม/สังกัด', 'REG_ID', 'ลำดับ', 'ชื่อ-นามสกุล', 'วันที่ได้รับอาหาร', 'การแพ้อาหาร'];
  mealSh.clearContents();
  mealSh.clearFormats();
  mealSh.appendRow(headers);
  mealSh.getRange(1, 1, 1, headers.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  mealSh.setFrozenRows(1);

  var dataRows = allRows.map(function (r) { return r.slice(0, 7); });
  if (dataRows.length) mealSh.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  mealSh.autoResizeColumns(1, headers.length);

  var summaryMap = {};
  var grand17 = 0, grand18 = 0, grandCount = 0;
  allRows.forEach(function (r) {
    var nm = r[0];
    var key = r[2] || r[4];
    if (!summaryMap[nm]) summaryMap[nm] = { d17: 0, d18: 0, keySet: {} };
    if (key && !summaryMap[nm].keySet[key]) summaryMap[nm].keySet[key] = true;
    if (r[7]) { summaryMap[nm].d17++; grand17++; }
    if (r[8]) { summaryMap[nm].d18++; grand18++; }
  });

  var sumStartRow = dataRows.length + 3;
  mealSh.getRange(sumStartRow, 1).setValue(summaryTitle).setFontWeight('bold').setFontSize(12);

  var sumHeaderRow = sumStartRow + 1;
  mealSh.getRange(sumHeaderRow, 1, 1, 4).setValues([['รายการ/ประเภท', 'จำนวนทีม/คน', '27 พ.ค. (คน)', '28 พ.ค. (คน)']])
    .setBackground('#162544').setFontColor('#fff').setFontWeight('bold');

  var sumKeys = Object.keys(summaryMap).sort();
  var sumRows = sumKeys.map(function (k) {
    var cnt = Object.keys(summaryMap[k].keySet).length;
    grandCount += cnt;
    return [k, cnt, summaryMap[k].d17, summaryMap[k].d18];
  });
  if (sumRows.length) mealSh.getRange(sumHeaderRow + 1, 1, sumRows.length, 4).setValues(sumRows);

  var totalRow = sumHeaderRow + 1 + sumRows.length;
  mealSh.getRange(totalRow, 1, 1, 4).setValues([['รวมทั้งหมด (นับซ้ำได้ถ้าคนเดียวอยู่หลายรายการ)', grandCount, grand17, grand18]])
    .setBackground('#f0fdf4').setFontWeight('bold');
  mealSh.autoResizeColumns(1, 4);

  return totalRow;
}

// ================================================================
//  ADD-ON: บันทึกรายชื่อซ้ำที่ถูกตัดออก ลงชีทถาวร (ดูย้อนหลังได้)
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  เรียกใช้อัตโนมัติจาก exportMealListMergedAllSources และ
//  applyMergeOverridesAndFinalize (ต้องแก้ 2 ฟังก์ชันนั้นตามคำแนะนำก่อน)
// ================================================================

var SH_DEDUPE_LOG = 'บันทึกชื่อซ้ำที่ถูกตัดออก';

// entries = [[ประเภทคนที่ถูกตัด, ชื่อที่ถูกตัดออก, รายละเอียด/เหตุผล], ...]
function _appendDedupeLog(stepLabel, entries) {
  if (!entries || !entries.length) return;
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH_DEDUPE_LOG);
  if (!sh) {
    sh = ss.insertSheet(SH_DEDUPE_LOG);
    sh.appendRow(['วันที่/เวลา', 'ขั้นตอน', 'ประเภท', 'ชื่อที่ถูกตัดออก', 'รายละเอียด']);
    sh.getRange(1, 1, 1, 5).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  var rows = entries.map(function (e) {
    return [now, stepLabel, e[0], e[1], e[2]];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  sh.autoResizeColumns(1, 5);
}

// ── ดูสรุปย้อนหลังทั้งหมดแบบเร็วๆ (ไม่ต้องเปิดชีทเอง) ──────────
function viewDedupeLog() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH_DEDUPE_LOG);
  var ui = SpreadsheetApp.getUi();
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('ยังไม่มีบันทึกชื่อซ้ำที่ถูกตัดออกเลย');
    return;
  }
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var msg = data.map(function (r) {
    return r[0] + ' | ' + r[1] + ' | ' + r[2] + ' | ตัดออก: ' + r[3] + ' | ' + r[4];
  }).join('\n');
  var shown = msg.length > 3800 ? msg.slice(0, 3800) + '\n\n...(ยาวเกิน ดูฉบับเต็มที่ชีท "' + SH_DEDUPE_LOG + '")' : msg;
  ui.alert('บันทึกชื่อซ้ำที่ถูกตัดออก (' + data.length + ' รายการ)', shown, ui.ButtonSet.OK);
}

// ================================================================
//  ADD-ON: โหลดคู่ชื่อที่แอดมินยืนยันใน MERGE_OVERRIDE ว่า "ไม่ซ้ำ"
//          (คนละคนกันจริง — อย่าตัด/อย่ารวมอัตโนมัติ)
//  วิธีติดตั้ง: copy วางต่อท้ายไฟล์ .gs หลัก
//  ใช้ร่วมกับ patch ที่แก้ exportMealListMergedAllSources และ
//  applyMergeOverridesAndFinalize ตามคำแนะนำ
// ================================================================

// คืนค่า object: pairKey (ชื่อดิบ 2 ฝั่งเรียงตามตัวอักษร คั่นด้วย '||') → true
// เฉพาะแถวที่คำตัดสินขึ้นต้นด้วย "ไม่ซ้ำ" เท่านั้น
function _loadKeepBothOverrides() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ovSh = ss.getSheetByName(SH_MERGE_OVERRIDE);
  var result = {};
  if (!ovSh || ovSh.getLastRow() < 2) return result;

  ovSh.getRange(2, 1, ovSh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var n1 = String(r[0] || '').trim();
    var n2 = String(r[1] || '').trim();
    var verdict = String(r[2] || '').trim();
    if (!n1 || !n2) return;
    // "ไม่ซ้ำ" ต้องขึ้นต้นด้วยคำนี้เป๊ะ (ตรงข้ามกับเช็ค "ซ้ำ" ในฟังก์ชันอื่น)
    if (verdict.indexOf('ไม่ซ้ำ') !== 0) return;
    var key = [n1, n2].sort().join('||');
    result[key] = true;
  });
  return result;
}

// ================================================================
//  ADD-ON: ชีทใบเสร็จเรียงตามสถาบัน/สังกัด (ดึงจาก PAYMENTS + REGISTRATIONS)
//  คอลัมน์ที่ได้: REG_ID | ชื่อ-นามสกุล | สถาบัน/สังกัด | ที่อยู่ใบเสร็จ | เลขที่ใบเสร็จ
//  เรียงตามสถาบัน/สังกัดให้คนสถาบันเดียวกันอยู่ติดกัน (แรเงาสลับสีแยกกลุ่ม)
//  อ่านเลขที่ใบเสร็จจากคอลัมน์ M ใน PAYMENTS (หา header อัตโนมัติ ถ้าหาไม่เจอ fallback เป็นคอลัมน์ M)
//  ไม่แก้ไข/ไม่ลบข้อมูลในชีทอื่นใดๆ ทั้งสิ้น สร้าง/อัปเดตเฉพาะชีท "ใบเสร็จเรียงตามสถาบัน" เท่านั้น
// ================================================================

function exportReceiptsByAddress(silent) {
  var ss  = SpreadsheetApp.openById(SS_ID);
  var ui  = SpreadsheetApp.getUi();
  var psh = ss.getSheetByName(SH_PAY);
  if (!psh || psh.getLastRow() < 2) {
    if (!silent) ui.alert('ไม่พบข้อมูลใน PAYMENTS');
    return;
  }

  var lastCol = psh.getLastColumn();
  var hdr = psh.getRange(1, 1, 1, lastCol).getValues()[0];
  var iRegId      = hdr.indexOf('REG_ID');        if (iRegId < 0) iRegId = 0;
  var iAddr       = hdr.indexOf('ที่อยู่ใบเสร็จ'); if (iAddr < 0) iAddr = 9;   // col J fallback
  var iReceiptNo  = hdr.indexOf('เลขที่ใบเสร็จ');   if (iReceiptNo < 0) iReceiptNo = 12; // col M fallback

  // ── ดึงชื่อเต็ม + สังกัดจาก REGISTRATIONS (PAYMENTS ไม่มีคอลัมน์นี้) ──
  var infoMap = {};
  var rsh = ss.getSheetByName(SH_REG);
  if (rsh && rsh.getLastRow() > 1) {
    var rdata = rsh.getDataRange().getValues();
    var rhdr  = rdata[0];
    var iN = rhdr.indexOf('ชื่อเต็ม'), iPre = rhdr.indexOf('คำนำหน้า'),
        iFn = rhdr.indexOf('ชื่อ'), iLn = rhdr.indexOf('นามสกุล'),
        iOrg = rhdr.indexOf('สังกัด/หน่วยงาน');
    rdata.slice(1).forEach(function(r) {
      var rid = String(r[0] || '').trim();
      if (!rid) return;
      var nm = String(r[iN] || '').trim();
      if (!nm) nm = ((String(r[iPre]||'')+' '+String(r[iFn]||'')+' '+String(r[iLn]||'')).trim().replace(/\s+/g,' '));
      infoMap[rid] = { name: nm, org: String(r[iOrg] || '').trim() };
    });
  }

  // ── อ่าน PAYMENTS — เก็บแถวล่าสุดต่อ REG_ID เผื่อมีการส่งสลิปซ้ำหลายครั้ง ──
  var pdata = psh.getRange(2, 1, psh.getLastRow() - 1, lastCol).getValues();
  var byRegId = {};
  pdata.forEach(function(r) {
    var rid = String(r[iRegId] || '').trim();
    if (!rid) return;
    // [ADM2-13] แถวหลังที่ว่าง ห้ามทับค่าที่ถูกต้องของแถวก่อนหน้า
    var addr = String(r[iAddr] || '').trim(), rcp = String(r[iReceiptNo] || '').trim();
    var prev = byRegId[rid];
    byRegId[rid] = {
      regId: rid,
      addr: addr || (prev && prev.addr) || '',
      receiptNo: rcp || (prev && prev.receiptNo) || ''
    };
  });

  var rows = Object.keys(byRegId).map(function(rid) {
    var p = byRegId[rid];
    var info = infoMap[rid] || {};
    return { regId: rid, name: info.name || '', org: info.org || '', addr: p.addr, receiptNo: p.receiptNo };
  }).filter(function(r) { return r.addr || r.receiptNo; }); // เอาเฉพาะที่มีข้อมูลใบเสร็จ

  // ── เรียงตามสถาบันก่อน แล้วเรียงชื่อภายในกลุ่มเดียวกัน ──
  rows.sort(function(a, b) {
    if (a.org !== b.org) return a.org.localeCompare(b.org, 'th');
    return (a.name || '').localeCompare(b.name || '', 'th');
  });

  // ── เขียนชีท ──────────────────────────────────────────────
  var shName = 'ใบเสร็จเรียงตามสถาบัน';
  var sh = ss.getSheetByName(shName);
  if (!sh) { sh = ss.insertSheet(shName); } else { sh.clearContents(); sh.clearFormats(); }

  var headers = ['REG_ID', 'ชื่อ-นามสกุล', 'สถาบัน/สังกัด', 'ที่อยู่ใบเสร็จ', 'เลขที่ใบเสร็จ'];
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);

  var dataRows = rows.map(function(r) { return [r.regId, r.name, r.org, r.addr, r.receiptNo]; });
  if (dataRows.length) sh.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

  // ── แรเงาสลับสีตามกลุ่มสถาบัน ให้เห็นง่ายว่าใครอยู่สถาบันเดียวกัน ──
  var curOrg = null, shadeOn = false, shadeBgs = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].org !== curOrg) { curOrg = rows[i].org; shadeOn = !shadeOn; }
    var shadeColor = shadeOn ? '#f0f6ff' : '#ffffff';
    shadeBgs.push(headers.map(function() { return shadeColor; }));
  }
  if (shadeBgs.length) sh.getRange(2, 1, shadeBgs.length, headers.length).setBackgrounds(shadeBgs); // [ADM2-13] batch

  sh.autoResizeColumns(1, headers.length);

  if (!silent) {
    var orgSet = {};
    rows.forEach(function(r) { orgSet[r.org] = true; });
    ui.alert(
      'สร้างชีท "' + shName + '" เสร็จแล้ว',
      'จำนวนคนทั้งหมด: ' + rows.length + ' คน\n' +
      'จำนวนสถาบัน/สังกัดที่ต่างกัน: ' + Object.keys(orgSet).length +
      (rows.length ? '' : '\n\n⚠️ ไม่พบข้อมูลใบเสร็จเลย — ตรวจสอบว่าคอลัมน์ "ที่อยู่ใบเสร็จ"/"เลขที่ใบเสร็จ" ใน PAYMENTS มีข้อมูลหรือยัง'),
      ui.ButtonSet.OK
    );
  }
}

// ================================================================
//  ADD-ON: ชีทรายชื่อผู้มีใบประกอบวิชาชีพ (รวมสมาชิกทีมทุกคน)
//  คอลัมน์ที่ได้: REG_ID | ชื่อ-นามสกุล | ใบประกอบ | เลขบัตรประชาชน
//  ดึงจาก 2 แหล่ง:
//   1) REGISTRATIONS — ผู้ลงทะเบียนหลัก/หัวหน้าทีม (คอลัมน์ เลขใบประกอบวิชาชีพ)
//   2) ชีทรายการแข่งขันทุกอัน — เฉพาะสมาชิกทีมคนที่ 2-4 (คนที่ 1 ข้ามเพราะ
//      ซ้ำกับหัวหน้าทีมใน REGISTRATIONS อยู่แล้ว กันไม่ให้คนเดียวกันขึ้นซ้ำแถว)
//  เอาเฉพาะแถวที่มีเลขใบประกอบวิชาชีพกรอกไว้จริงเท่านั้น
//  ไม่แก้ไข/ไม่ลบข้อมูลในชีทอื่นใดๆ ทั้งสิ้น สร้าง/อัปเดตเฉพาะชีท "รายชื่อผู้มีใบประกอบวิชาชีพ" เท่านั้น
// ================================================================

function exportLicensedPersonsList(silent) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ui = SpreadsheetApp.getUi();
  var rows = [];

  // ── 1) REGISTRATIONS — ผู้ลงทะเบียนหลัก/หัวหน้าทีม ──────────────
  var rsh = ss.getSheetByName(SH_REG);
  if (rsh && rsh.getLastRow() > 1) {
    var rdata = rsh.getDataRange().getValues();
    var rhdr  = rdata[0];
    var iN    = rhdr.indexOf('ชื่อเต็ม');
    var iPre  = rhdr.indexOf('คำนำหน้า'), iFn = rhdr.indexOf('ชื่อ'), iLn = rhdr.indexOf('นามสกุล');
    var iLic  = rhdr.indexOf('เลขใบประกอบวิชาชีพ');
    var iCard = rhdr.indexOf('เลขบัตรประชาชน');

    rdata.slice(1).forEach(function(r) {
      var rid = String(r[0] || '').trim();
      if (!rid) return;
      var lic = String(r[iLic] || '').trim();
      if (!lic) return; // ไม่มีใบประกอบ ข้าม

      var nm = String(r[iN] || '').trim();
      if (!nm) nm = ((String(r[iPre]||'')+' '+String(r[iFn]||'')+' '+String(r[iLn]||'')).trim().replace(/\s+/g,' '));

      rows.push([rid, nm, lic, String(r[iCard] || '').trim()]);
    });
  }

  // ── 2) ชีทรายการแข่งขันทุกอัน — เฉพาะสมาชิกทีมคนที่ 2-4 ──────────
  Object.keys(COMP_SHEET).forEach(function(cid) {
    if (cid === 'c00') return; // ออนไลน์ไม่มีทีม/ใบประกอบแยก
    var shName = COMP_SHEET[cid];
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;

    var lastCol = sh.getLastColumn();
    var headerRowIdx = -1, hdr = null;
    for (var scanRow = 1; scanRow <= Math.min(3, sh.getLastRow()); scanRow++) {
      var candidate = sh.getRange(scanRow, 1, 1, lastCol).getValues()[0];
      if (candidate.indexOf('REG_ID') >= 0) { headerRowIdx = scanRow; hdr = candidate; break; }
    }
    if (headerRowIdx < 0 || sh.getLastRow() <= headerRowIdx) return;

    var iRegId = hdr.indexOf('REG_ID');

    // เตรียมคอลัมน์เฉพาะสมาชิกที่ 2, 3, 4 เท่านั้น (ข้ามคนที่ 1 เพราะซ้ำกับ REGISTRATIONS)
    var slotCols = [];
    for (var n = 2; n <= 4; n++) {
      slotCols.push({
        pre: hdr.indexOf('คำนำหน้า ' + n),
        fn:  hdr.indexOf('ชื่อ ' + n),
        ln:  hdr.indexOf('นามสกุล ' + n),
        combined: hdr.indexOf('ชื่อ สมาชิก ' + n),
        idc: hdr.indexOf('เลขบัตร ' + n),
        lic: hdr.indexOf('ใบประกอบ ' + n)
      });
    }

    var data = sh.getRange(headerRowIdx + 1, 1, sh.getLastRow() - headerRowIdx, lastCol).getValues();
    data.forEach(function(r) {
      var regId = String(r[iRegId] || '').trim();
      if (!regId) return;

      slotCols.forEach(function(col) {
        if (col.lic < 0) return; // ชีทนี้ไม่มีคอลัมน์ใบประกอบของสมาชิกคนนี้
        var lic = String(r[col.lic] || '').trim();
        if (!lic) return; // ไม่มีใบประกอบ ข้าม

        var fullName = '';
        if (col.fn >= 0) {
          var pre = col.pre >= 0 ? r[col.pre] : '';
          var ln  = col.ln  >= 0 ? r[col.ln]  : '';
          fullName = (String(pre||'')+' '+String(r[col.fn]||'')+' '+String(ln||'')).trim();
        } else if (col.combined >= 0) {
          fullName = String(r[col.combined] || '').trim();
        }
        if (!fullName) return;

        var idc = col.idc >= 0 ? String(r[col.idc] || '').trim() : '';
        rows.push([regId, fullName, lic, idc]);
      });
    });
  });

  // ── เขียนชีท ──────────────────────────────────────────────────
  var shName = 'รายชื่อผู้มีใบประกอบวิชาชีพ';
  var sh = ss.getSheetByName(shName);
  if (!sh) { sh = ss.insertSheet(shName); } else { sh.clearContents(); sh.clearFormats(); }

  var headers = ['REG_ID', 'ชื่อ-นามสกุล', 'ใบประกอบ', 'เลขบัตรประชาชน'];
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length).setBackground('#0f1923').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);

  // เรียงตามชื่อให้ดูง่าย
  rows.sort(function(a, b) { return (a[1] || '').localeCompare(b[1] || '', 'th'); });

  if (rows.length) {
    sh.getRange(2, 3, rows.length, 2).setNumberFormat('@'); // [ADM2-17] กันเลขบัตร/ใบประกอบถูกแปลงเป็นตัวเลข
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sh.autoResizeColumns(1, headers.length);

  if (!silent) {
    ui.alert(
      'สร้างชีท "' + shName + '" เสร็จแล้ว',
      'พบผู้มีใบประกอบวิชาชีพทั้งหมด: ' + rows.length + ' คน',
      ui.ButtonSet.OK
    );
  }
}

// ── PromptPay (EMVCo Tag 29) — อัลกอริทึมเดียวกับ ppPayload ในหน้าเว็บ ──
function _promptPayPayload(id, amount){
  var n = String(id == null ? '' : id).replace(/[\s-]/g, ''), sub;
  if (/^0\d{9}$/.test(n)) sub = '0113' + '0066' + n.slice(1);
  else if (/^\d{13}$/.test(n)) sub = '0213' + n;
  else if (/^\d{15}$/.test(n)) sub = '0315' + n;
  else return null;
  function f(t, v){ return t + ('0' + v.length).slice(-2) + v; }
  var amt = Number(amount), has = isFinite(amt) && amt > 0;
  var body = '000201' + (has ? '010212' : '010211') + f('29', '0016A000000677010111' + sub) + '5802TH' + '5303764' + (has ? f('54', amt.toFixed(2)) : '') + '6304';
  var c = 0xFFFF;
  for (var i = 0; i < body.length; i++){ c ^= body.charCodeAt(i) << 8; for (var k = 0; k < 8; k++) c = (c & 0x8000) ? ((c << 1) ^ 0x1021) & 0xFFFF : (c << 1) & 0xFFFF; }
  return body + ('000' + c.toString(16).toUpperCase()).slice(-4);
}
