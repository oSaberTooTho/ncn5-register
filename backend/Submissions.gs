// ============================================================
//  Submissions.gs — ส่งผลงานเข้าประกวด (แทน Google Form ภายนอก)
//  วางเป็นไฟล์ที่ 2 ในโปรเจกต์ Apps Script (ถัดจาก Code.gs) — ทุกไฟล์ .gs ใช้ global scope เดียวกัน
//  ใช้ helper จาก Code.gs: _safe _esc _getRegInfo COMP COMP_SHEET SH_REG SH_COMP SS_ID DRIVE_FOLDER_ID CONF_NAME_*
//  Script Property (ไม่บังคับ): SUBMIT_DEADLINE = YYYY-MM-DD (Asia/Bangkok, นับรวมวันนั้น)
// ============================================================

var SH_SUBMIT = 'SUBMISSIONS';

// SIM / ตอบคำถาม / ประชุม online-onsite ไม่ต้องส่งผลงาน
var SUBMIT_COMP = {
  c05:  { kind: 'paper' },
  c06:  { kind: 'paper' },
  c07a: { kind: 'paper' },
  c07b: { kind: 'paper' },
  c10:  { kind: 'innovation' },
  c11:  { kind: 'innovation' }
};

var SUBMIT_MAX_BYTES = 10 * 1024 * 1024;
var SUBMIT_MAX_TEXT  = 2500;

var SUBMIT_MIME = {
  'application/pdf': 'pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
  'image/jpeg': 'jpg',
  'image/png': 'png'
};
var _DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
var _PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
// ชนิดไฟล์ที่รับ แยกตาม kind
var SUBMIT_MAIN_OK  = { paper: ['application/pdf', _DOCX], innovation: ['application/pdf', _DOCX, _PPTX] };
var SUBMIT_EXTRA_OK = ['application/pdf', _PPTX, 'image/jpeg', 'image/png'];

var SUBMIT_HDR = [
  'SUB_ID', 'REG_ID', 'วันที่ส่ง', 'รหัสรายการ', 'ชื่อรายการ', 'ชื่อผลงาน (TH)', 'ชื่อผลงาน (EN)',
  'ผู้นำเสนอ', 'ผู้ร่วมผลงาน', 'บทคัดย่อ/รายละเอียด', 'ลิงก์วิดีโอ', 'ไฟล์หลัก', 'ไฟล์เพิ่มเติม',
  'เวอร์ชัน', 'สถานะชำระ ณ ตอนส่ง', 'สถานะผลงาน', 'อีเมล', 'สังกัด'
];
var SUBMIT_WIDTHS = [190, 150, 140, 80, 220, 260, 220, 160, 200, 360, 200, 260, 260, 70, 120, 150, 200, 200];

function _extraHandlers() {
  return {
    getSubmissionContext: _getSubmissionContext,
    submitWork: _submitWork
  };
}

function _extraSetupSheets(ss, makeSheet, log) {
  makeSheet(SH_SUBMIT, SUBMIT_HDR, SUBMIT_WIDTHS);
}

// ── helpers ────────────────────────────────────────────────

// คืน null ถ้ายังไม่หมดเขต, ไม่งั้นข้อความ
function _submitClosedMsg() {
  var dl = _cfg('SUBMIT_DEADLINE');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dl)) return null;
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  if (today <= dl) return null;
  return 'ปิดรับผลงานแล้ว (หมดเขต ' + dl.slice(8, 10) + '/' + dl.slice(5, 7) + '/' + dl.slice(0, 4) + ')';
}

function _submitRegRow(ss, regId) {
  var sh = ss.getSheetByName(SH_REG);
  if (!sh || sh.getLastRow() < 2 || !regId) return null;
  var lc = sh.getLastColumn(), hdr = sh.getRange(1, 1, 1, lc).getValues()[0];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, lc).getValues();
  function ix(n, d) { var i = hdr.indexOf(n); return i < 0 ? d : i; }
  var iName = ix('ชื่อเต็ม', 5), iOrg = ix('สังกัด/หน่วยงาน', 11), iMail = ix('อีเมล', 6), iSt = ix('สถานะชำระ', 17);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === regId) {
      return {
        name: String(data[i][iName] || '').replace(/^'/, '').trim(),
        org: String(data[i][iOrg] || '').replace(/^'/, '').trim(),
        email: String(data[i][iMail] || '').trim(),
        payStatus: String(data[i][iSt] || '').trim()
      };
    }
  }
  return null;
}

// cid ที่ผู้ลงทะเบียนสมัครไว้ (COMPETITIONS + ชีทแยกรายการ) เฉพาะที่ต้องส่งผลงาน
function _submitEligibleCids(ss, regId) {
  var have = {};
  var csh = ss.getSheetByName(SH_COMP);
  if (csh && csh.getLastRow() >= 2) {
    var lc = csh.getLastColumn(), hdr = csh.getRange(1, 1, 1, lc).getValues()[0];
    var iC = hdr.indexOf('รหัสรายการ'); if (iC < 0) iC = 2;
    csh.getRange(2, 1, csh.getLastRow() - 1, lc).getValues().forEach(function (r) {
      if (String(r[0]).trim() === regId) have[String(r[iC]).trim()] = true;
    });
  }
  Object.keys(SUBMIT_COMP).forEach(function (cid) {
    if (have[cid]) return;
    var sh = ss.getSheetByName(COMP_SHEET[cid]);
    if (!sh || sh.getLastRow() < 3) return;
    sh.getRange(3, 2, sh.getLastRow() - 2, 1).getValues().forEach(function (r) {
      if (String(r[0]).trim() === regId) have[cid] = true;
    });
  });
  return Object.keys(SUBMIT_COMP).filter(function (cid) { return have[cid]; });
}

function _submitRows(ss, regId) {
  var sh = ss.getSheetByName(SH_SUBMIT);
  if (!sh || sh.getLastRow() < 2) return [];
  var lc = Math.max(sh.getLastColumn(), SUBMIT_HDR.length);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, lc).getValues();
  var out = [];
  data.forEach(function (r, i) { if (String(r[1]).trim() === regId) out.push({ row: i + 2, v: r }); });
  return out;
}

function _maskEmail(e) {
  e = String(e || '');
  var at = e.indexOf('@');
  if (at < 1) return '';
  return e.charAt(0) + '***' + e.slice(at);
}

function _submitDate(v) {
  try { return (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : String(v || ''); }
  catch (e) { return String(v || ''); }
}

// ── getSubmissionContext ───────────────────────────────────

function _getSubmissionContext(p) {
  try {
    var regId = String((p && p.regId) || '').trim();
    if (!regId) return { success: false, message: 'กรุณากรอก REG ID' };
    var closed = _submitClosedMsg();
    if (closed) return { success: false, closed: true, message: closed };
    var ss = SpreadsheetApp.openById(SS_ID);
    var reg = _submitRegRow(ss, regId);
    if (!reg) return { success: false, message: 'ไม่พบ REG ID นี้ในระบบ' };
    var comps = _submitEligibleCids(ss, regId).map(function (cid) {
      return { cid: cid, nm: COMP[cid].nm, kind: SUBMIT_COMP[cid].kind };
    });
    var subs = _submitRows(ss, regId).map(function (o) {
      var v = o.v, nm = String(v[4] || '');
      return {
        subId: String(v[0]), cid: String(v[3]), nm: nm, titleTh: String(v[5]), version: Number(v[13]) || 1,
        sentAt: _submitDate(v[2]), fileName: String(v[11]), status: String(v[15])
      };
    });
    return {
      success: true, regId: regId, name: reg.name, org: reg.org, email: _maskEmail(reg.email),
      payStatus: reg.payStatus, comps: comps, submissions: subs
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ── submitWork ─────────────────────────────────────────────

function _submitCheckFile(b64, mime, allowed, label) {
  mime = String(mime || '').toLowerCase();
  if (allowed.indexOf(mime) < 0) {
    var ex = allowed.map(function (m) { return SUBMIT_MIME[m].toUpperCase(); }).join(', ');
    return label + ': รองรับเฉพาะไฟล์ ' + ex;
  }
  b64 = String(b64 || '');
  if (!b64) return label + ': ไม่มีข้อมูลไฟล์';
  if (b64.length * 0.75 > SUBMIT_MAX_BYTES) return label + ': ไฟล์ใหญ่เกิน 10MB';
  return '';
}

// ตรวจ magic bytes คร่าวๆ กันไฟล์ปลอมนามสกุล (byte เป็น signed ใน Apps Script → & 255)
function _submitMagicOk(bytes, mime) {
  var b = [bytes[0] & 255, bytes[1] & 255, bytes[2] & 255, bytes[3] & 255];
  if (mime === 'application/pdf') return b[0] === 37 && b[1] === 80 && b[2] === 68 && b[3] === 70;
  if (mime === 'image/png') return b[0] === 137 && b[1] === 80 && b[2] === 78 && b[3] === 71;
  if (mime === 'image/jpeg') return b[0] === 255 && b[1] === 216;
  return b[0] === 80 && b[1] === 75; // docx / pptx = zip "PK"
}

function _submitFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function _submitFileName(regId, ver, title, ext) {
  var t = String(title || '').replace(/[\\\/:*?"<>|\r\n\t]/g, '').replace(/\s+/g, '_').slice(0, 40);
  return regId.replace(/[^\w\-]/g, '') + '_v' + ver + (t ? '_' + t : '') + '.' + ext;
}

function _submitWork(p) {
  var lock = null;
  try {
    p = p || {};
    var regId = String(p.regId || '').trim();
    var cid = String(p.cid || '').trim();
    var closed = _submitClosedMsg();
    if (closed) return { success: false, closed: true, message: closed };
    if (!regId) return { success: false, message: 'กรุณากรอก REG ID' };
    var cfg = SUBMIT_COMP[cid];
    if (!cfg || !COMP[cid]) return { success: false, message: 'รายการนี้ไม่ต้องส่งผลงานผ่านระบบ' };

    var titleTh = String(p.titleTh || '').trim().slice(0, 300);
    var titleEn = String(p.titleEn || '').trim().slice(0, 300);
    var presenter = String(p.presenter || '').trim().slice(0, 200);
    var authors = String(p.authors || '').trim().slice(0, 1500);
    var abstract = String(p.abstract || '').trim();
    var videoUrl = String(p.videoUrl || '').trim();
    if (!titleTh) return { success: false, message: 'กรุณากรอกชื่อผลงาน (ภาษาไทย)' };
    if (!presenter) return { success: false, message: 'กรุณากรอกชื่อผู้นำเสนอ' };
    if (p.consent !== true) return { success: false, message: 'กรุณาติ๊กรับรองผลงานก่อนส่ง' };
    if (!abstract) return { success: false, message: cfg.kind === 'paper' ? 'กรุณากรอกบทคัดย่อ' : 'กรุณากรอกรายละเอียดผลงาน' };
    if (abstract.length > SUBMIT_MAX_TEXT) return { success: false, message: 'บทคัดย่อ/รายละเอียดยาวเกิน ' + SUBMIT_MAX_TEXT + ' ตัวอักษร (ปัจจุบัน ' + abstract.length + ')' };
    if (cfg.kind === 'innovation' && videoUrl) {
      if (videoUrl.length > 500 || !/^https:\/\/[^\s<>"']+$/i.test(videoUrl)) return { success: false, message: 'ลิงก์วิดีโอต้องขึ้นต้นด้วย https://' };
    } else { videoUrl = ''; }

    var mainMime = String(p.fileType || '').toLowerCase();
    var err = _submitCheckFile(p.fileB64, mainMime, SUBMIT_MAIN_OK[cfg.kind], 'ไฟล์หลัก');
    if (err) return { success: false, message: err };
    var hasExtra = cfg.kind === 'innovation' && !!p.file2B64;
    var extraMime = String(p.file2Type || '').toLowerCase();
    if (hasExtra) {
      err = _submitCheckFile(p.file2B64, extraMime, SUBMIT_EXTRA_OK, 'ไฟล์เพิ่มเติม');
      if (err) return { success: false, message: err };
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var reg = _submitRegRow(ss, regId);
    if (!reg) return { success: false, message: 'ไม่พบ REG ID นี้ในระบบ' };
    if (_submitEligibleCids(ss, regId).indexOf(cid) < 0) return { success: false, message: 'REG ID นี้ไม่ได้ลงทะเบียนรายการนี้' };

    var mainBytes = Utilities.base64Decode(String(p.fileB64));
    if (mainBytes.length > SUBMIT_MAX_BYTES) return { success: false, message: 'ไฟล์หลัก: ไฟล์ใหญ่เกิน 10MB' };
    if (!_submitMagicOk(mainBytes, mainMime)) return { success: false, message: 'ไฟล์หลัก: เนื้อไฟล์ไม่ตรงกับชนิดที่ระบุ' };
    var extraBytes = null;
    if (hasExtra) {
      extraBytes = Utilities.base64Decode(String(p.file2B64));
      if (extraBytes.length > SUBMIT_MAX_BYTES) return { success: false, message: 'ไฟล์เพิ่มเติม: ไฟล์ใหญ่เกิน 10MB' };
      if (!_submitMagicOk(extraBytes, extraMime)) return { success: false, message: 'ไฟล์เพิ่มเติม: เนื้อไฟล์ไม่ตรงกับชนิดที่ระบุ' };
    }

    lock = LockService.getScriptLock();
    lock.waitLock(30000);

    var sh = _sheet(ss, SH_SUBMIT, SUBMIT_HDR);
    var prev = _submitRows(ss, regId).filter(function (o) { return String(o.v[3]).trim() === cid; });
    var ver = prev.length + 1;
    var subId = regId + '-' + cid + '-v' + ver;

    var root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var dir = _submitFolder(_submitFolder(root, 'SUBMISSIONS'), COMP_SHEET[cid]);
    var mainName = _submitFileName(regId, ver, titleTh, SUBMIT_MIME[mainMime]);
    dir.createFile(Utilities.newBlob(mainBytes, mainMime, mainName));
    var extraName = '';
    if (hasExtra) {
      extraName = _submitFileName(regId, ver, titleTh + '_extra', SUBMIT_MIME[extraMime]);
      dir.createFile(Utilities.newBlob(extraBytes, extraMime, extraName));
    }

    // เวอร์ชันก่อนหน้า → แทนที่แล้ว
    prev.forEach(function (o) { sh.getRange(o.row, 16).setValue('แทนที่แล้ว (v' + ver + ')'); });

    var now = new Date();
    var row = [
      subId, regId, now, cid, COMP[cid].nm, _safe(titleTh), _safe(titleEn), _safe(presenter), _safe(authors),
      _safe(abstract), _safe(videoUrl), mainName, extraName, ver, reg.payStatus, 'ล่าสุด', reg.email, _safe(reg.org)
    ];
    var at = sh.getLastRow() + 1;
    if (at > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), at - sh.getMaxRows());
    sh.getRange(at, 1, 1, row.length).setValues([row]);
    try { sh.getRange(at, 3).setNumberFormat('dd/MM/yyyy HH:mm'); } catch (fe) {}

    try { lock.releaseLock(); lock = null; } catch (le) {}

    var mailed = false;
    try {
      if (reg.email) { _submitSendMail(reg, regId, cid, titleTh, ver, subId, mainName, extraName, now); mailed = true; }
    } catch (me) { Logger.log('SUBMIT MAIL ERROR [' + subId + ']: ' + me.message); }

    return {
      success: true, subId: subId, version: ver, mailed: mailed, email: _maskEmail(reg.email),
      message: 'ส่งผลงานเรียบร้อย (เวอร์ชัน ' + ver + ')' + (ver > 1 ? ' — เวอร์ชันนี้แทนที่เวอร์ชันก่อนหน้า' : '')
    };
  } catch (e) {
    Logger.log('submitWork err: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

function _submitSendMail(reg, regId, cid, titleTh, ver, subId, mainName, extraName, when) {
  var url = _webAppUrl();
  var dateStr = Utilities.formatDate(when, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var td = 'padding:8px 10px;border-top:1px solid #eee;word-break:break-word;';
  var files = _esc(mainName) + (extraName ? '<br>' + _esc(extraName) : '');
  var html =
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%"><tr><td align="center"><table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;font-family:Sarabun,Arial,sans-serif;color:#1f2937;border:1px solid #e3d3a8;border-radius:8px;border-collapse:separate;overflow:hidden"><tr><td style="text-align:left">' +
    '<div style="background:#fbf6e9;padding:20px 20px;border-bottom:4px solid #b8862b">' +
      '<p style="margin:0 0 4px;font-size:9px;color:#b8862b;letter-spacing:.1em;text-transform:uppercase">Work Submission Received</p>' +
      '<h1 style="margin:0 0 3px;font-size:15px;font-weight:700;color:#0b5844">' + CONF_NAME_TH + '</h1>' +
      '<p style="margin:0 0 12px;font-size:12px;color:#0b5844">' + CONF_NAME_EN + '</p>' +
      '<div style="background:#fff;border:1px solid #e3d3a8;border-radius:6px;padding:8px 12px">' +
        '<p style="margin:0 0 2px;font-size:9px;color:#b8862b;letter-spacing:.08em;text-transform:uppercase">SUBMISSION ID</p>' +
        '<p style="margin:0;font-family:\'Courier New\',monospace;font-size:14px;font-weight:700;color:#b8862b;letter-spacing:.06em;word-break:break-all">' + _esc(subId) + '</p>' +
      '</div>' +
    '</div>' +
    '<div style="padding:20px 20px;background:#fff">' +
      '<p style="font-size:14px;margin:0 0 14px">เรียน / Dear <b>' + _esc(reg.name) + '</b>,</p>' +
      '<p style="font-size:13px;color:#555;margin:0 0 16px;line-height:1.7">ระบบได้รับผลงานของท่านเรียบร้อยแล้ว<br>Your work submission has been received.</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:14px">' +
        '<tr style="background:#f6efe1"><td style="padding:8px 10px;color:#555;width:32%">รายการ / Event</td><td style="padding:8px 10px;word-break:break-word">' + _esc(COMP[cid].nm) + '</td></tr>' +
        '<tr><td style="' + td + 'color:#555">ชื่อผลงาน / Title</td><td style="' + td + 'font-weight:700">' + _esc(titleTh) + '</td></tr>' +
        '<tr><td style="' + td + 'color:#555">เวอร์ชัน / Version</td><td style="' + td + '">v' + ver + '</td></tr>' +
        '<tr><td style="' + td + 'color:#555">ไฟล์ / Files</td><td style="' + td + 'font-family:monospace;font-size:12px">' + files + '</td></tr>' +
        '<tr><td style="' + td + 'color:#555">เวลาที่ส่ง / Sent</td><td style="' + td + '">' + dateStr + '</td></tr>' +
        '<tr><td style="' + td + 'color:#555">REG ID</td><td style="' + td + 'font-family:monospace">' + _esc(regId) + '</td></tr>' +
      '</table>' +
      '<div style="background:#fffbeb;border-left:4px solid #d97706;padding:12px 14px;margin:12px 0;border-radius:0 8px 8px 0;font-size:12px;color:#92400e;line-height:1.7">' +
        'หากส่งผลงานซ้ำในรายการเดียวกัน <b>เวอร์ชันล่าสุดจะแทนที่เวอร์ชันก่อนหน้า</b> — คณะกรรมการจะพิจารณาเฉพาะเวอร์ชันล่าสุด<br>' +
        'If you submit again for the same event, the latest version replaces the earlier ones.' +
      '</div>' +
      '<div style="text-align:center;margin-top:20px"><a href="' + url + '" style="display:inline-block;background:#0b5844;color:#fff;padding:12px 28px;text-decoration:none;font-weight:700;font-size:14px;border-radius:6px">เปิดระบบ / Open System →</a></div>' +
    '</div>' +
    '<div style="background:#f6efe1;padding:12px 20px;font-size:11px;color:#6b5a3a;border-top:1px solid #eee;text-align:center">' +
      'คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏนครปฐม<br>Faculty of Nursing, Nakhon Pathom Rajabhat University<br>อีเมลนี้ออกโดยระบบอัตโนมัติ · Auto-generated — please do not reply' +
    '</div></td></tr></table></td></tr></table>';
  GmailApp.sendEmail(reg.email,
    '[Submission Received / ได้รับผลงานแล้ว] ' + subId,
    'Dear ' + reg.name + '\nSubmission: ' + subId + '\nTitle: ' + titleTh + '\nVersion: v' + ver +
      '\nLatest version replaces earlier ones.\n' + url,
    { name: MAIL_SENDER_NAME, htmlBody: html, name: CONF_NAME_EN });
}
