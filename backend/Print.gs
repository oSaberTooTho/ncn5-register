// ============================================================
//  Print.gs — หน้าพิมพ์ A4: ใบยืนยันการลงทะเบียน + คูปองอาหาร
//  เรียกจาก doGet (?p=print&rid=<REG_ID>&k=<token>) · ลิงก์อยู่ในอีเมลยืนยัน
//  ใช้ฟังก์ชัน/ค่าคงที่จาก Code.gs (วางเป็นไฟล์สคริปต์ที่ 3 ได้เลย)
// ============================================================

// โทเคนพิมพ์ = 10 ตัวอักษร hex แรกของ HMAC-SHA256(regId, PRINT_SECRET)
function _printSecret(){
  var p=PropertiesService.getScriptProperties(), s=p.getProperty('PRINT_SECRET');
  if(!s){ s=Utilities.getUuid()+Utilities.getUuid(); p.setProperty('PRINT_SECRET',s); }
  return s;
}
function _printHmac(str,len){
  var sig=Utilities.computeHmacSha256Signature(String(str||''),_printSecret()), h='';
  for(var i=0;i<sig.length;i++){ var b=(sig[i]+256)%256; h+=(b<16?'0':'')+b.toString(16); }
  return h.slice(0,len);
}
function _printToken(regId){ return _printHmac(String(regId||'').trim(),10); }
function _printTokenOk(regId,k){
  var a=_printToken(regId), b=String(k||''), d=a.length^b.length;
  for(var i=0;i<a.length;i++) d|=a.charCodeAt(i)^(b.charCodeAt(i)||0);
  return d===0;
}
function _printUrl(regId){
  var site=_siteUrl();
  if(site) return site+'/print.html?rid='+encodeURIComponent(regId)+'&k='+_printToken(regId);
  var base=_webAppUrl()||'';
  return base+'?p=print&rid='+encodeURIComponent(regId)+'&k='+_printToken(regId);
}
// ปุ่มในอีเมล (inline style, ตาราง/ลิงก์ล้วน)
function _printButtonHtml(regId){
  try{
    var u=_esc(_printUrl(regId));
    return '<div style="text-align:center;margin:14px 0 4px"><a href="'+u+'" style="display:inline-block;background:#0b5844;color:#fff;padding:11px 26px;text-decoration:none;font-weight:700;font-size:14px;border-radius:24px">พิมพ์ใบยืนยัน / คูปอง (A4)</a></div>'+
      '<p style="text-align:center;font-size:11px;color:#6b5a3a;margin:4px 0 14px">เปิดลิงก์แล้วกดพิมพ์ เลือกกระดาษ A4</p>';
  }catch(e){ Logger.log('PRINT BTN ERROR ['+regId+']: '+e.message); return ''; }
}

function _printErrorPage(msg,hint){
  var h='<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ลิงก์ไม่ถูกต้อง</title></head>'+
    '<body style="font-family:Sarabun,Tahoma,Arial,sans-serif;background:#fbf6ec;color:#1f2937;text-align:center;padding:60px 20px"><h2 style="color:#7a1f1f">'+_esc(msg||'ลิงก์ไม่ถูกต้อง')+'</h2><p>'+_esc(hint||'กรุณาเปิดลิงก์จากอีเมลยืนยันอีกครั้ง หรือติดต่อผู้จัดงาน')+'</p></body></html>';
  return HtmlService.createHtmlOutput(h).setTitle('ลิงก์ไม่ถูกต้อง').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// ตรวจอีเมลที่ใช้ลงทะเบียน (ทางสำรองเมื่อโทเคนในลิงก์ใช้ไม่ได้ เช่น PRINT_SECRET ถูกสร้างใหม่)
function _printEmailOk(rid,em){
  em=String(em||'').trim().toLowerCase(); if(!em||!rid) return false;
  try{
    var sh=SpreadsheetApp.openById(SS_ID).getSheetByName(SH_REG); if(!sh||sh.getLastRow()<2) return false;
    var data=sh.getDataRange().getValues(), hdr=data[0], iE=hdr.indexOf('อีเมล'); if(iE<0) return false;
    for(var i=1;i<data.length;i++){ if(String(data[i][0]).trim().toLowerCase()===rid.toLowerCase()) return String(data[i][iE]||'').trim().toLowerCase()===em; }
  }catch(err){ Logger.log('_printEmailOk: '+err.message); }
  return false;
}
function _printAskEmailPage(rid){
  var base=_webAppUrl()||'';
  var h='<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ยืนยันตัวตน</title></head>'+
    '<body style="font-family:Sarabun,Tahoma,Arial,sans-serif;background:#fbf6ec;color:#1f2937;padding:40px 16px"><div style="max-width:420px;margin:0 auto;background:#fff;border:1px solid #eadfcb;border-radius:14px;padding:24px">'+
    '<h2 style="color:#0b5844;margin:0 0 6px">ยืนยันตัวตนก่อนพิมพ์</h2><p style="font-size:14px;color:#4b5563;margin:0 0 14px">ลิงก์นี้ใช้ไม่ได้แล้ว กรุณากรอกอีเมลที่ใช้ลงทะเบียน REG ID <b>'+_esc(rid)+'</b></p>'+
    '<form method="get" action="'+_esc(base)+'"><input type="hidden" name="p" value="print"><input type="hidden" name="rid" value="'+_esc(rid)+'">'+
    '<input type="email" name="em" required placeholder="อีเมลที่ใช้ลงทะเบียน" style="width:100%;box-sizing:border-box;padding:11px;border:1px solid #eadfcb;border-radius:10px;font-size:15px;margin-bottom:10px">'+
    '<button type="submit" style="width:100%;background:#0b5844;color:#fff;border:0;border-radius:24px;padding:12px;font-size:15px;font-weight:700">เปิดหน้าพิมพ์</button></form></div></body></html>';
  return HtmlService.createHtmlOutput(h).setTitle('ยืนยันตัวตน').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function _printServe(e){
  if(e.parameter&&e.parameter.mode==='batch') return _batchServe(e);
  var rid=String((e.parameter&&e.parameter.rid)||'').trim(), k=(e.parameter&&e.parameter.k)||'', em=(e.parameter&&e.parameter.em)||'';
  if(!rid) return _printErrorPage('ลิงก์ไม่ถูกต้อง');
  if(!_printTokenOk(rid,k) && !_printEmailOk(rid,em)) return em ? _printErrorPage('อีเมลไม่ตรงกับที่ลงทะเบียนไว้') : _printAskEmailPage(rid);
  var html=_printPageHtml(rid);
  if(!html) return _printErrorPage('ไม่พบข้อมูลการลงทะเบียน');
  return HtmlService.createHtmlOutput(html).setTitle('ใบยืนยันการลงทะเบียน '+rid)
    .addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── API สำหรับหน้า print.html ที่โฮสต์แยก (GitHub Pages): คืน HTML เต็มหน้าเป็น string ──
function _getPrintHtml(p){
  var rid=String((p&&p.rid)||'').trim(), k=(p&&p.k)||'', em=(p&&p.em)||'';
  if(!rid) return {success:false,message:'ลิงก์ไม่ถูกต้อง'};
  if(!_printTokenOk(rid,k) && !_printEmailOk(rid,em)) return em ? {success:false,message:'อีเมลไม่ตรงกับที่ลงทะเบียนไว้'} : {success:false,needEmail:true,message:'ลิงก์นี้ใช้ไม่ได้แล้ว กรุณากรอกอีเมลที่ใช้ลงทะเบียน'};
  var html=_printPageHtml(rid);
  if(!html) return {success:false,message:'ไม่พบข้อมูลการลงทะเบียน'};
  return {success:true,html:html};
}
function _getBatchHtml(p){
  p=p||{};
  var day=String(p.day||''), cid=String(p.cid||''), paid=String(p.paid||''), exp=String(p.exp||''), k=String(p.k||'');
  var bad={success:false,message:'ลิงก์ไม่ถูกต้อง กรุณาสร้างลิงก์ใหม่จากหน้าเจ้าหน้าที่'};
  if(!_batchParamsOk(day,cid,paid)||!/^\d{9,12}$/.test(exp)) return bad;
  var a=_batchSig(day,cid,paid,exp), dd=a.length^k.length;
  for(var i=0;i<a.length;i++) dd|=a.charCodeAt(i)^(k.charCodeAt(i)||0);
  if(dd!==0) return bad;
  if(Number(exp)<Math.floor(Date.now()/1000)) return {success:false,message:'ลิงก์หมดอายุ — ลิงก์พิมพ์รวมมีอายุ 2 ชั่วโมง กรุณาสร้างลิงก์ใหม่จากหน้าเจ้าหน้าที่'};
  return {success:true,html:_batchHtml(_batchBuild(day,cid,paid==='1'),day,cid,paid==='1')};
}

function _printCss(){
  return '@page{size:A4;margin:14mm}'+
  '*{box-sizing:border-box}html,body{margin:0;padding:0}'+
  'body{font-family:Sarabun,Prompt,Tahoma,Arial,sans-serif;font-size:13pt;color:#1f2937;background:#e5e7eb;-webkit-print-color-adjust:exact;print-color-adjust:exact}'+
  '.bar{position:sticky;top:0;z-index:5;background:#0b5844;color:#fff;padding:10px 16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;font-size:11pt}'+
  '.bar button{background:#b8862b;color:#fff;border:0;border-radius:20px;padding:8px 22px;font:700 12pt Sarabun,Tahoma,sans-serif;cursor:pointer}'+
  '.sheets{padding:16px 0}'+
  '.page{width:210mm;min-height:297mm;margin:0 auto 14px;background:#fff;padding:14mm;box-shadow:0 2px 12px rgba(0,0,0,.18);page-break-after:always;break-after:page}'+
  '.page:last-child{page-break-after:auto;break-after:auto}'+
  '.hd{border-bottom:3px solid #b8862b;padding-bottom:4mm;margin-bottom:5mm}'+
  '.hd .k{font-size:9pt;color:#b8862b;letter-spacing:.12em;text-transform:uppercase}'+
  '.hd h1{margin:1mm 0;font-size:14pt;color:#0b5844;line-height:1.3}.hd .en{font-size:11pt;color:#0b5844}.hd .dv{font-size:10pt;color:#6b5a3a;margin-top:1mm}'+
  '.rid{border:2px solid #b8862b;background:#fbf6ec;border-radius:3mm;padding:3mm 5mm;margin:0 0 5mm;text-align:center}'+
  '.rid small{display:block;font-size:9pt;color:#b8862b;letter-spacing:.12em}.rid b{font:700 22pt \'Courier New\',monospace;color:#0b5844;letter-spacing:.08em;word-break:break-all}'+
  'h2{font-size:13pt;color:#0b5844;margin:5mm 0 2mm;border-left:4px solid #b8862b;padding-left:3mm}'+
  'table.t{width:100%;border-collapse:collapse;font-size:12pt}table.t td,table.t th{padding:2mm 3mm;border:1px solid #e3d3a8;text-align:left;vertical-align:top}'+
  'table.t th{background:#f6efe1;color:#555;width:32%;font-weight:600}'+
  '.badge{display:inline-block;padding:0 3mm;border-radius:10mm;font-weight:700;font-size:11pt}.ok{background:#dcfce7;color:#166534}.wait{background:#fef3c7;color:#92400e}.bad{background:#fee2e2;color:#991b1b}'+
  '.sch{font-size:10pt;color:#6b7280}'+
  '.note{border-left:4px solid #b8862b;background:#fff8e1;padding:3mm 4mm;margin:4mm 0;font-size:11pt;color:#6b5a3a}'+
  '.ft{margin-top:8mm;padding-top:3mm;border-top:1px solid #e3d3a8;font-size:9pt;color:#6b7280;text-align:center}'+
  '.dh{display:flex;justify-content:space-between;align-items:center;background:#f3e6c4;border-bottom:3px solid #b8862b;border-radius:2mm 2mm 0 0;padding:3mm 4mm;margin-bottom:5mm}'+
  '.dh b{font-size:14pt;color:#0b5844}.dh span{background:#0b5844;color:#fff;border-radius:10mm;padding:0 4mm;font-weight:700;font-size:11pt}'+
  '.grid{display:grid;grid-template-columns:1fr 1fr;gap:6mm}'+
  '.card{break-inside:avoid;page-break-inside:avoid;display:flex;height:48mm;border:0.5mm dashed #b8862b;border-radius:3mm;background:#fbf6ec;overflow:hidden;position:relative}'+
  '.card .stub{width:24mm;background:#0b5844;color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;flex:none}'+
  '.card .stub i{font-style:normal;font-size:7pt;letter-spacing:.2em;color:#e8cc7a;font-weight:700}.card .stub s{text-decoration:none;font-size:26pt;font-weight:700;line-height:1.05}.card .stub em{font-style:normal;font-size:8pt;color:#f3e6c4}'+
  '.card .bd{padding:3mm 3.5mm;flex:1;min-width:0;display:flex;flex-direction:column;gap:1.2mm}'+
  '.card .no{align-self:flex-start;background:#f3e6c4;border:1px solid #b8862b;color:#7a5a12;font-size:8.5pt;font-weight:700;padding:0 2.5mm;border-radius:10mm}'+
  '.card .nm{font-size:11.5pt;font-weight:700;line-height:1.25;overflow:hidden;max-height:2.6em}.card .cp{font-size:8.5pt;color:#6b7280;line-height:1.2;max-height:2.4em;overflow:hidden}'+
  '.card .id{font:700 9pt \'Courier New\',monospace;color:#0b5844;background:#f6efe1;border:1px solid #b8862b;border-radius:1.5mm;padding:0 1.5mm;word-break:break-all;margin-top:auto}'+
  '.card .al{font-size:8.5pt;font-weight:700;color:#b91c1c}'+
  '.card .cut{position:absolute;right:2mm;top:1mm;font-size:9pt;color:#c9b887}'+
  '@media screen and (max-width:230mm){.page{width:100%;min-height:0;padding:6mm}.grid{grid-template-columns:1fr}.card{height:auto;min-height:48mm}}'+
  '@media print{body{background:#fff;font-size:13pt}.bar{display:none!important}.sheets{padding:0}.page{width:auto;min-height:0;margin:0;padding:0;box-shadow:none}}';
}

function _printData(regId){
  var ss=SpreadsheetApp.openById(SS_ID), sh=ss.getSheetByName(SH_REG);
  if(!sh||sh.getLastRow()<2) return null;
  var v=sh.getDataRange().getValues(), h=v[0], id=String(regId).trim(), row=null;
  for(var i=1;i<v.length;i++) if(String(v[i][0]).trim()===id){ row=v[i]; break; }
  if(!row) return null;
  var g=function(n){var x=h.indexOf(n); return x<0?'':row[x];};
  var o={rid:id,name:String(g('ชื่อเต็ม')||'').trim(),org:String(g('สังกัด/หน่วยงาน')||''),type:String(g('ประเภทผู้ลงทะเบียน')||''),
    email:String(g('อีเมล')||''),fee:Number(g('ยอดรวม')||0),status:String(g('สถานะชำระ')||''),
    n17:Number(g('อาหาร_27พค')||0),n18:Number(g('อาหาร_28พค')||0),comps:[]};
  if(!o.name) o.name=((g('คำนำหน้า')||'')+' '+(g('ชื่อ')||'')+' '+(g('นามสกุล')||'')).replace(/\s+/g,' ').trim();
  try{
    var cs=ss.getSheetByName(SH_COMP);
    if(cs&&cs.getLastRow()>=2){
      var cv=cs.getDataRange().getValues();
      for(var j=1;j<cv.length;j++){
        if(String(cv[j][0]).trim()!==id) continue;
        var cid=String(cv[j][2]||'').trim(), c=COMP[cid]; if(!c) continue;
        o.comps.push({nm:c.nm,sch:_schedLine(cid)});
      }
    }
  }catch(e){ Logger.log('PRINT COMP ERROR ['+id+']: '+e.message); }
  o.coupons={d17:[],d18:[]}; o.isFree=false;
  try{
    var b=_couponListFromSheets(ss,id,o.name); o.isFree=!!b.isFree;
    o.coupons=_mealCoupons(b.compList,b.isFree,o.name);
  }catch(e2){ Logger.log('PRINT COUPON ERROR ['+id+']: '+e2.message); }
  o.paid=(o.status==='ชำระแล้ว')||o.isFree;
  return o;
}

function _printCard(x,day,rid){
  var dp=day.split(' ');
  return '<div class="card"><div class="stub"><i>LUNCH</i><s>'+_esc(dp[0])+'</s><em>'+_esc(dp[1]||'')+' 2570</em></div>'+
    '<div class="bd"><span class="no">คูปอง '+x.n+'/'+x.total+'</span><div class="nm">'+_esc(x.name)+'</div>'+
    (x.comp?'<div class="cp">'+_esc(x.comp)+'</div>':'')+
    (x.allergy?'<div class="al">⚠ แพ้อาหาร: '+_esc(x.allergy)+'</div>':'')+
    '<div class="id">'+_esc(rid)+'</div></div><span class="cut">✂</span></div>';
}
function _printDayPages(arr,day,rid){
  var out='', per=8;
  for(var s=0;s<arr.length;s+=per){
    var part=arr.slice(s,s+per);
    out+='<section class="page"><div class="dh"><b>คูปองอาหารกลางวัน · '+day+' 2570'+(s>0?' (ต่อ)':'')+'</b><span>'+arr.length+' ที่</span></div><div class="grid">'+
      part.map(function(x){return _printCard(x,day,rid);}).join('')+'</div>'+
      '<div class="note" style="margin-top:6mm">1 คูปอง = อาหาร 1 ที่ · ใช้ได้เฉพาะวันที่ระบุ · ตัดตามเส้นประแล้วแสดงที่จุดรับอาหาร</div></section>';
  }
  return out;
}

function _printPageHtml(regId,opts){
  opts=opts||{}; var pdf=!!opts.pdf;
  var d=opts.data||_printData(regId); if(!d) return '';
  var now=Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm');
  var stCls=d.paid?'ok':(d.status==='ปฏิเสธ'||d.status==='ยกเลิก'?'bad':'wait');
  var stTxt=d.isFree&&!d.status?'ไม่มีค่าธรรมเนียม':(d.status||'รอชำระเงิน');
  var evs=d.comps.length?d.comps.map(function(c){return '<tr><td>'+_esc(c.nm)+(c.sch?'<div class="sch">'+_esc(c.sch)+'</div>':'')+'</td></tr>';}).join(''):'';
  var c17=d.coupons.d17.length||d.n17, c18=d.coupons.d18.length||d.n18;
  var p1='<section class="page"><div class="hd"><div class="k">Registration Confirmation</div><h1>'+_esc(CONF_NAME_TH)+'</h1><div class="en">'+_esc(CONF_NAME_EN)+'</div>'+
    '<div class="dv">'+_esc(CONF_DATE_TH)+' · '+_esc(CONF_VENUE_TH)+'</div></div>'+
    '<h2 style="margin-top:0">ใบยืนยันการลงทะเบียน</h2>'+
    '<div class="rid"><small>REG ID</small><b>'+_esc(d.rid)+'</b></div>'+
    '<table class="t"><tr><th>ชื่อ-นามสกุล</th><td><b>'+_esc(d.name)+'</b></td></tr>'+
    '<tr><th>สังกัด/หน่วยงาน</th><td>'+_esc(d.org||'-')+'</td></tr>'+
    '<tr><th>ประเภทผู้ลงทะเบียน</th><td>'+_esc(d.type||'-')+'</td></tr>'+
    '<tr><th>อีเมล</th><td>'+_esc(d.email||'-')+'</td></tr>'+
    '<tr><th>ค่าลงทะเบียน</th><td>'+(d.isFree&&!d.fee?'ไม่มีค่าธรรมเนียม':d.fee.toLocaleString()+' บาท')+'</td></tr>'+
    '<tr><th>สถานะ</th><td><span class="badge '+stCls+'">'+_esc(stTxt)+'</span></td></tr></table>'+
    (evs?'<h2>รายการที่ลงทะเบียน / กำหนดการ</h2><table class="t">'+evs.replace(/<td>/g,'<td style="width:auto">')+'</table>':'')+
    '<h2>สิทธิ์อาหารกลางวัน</h2><table class="t"><tr><th>'+MD17+' 2570</th><td>'+(c17?c17+' ที่':'ไม่มีสิทธิ์')+'</td></tr><tr><th>'+MD18+' 2570</th><td>'+(c18?c18+' ที่':'ไม่มีสิทธิ์')+'</td></tr></table>'+
    (d.paid?'':'<div class="note">คูปองจะพิมพ์ได้หลังเจ้าหน้าที่ยืนยันการชำระเงิน</div>')+
    '<div class="ft">พิมพ์เมื่อ '+now+' · เอกสารนี้ออกโดยระบบอัตโนมัติ</div></section>';
  var dp=pdf?_printDayPagesPdf:_printDayPages;
  var days=d.paid?dp(d.coupons.d17,MD17,d.rid)+dp(d.coupons.d18,MD18,d.rid):'';
  if(pdf){
    return '<!doctype html><html lang="th"><head><meta charset="utf-8"><title>INHI-'+_esc(d.rid)+'</title>'+
      '<style>'+_printCssPdf()+'</style></head><body>'+p1+days+'</body></html>';
  }
  return '<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<title>ใบยืนยันการลงทะเบียน '+_esc(d.rid)+'</title>'+
    '<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&family=Prompt:wght@500;700&display=swap" rel="stylesheet">'+
    '<style>'+_printCss()+'</style></head><body>'+
    '<div class="bar"><button type="button" onclick="window.print()">พิมพ์</button><span>เลือกขนาดกระดาษ A4 · พิมพ์สี</span></div>'+
    '<div class="sheets">'+p1+days+'</div></body></html>';
}

// ── PDF แนบอีเมล ────────────────────────────────────────────
// CSS สำหรับตัวแปลง HTML→PDF ของ Drive: ตาราง/บล็อกล้วน ไม่ใช้ grid/flex, ไม่ดึงฟอนต์ภายนอก
function _printCssPdf(){
  return '@page{size:A4;margin:14mm}'+
  'body{font-family:\'Sarabun\',\'TH Sarabun New\',\'Noto Sans Thai\',Tahoma,sans-serif;font-size:12pt;color:#1f2937;margin:0;padding:0}'+
  '.page{page-break-after:always}.page.last{page-break-after:auto}'+
  '.hd{border-bottom:3px solid #b8862b;padding-bottom:4mm;margin-bottom:5mm}'+
  '.hd .k{font-size:9pt;color:#b8862b}.hd h1{margin:1mm 0;font-size:14pt;color:#0b5844}.hd .en{font-size:11pt;color:#0b5844}.hd .dv{font-size:10pt;color:#6b5a3a}'+
  '.rid{border:2px solid #b8862b;background:#fbf6ec;padding:3mm 5mm;margin:0 0 5mm 0;text-align:center}'+
  '.rid small{display:block;font-size:9pt;color:#b8862b}.rid b{font-family:\'Courier New\',monospace;font-size:20pt;color:#0b5844}'+
  'h2{font-size:13pt;color:#0b5844;margin:5mm 0 2mm 0;border-left:4px solid #b8862b;padding-left:3mm}'+
  'table.t{width:182mm;border-collapse:collapse;font-size:12pt}table.t td,table.t th{padding:2mm 3mm;border:1px solid #e3d3a8;text-align:left;vertical-align:top}'+
  'table.t th{background:#f6efe1;color:#555;width:55mm}'+
  '.badge{font-weight:bold}.ok{color:#166534}.wait{color:#92400e}.bad{color:#991b1b}'+
  '.sch{font-size:10pt;color:#6b7280}'+
  '.note{border-left:4px solid #b8862b;background:#fff8e1;padding:3mm 4mm;margin:4mm 0;font-size:11pt;color:#6b5a3a}'+
  '.ft{margin-top:8mm;padding-top:3mm;border-top:1px solid #e3d3a8;font-size:9pt;color:#6b7280;text-align:center}'+
  'table.dh{width:182mm;border-collapse:collapse;background:#f3e6c4;border-bottom:3px solid #b8862b;margin-bottom:5mm}table.dh td{padding:3mm 4mm}'+
  'table.dh b{font-size:14pt;color:#0b5844}table.dh .n{text-align:right;font-weight:bold;color:#0b5844}'+
  'table.grid{width:182mm;border-collapse:separate;border-spacing:0}table.grid td.gc{width:88mm;padding:0 3mm 6mm 0;vertical-align:top}table.grid td.gc.r{padding:0 0 6mm 3mm}'+
  'table.card{width:88mm;height:46mm;border:1px dashed #b8862b;background:#fbf6ec;border-collapse:collapse;page-break-inside:avoid}'+
  'table.card td.stub{width:22mm;background:#0b5844;color:#ffffff;text-align:center;vertical-align:middle}'+
  'td.stub i{font-style:normal;font-size:7pt;color:#e8cc7a;display:block}td.stub s{text-decoration:none;font-size:24pt;font-weight:bold;display:block}td.stub em{font-style:normal;font-size:8pt;color:#f3e6c4;display:block}'+
  'table.card td.bd{padding:3mm;vertical-align:top}'+
  '.no{font-size:8.5pt;font-weight:bold;color:#7a5a12}.nm{font-size:11.5pt;font-weight:bold;margin-top:1mm}.cp{font-size:8.5pt;color:#6b7280}'+
  '.id{font-family:\'Courier New\',monospace;font-size:9pt;font-weight:bold;color:#0b5844;margin-top:2mm}.al{font-size:8.5pt;font-weight:bold;color:#b91c1c}';
}
function _printCardPdf(x,day,rid){
  var dp=day.split(' ');
  return '<table class="card"><tr><td class="stub"><i>LUNCH</i><s>'+_esc(dp[0])+'</s><em>'+_esc(dp[1]||'')+' 2570</em></td>'+
    '<td class="bd"><div class="no">คูปอง '+x.n+'/'+x.total+'</div><div class="nm">'+_esc(x.name)+'</div>'+
    (x.comp?'<div class="cp">'+_esc(x.comp)+'</div>':'')+
    (x.allergy?'<div class="al">แพ้อาหาร: '+_esc(x.allergy)+'</div>':'')+
    '<div class="id">'+_esc(rid)+'</div></td></tr></table>';
}
function _printDayPagesPdf(arr,day,rid){
  var out='', per=8;
  for(var s=0;s<arr.length;s+=per){
    var part=arr.slice(s,s+per), rows='';
    for(var i=0;i<part.length;i+=2){
      rows+='<tr><td class="gc">'+_printCardPdf(part[i],day,rid)+'</td><td class="gc r">'+(part[i+1]?_printCardPdf(part[i+1],day,rid):'')+'</td></tr>';
    }
    out+='<div class="page"><table class="dh"><tr><td><b>คูปองอาหารกลางวัน · '+day+' 2570'+(s>0?' (ต่อ)':'')+'</b></td><td class="n">'+arr.length+' ที่</td></tr></table>'+
      '<table class="grid">'+rows+'</table>'+
      '<div class="note">1 คูปอง = อาหาร 1 ที่ · ใช้ได้เฉพาะวันที่ระบุ · ตัดตามเส้นประแล้วแสดงที่จุดรับอาหาร</div></div>';
  }
  return out;
}

// สร้าง PDF A4 (ใบยืนยัน [+ คูปองถ้าชำระแล้ว/ฟรี]) — คืน Blob หรือ null (ล้มเหลวได้ อีเมลยังส่งตามปกติ)
// เรียกเฉพาะในคิวอีเมล/ส่งซ้ำ ห้ามเรียกบน request path
function _printPdfBlob(regId){
  try{
    var d=_printData(regId); if(!d) return null;
    var html=_printPageHtml(regId,{pdf:true,data:d}); if(!html) return null;
    var name='INHI-'+regId+(d.paid?'-คูปอง':'-ใบยืนยัน')+'.pdf';
    var pdf=Utilities.newBlob(html,'text/html','INHI-'+regId+'.html').getAs('application/pdf').setName(name);
    _printSavePdf(pdf,regId);
    return pdf;
  }catch(e){ Logger.log('PRINT PDF ERROR ['+regId+']: '+e.message); return null; }
}
// เก็บสำเนาใน DRIVE_FOLDER_ID/PDF/<regId>.pdf (ทับของเดิม)
function _printSavePdf(pdf,regId){
  try{
    if(!DRIVE_FOLDER_ID) return;
    var root=DriveApp.getFolderById(DRIVE_FOLDER_ID), it=root.getFoldersByName('PDF'), dir=it.hasNext()?it.next():root.createFolder('PDF');
    var nm=regId+'.pdf', old=dir.getFilesByName(nm);
    while(old.hasNext()){ try{ old.next().setTrashed(true); }catch(e1){} }
    dir.createFile(Utilities.newBlob(pdf.getBytes(),'application/pdf',nm));
  }catch(e){ Logger.log('PRINT PDF SAVE ERROR ['+regId+']: '+e.message); }
}
// ข้อความบรรทัดในอีเมลเมื่อแนบ PDF
function _printAttachNote(){
  return '<p style="text-align:center;font-size:12px;color:#0b5844;font-weight:700;margin:6px 0 2px">แนบไฟล์ PDF สำหรับพิมพ์ (A4) มาในอีเมลนี้แล้ว</p>';
}


// ============================================================
//  พิมพ์คูปองรวมตาม session (เจ้าหน้าที่) — ?p=print&mode=batch&day=&cid=&paid=&exp=&k=
// ============================================================
var BATCH_DAYS={'17':MD17+' 2570','18':MD18+' 2570','both':MD17+'–'+MD18+' 2570'};
function _batchSig(day,cid,paid,exp){ return _printHmac(['batch',day,cid,paid,exp].join('|'),12); }
function _batchCidLabel(cid){ return cid==='all'?'ทุกรายการ':cid==='committee'?'กรรมการ':(COMP[cid]?COMP[cid].nm:cid); }
function _batchParamsOk(day,cid,paid){
  return (day==='17'||day==='18'||day==='both') && (paid==='0'||paid==='1') &&
    (cid==='all'||cid==='committee'||(COMP[cid]&&cid!=='c00'));
}
// สร้างรายการคูปองที่จะพิมพ์: [{day:'17'|'18',dayLbl,seq,n,total,name,comp,allergy,rid,org,unpaid}] เรียง comp → ชื่อผู้ลงทะเบียน → n
// ── ดัชนีในหน่วยความจำสำหรับ batch: อ่านแต่ละชีทครั้งเดียวต่อคำขอ ──
// COMPETITIONS → {regId:[cid,...]} (ลำดับ/กติกาเดียวกับ _regCids; null = ไม่มีแถว → วนทุกชีทเหมือนเดิม)
function _batchCompIndex(ss){
  var sh=ss.getSheetByName(SH_COMP);
  if(!sh||sh.getLastRow()<2) return null;
  var v=sh.getRange(1,1,sh.getLastRow(),Math.max(3,Math.min(sh.getLastColumn(),4))).getValues();
  var iR=v[0].indexOf('REG_ID'); if(iR<0) iR=0;
  var iC=v[0].indexOf('รหัสรายการ'); if(iC<0) iC=2;
  var out={};
  for(var i=1;i<v.length;i++){
    var id=String(v[i][iR]).trim(), c=String(v[i][iC]||'').trim();
    if(c&&COMP_SHEET[c]){ var a=out[id]||(out[id]=[]); if(a.indexOf(c)<0) a.push(c); }
  }
  return out;
}
// ชีทรายการ 1 ชีท → {regId:{days,members}} (แถวแรกที่ตรงต่อ regId; กติกา header/สมาชิกเดียวกับ _couponListFromSheets); wantIds=null → ทุก regId
function _batchMemberIndex(ss,cid,wantIds){
  var sh=ss.getSheetByName(COMP_SHEET[cid]); if(!sh||sh.getLastRow()<2) return {};
  var all=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues(), hr=-1;
  for(var t=0;t<Math.min(3,all.length)&&hr<0;t++) if(all[t].indexOf('REG_ID')>=0) hr=t;
  if(hr<0||all.length<=hr+1) return {};
  var H=all[hr], iR=H.indexOf('REG_ID'), iD=H.indexOf('จำนวนวัน'), cols=[], out={};
  for(var n=1;n<=4;n++){
    var iFn=H.indexOf('ชื่อ '+n); if(iFn<0) iFn=H.indexOf('ชื่อ สมาชิก '+n);
    cols.push([H.indexOf('คำนำหน้า '+n),iFn,H.indexOf('นามสกุล '+n),H.indexOf('แพ้อาหาร '+n)]);
  }
  for(var k=hr+1;k<all.length;k++){
    var row=all[k], id=String(row[iR]).trim();
    if(out[id]||(wantIds&&!wantIds[id])) continue;
    var g=function(x){return x<0?'':String(row[x]||'').trim();}, mem=[];
    for(var m=0;m<4;m++){
      var c=cols[m];
      mem.push({name:g(c[1])?(g(c[0])+' '+g(c[1])+' '+g(c[2])).trim():'',allergy:g(c[3])});
    }
    out[id]={days:Number(row[iD])===2?2:1,members:mem};
  }
  return out;
}
// เหมือน _couponListFromSheets แต่ใช้ดัชนีที่อ่านไว้แล้ว (ไม่แตะชีท)
// compRowsByReg=ผล _batchCompIndex (หรือ null) · memberRowsByReg={cid:{regId:{days,members}}}
function _couponListFromRows(regId,fullName,isFree,compRowsByReg,memberRowsByReg,al0){
  var id=String(regId||'').trim();
  if(isFree) return {isFree:true,compList:[{cid:'',days:1,members:[{name:'',allergy:al0||''}]}]};
  var myCids=compRowsByReg?(compRowsByReg[id]||null):null, list=[];
  Object.keys(COMP_SHEET).forEach(function(cid){
    if(cid==='c00') return;
    if(myCids&&myCids.indexOf(cid)<0) return;
    var r=(memberRowsByReg[cid]||{})[id];
    if(r) list.push({cid:cid,days:r.days,members:r.members});
  });
  return {isFree:false,compList:list};
}
function _batchBuild(day,cid,onlyPaid){
  var ss=SpreadsheetApp.openById(SS_ID), sh=ss.getSheetByName(SH_REG), items=[], regs={};
  if(!sh||sh.getLastRow()<2) return {items:items,regs:regs};
  var v=sh.getDataRange().getValues(), h=v[0];
  var iF=h.indexOf('ฟรี/ชำระ'), iS=h.indexOf('สถานะชำระ'), iN=h.indexOf('ชื่อเต็ม'), iO=h.indexOf('สังกัด/หน่วยงาน'), iT=h.indexOf('ประเภทผู้ลงทะเบียน'), iA=h.indexOf('การแพ้อาหาร');
  var wantComp=(cid==='all'||cid==='committee')?'':COMP[cid].nm;
  var compIdx=_batchCompIndex(ss), nmCid={};
  Object.keys(COMP).forEach(function(c){ nmCid[COMP[c].nm]=c; });
  // แถวแรกต่อ REG_ID → isFree/แพ้อาหาร (เหมือน _couponListFromSheets)
  var info={}, i, rid0;
  for(i=1;i<v.length;i++){
    rid0=String(v[i][0]||'').trim(); if(!rid0||info[rid0]) continue;
    info[rid0]={free:String(v[i][iF]||'')==='ฟรี'||[TYPE_LBL.committee_dir,TYPE_LBL.committee_judge].indexOf(String(v[i][iT]||''))>=0, al:String(v[i][iA]||'')};
  }
  // ผู้ลงทะเบียนที่ผ่านตัวกรองสถานะ → ชีทที่ต้องอ่าน (อ่านครั้งเดียวต่อชีท)
  var cand={}, nCand=0, need={};
  for(i=1;i<v.length;i++){
    var r0=String(v[i][0]||'').trim(); if(!r0) continue;
    var st0=String(v[i][iS]||''), fr0=String(v[i][iF]||'')==='ฟรี', pd0=(st0==='ชำระแล้ว')||fr0;
    if(onlyPaid?!pd0:(st0==='ปฏิเสธ'||st0==='ยกเลิก')) continue;
    if(!cand[r0]){ cand[r0]=1; nCand++; }
    if(info[r0].free||cid==='committee') continue;
    var mc=compIdx?(compIdx[r0]||null):null;
    if(cid!=='all'&&mc&&mc.indexOf(cid)<0) continue; // ไม่มีรายการนี้ → ไม่ออกคูปองใน cid นี้ (ไม่ต้องอ่านชีท)
    Object.keys(COMP_SHEET).forEach(function(c){ if(c!=='c00'&&(!mc||mc.indexOf(c)>=0)) need[c]=1; });
  }
  var memIdx={}, needCids=Object.keys(COMP_SHEET).filter(function(c){return need[c];});
  if(nCand>2000) Logger.log('BATCH: '+nCand+' registrants > 2000 → โหลดชีทรายการทีละกลุ่ม ('+needCids.length+' ชีท) เก็บเฉพาะแถวที่ต้องใช้');
  var grp=nCand>2000?3:needCids.length||1;
  for(var gi=0;gi<needCids.length;gi+=grp){
    needCids.slice(gi,gi+grp).forEach(function(c){ memIdx[c]=_batchMemberIndex(ss,c,cand); });
    if(nCand>2000) Logger.log('BATCH: loaded '+Math.min(gi+grp,needCids.length)+'/'+needCids.length+' sheets');
  }
  for(i=1;i<v.length;i++){
    var rid=String(v[i][0]||'').trim(); if(!rid) continue;
    var st=String(v[i][iS]||''), free=String(v[i][iF]||'')==='ฟรี', paid=(st==='ชำระแล้ว')||free;
    if(onlyPaid?!paid:(st==='ปฏิเสธ'||st==='ยกเลิก')) continue;
    var name=String(v[i][iN]||'').trim(), org=String(v[i][iO]||'');
    var b; try{ b=_couponListFromRows(rid,name,info[rid].free,compIdx,memIdx,info[rid].al); }catch(err){ Logger.log('BATCH ['+rid+']: '+err.message); continue; }
    var cp=_mealCoupons(b.compList,b.isFree,name);
    ['17','18'].forEach(function(d){
      if(day!=='both'&&day!==d) return;
      (cp['d'+d]||[]).forEach(function(x){
        var comp=x.comp||'กรรมการ';
        if(cid==='all'){} else if(cid==='committee'){ if(comp!=='กรรมการ') return; } else if(comp!==wantComp) return;
        items.push({day:d,dayLbl:d==='17'?MD17:MD18,cid:nmCid[comp]||'committee',n:x.n,total:x.total,name:x.name,comp:comp,allergy:x.allergy,rid:rid,org:org,regName:name,unpaid:!paid});
        regs[rid]=1;
      });
    });
  }
  items.sort(function(a,b){
    if(a.day!==b.day) return a.day<b.day?-1:1;
    if(a.comp!==b.comp) return a.comp<b.comp?-1:1;
    if(a.regName!==b.regName) return a.regName<b.regName?-1:1;
    if(a.rid!==b.rid) return a.rid<b.rid?-1:1;
    return a.n-b.n;
  });
  var seq={}; // เลขลำดับเริ่มใหม่ทุก session (วัน+รายการ)
  items.forEach(function(it){ var k=it.day+'|'+it.comp; it.seq=seq[k]=(seq[k]||0)+1; });
  return {items:items,regs:regs};
}
function _batchServe(e){
  var q=e.parameter||{}, day=String(q.day||''), cid=String(q.cid||''), paid=String(q.paid||''), exp=String(q.exp||''), k=String(q.k||'');
  if(!_batchParamsOk(day,cid,paid)||!/^\d{9,12}$/.test(exp)) return _printErrorPage('ลิงก์ไม่ถูกต้อง','กรุณาสร้างลิงก์ใหม่จากหน้าเจ้าหน้าที่');
  var a=_batchSig(day,cid,paid,exp), dd=a.length^k.length;
  for(var i=0;i<a.length;i++) dd|=a.charCodeAt(i)^(k.charCodeAt(i)||0);
  if(dd!==0) return _printErrorPage('ลิงก์ไม่ถูกต้อง','กรุณาสร้างลิงก์ใหม่จากหน้าเจ้าหน้าที่');
  if(Number(exp)<Math.floor(Date.now()/1000)) return _printErrorPage('ลิงก์หมดอายุ','ลิงก์พิมพ์รวมมีอายุ 2 ชั่วโมง กรุณาสร้างลิงก์ใหม่จากหน้าเจ้าหน้าที่');
  var b=_batchBuild(day,cid,paid==='1');
  var html=_batchHtml(b,day,cid,paid==='1');
  return HtmlService.createHtmlOutput(html).setTitle('คูปองรวม · '+BATCH_DAYS[day])
    .addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function _batchCss(){
  return _printCss()+
  '.card .sq{position:absolute;right:2mm;bottom:1.5mm;font:700 9pt \'Courier New\',monospace;color:#7a5a12}'+
  '.card .og{font-size:7.5pt;color:#8a7a55;line-height:1.15;max-height:2.3em;overflow:hidden}'+
  '.card .rw{display:flex;align-items:center;gap:2mm}.card .unp{border:0.6mm solid #b91c1c;color:#b91c1c;font-weight:700;font-size:9pt;padding:0 2mm;border-radius:1.5mm;transform:rotate(-4deg)}'+
  '.card .bd{padding-right:3.5mm}.card .id{margin-right:11mm}'+
  '.card .chip{font:700 7.5pt \'Courier New\',monospace;color:#0b5844;border:0.4mm solid #0b5844;border-radius:1.2mm;padding:0 1.2mm;line-height:1.3}'+
  '.sh{border-bottom:2px solid #b8862b;padding-bottom:2mm;margin-bottom:5mm;color:#0b5844}'+
  '.sh .sn{font-size:17pt;font-weight:700;line-height:1.25}.sh .ss{font-size:10.5pt;color:#6b5a3a;margin-top:1mm}'+
  '.sh .sm{display:flex;justify-content:space-between;gap:4mm;font-size:10.5pt;font-weight:700;margin-top:1mm}.sh .sm span:last-child{white-space:nowrap}'+
  '.sh.c{padding-bottom:1.5mm;margin-bottom:4mm}.sh.c .sn{font-size:12pt}.sh.c .sm{margin-top:0;font-weight:400}'+
  'table.t.sm th,table.t.sm td{font-size:11pt;width:auto}table.t.sm th{width:auto;text-align:center}table.t.sm td.r{text-align:right}'+
  'table.t.sm tr.tot td{font-weight:700;background:#f6efe1}'+
  '@media screen and (max-width:230mm){.sh .sm{flex-direction:column}}';
}
function _batchCard(x){
  var dp=(x.dayLbl).split(' ');
  return '<div class="card"><div class="stub"><i>LUNCH</i><s>'+_esc(dp[0])+'</s><em>'+_esc(dp[1]||'')+' 2570</em></div>'+
    '<div class="bd"><div class="rw"><span class="no">คูปอง '+x.n+'/'+x.total+'</span><span class="chip">'+_esc(x.cid)+'</span>'+(x.unpaid?'<span class="unp">ยังไม่ชำระ</span>':'')+'</div><div class="nm">'+_esc(x.name)+'</div>'+
    '<div class="cp">'+_esc(x.comp)+'</div>'+
    (x.org?'<div class="og">'+_esc(x.org)+'</div>':'')+
    (x.allergy?'<div class="al">⚠ แพ้อาหาร: '+_esc(x.allergy)+'</div>':'')+
    '<div class="id">'+_esc(x.rid)+'</div></div>'+
    '<span class="sq">#'+('000'+x.seq).slice(-3)+'</span></div>';
}
function _batchHtml(b,day,cid,onlyPaid){
  var items=b.items, now=Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm');
  var dayTxt=BATCH_DAYS[day], cidTxt=_batchCidLabel(cid), per=8;
  // session = วัน + รายการ (เรียงวัน 27 → 28; ภายในวันเรียงตามชื่อรายการ — items เรียงมาแล้ว)
  var groups=[], gi={};
  items.forEach(function(it){
    var k=it.day+'|'+it.comp, g=gi[k];
    if(!g){ g=gi[k]={day:it.day,dayLbl:it.dayLbl,comp:it.comp,cid:it.cid,list:[],regs:{}}; groups.push(g); }
    g.list.push(it); g.regs[it.rid]=1;
  });
  var sheetNo=2; // แผ่นที่ 1 = หน้าปก
  groups.forEach(function(g){ g.np=Math.ceil(g.list.length/per); g.p0=sheetNo; sheetNo+=g.np; });
  var total=sheetNo-1;
  // หน้าปก: สรุปตามรายการ + ช่วงหน้า
  var by={}, order=[];
  groups.forEach(function(g){
    var r=by[g.comp]; if(!r){ r=by[g.comp]={regs:{},c17:0,c18:0,pg:[]}; order.push(g.comp); }
    Object.keys(g.regs).forEach(function(x){ r.regs[x]=1; });
    if(g.day==='17') r.c17+=g.list.length; else r.c18+=g.list.length;
    r.pg.push({day:g.dayLbl,rng:g.np>1?(g.p0+'–'+(g.p0+g.np-1)):String(g.p0)});
  });
  order.sort();
  var t17=0,t18=0, rows=order.map(function(c){
    var r=by[c], n=Object.keys(r.regs).length; t17+=r.c17; t18+=r.c18;
    var pg=r.pg.map(function(x){ return (day==='both'?_esc(x.day.split(' ')[0])+' : ':'')+x.rng; }).join('<br>');
    return '<tr><td>'+_esc(c)+'</td><td class="r">'+n+'</td><td class="r">'+r.c17+'</td><td class="r">'+r.c18+'</td><td class="r"><b>'+(r.c17+r.c18)+'</b></td><td class="r pg">'+pg+'</td></tr>';
  }).join('');
  var nReg=Object.keys(b.regs).length;
  if(!rows) rows='<tr><td colspan="6" style="text-align:center;color:#6b7280">ไม่พบคูปองตามเงื่อนไขนี้</td></tr>';
  var cover='<section class="page"><div class="hd"><div class="k">Meal Coupons · Summary</div><h1>สรุปคูปองอาหาร · '+_esc(dayTxt)+' · '+_esc(cidTxt)+'</h1><div class="en">'+_esc(CONF_NAME_TH)+'</div></div>'+
    '<table class="t sm"><tr><th style="text-align:left">รายการ</th><th>ทีม/คน</th><th>'+MD17+'</th><th>'+MD18+'</th><th>คูปองรวม</th><th>หน้าที่</th></tr>'+rows+
    '<tr class="tot"><td>รวมทั้งหมด</td><td class="r">'+nReg+'</td><td class="r">'+t17+'</td><td class="r">'+t18+'</td><td class="r">'+items.length+'</td><td class="r">'+(groups.length?'2–'+total:'-')+'</td></tr></table>'+
    '<div class="note">วัน: '+_esc(dayTxt)+' · รายการ: '+_esc(cidTxt)+' · '+(onlyPaid?'เฉพาะที่ชำระแล้ว/กรรมการ':'รวมที่ยังไม่ชำระ (ประทับ "ยังไม่ชำระ")')+'<br>สร้างเมื่อ '+now+' · คูปองทั้งหมด '+items.length+' ใบ · หน้าที่ = แผ่นที่ในชุดพิมพ์นี้ (หน้าปก = 1) · แต่ละ session เริ่มหน้าใหม่</div></section>';
  var body=groups.map(function(g){
    var nTeam=Object.keys(g.regs).length, sch=g.cid==='committee'?'ทั้ง 2 วัน':_schedLine(g.cid), out='';
    for(var pi=0;pi<g.np;pi++){
      var part=g.list.slice(pi*per,(pi+1)*per);
      var meta='<div class="sm"><span>คูปอง '+_esc(g.dayLbl)+' 2570 · '+g.list.length+' ใบ · '+nTeam+' ทีม/คน</span><span>หน้า '+(pi+1)+'/'+g.np+' ของ session</span></div>';
      var hdr=pi===0
        ? '<div class="sh"><div class="sn">'+_esc(g.comp)+'</div>'+(sch?'<div class="ss">'+_esc(sch)+'</div>':'')+meta+'</div>'
        : '<div class="sh c"><div class="sn">'+_esc(g.comp)+'</div>'+meta+'</div>';
      out+='<section class="page">'+hdr+'<div class="grid">'+part.map(_batchCard).join('')+'</div></section>';
    }
    return out;
  }).join('');
  return '<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<title>คูปองรวม '+_esc(dayTxt)+'</title>'+
    '<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&family=Prompt:wght@500;700&display=swap" rel="stylesheet">'+
    '<style>'+_batchCss()+'</style></head><body>'+
    '<div class="bar"><button type="button" onclick="window.print()">พิมพ์ / บันทึกเป็น PDF</button><span>A4 · พิมพ์สี · '+items.length+' ใบ · '+groups.length+' session · '+total+' หน้า</span></div>'+
    '<div class="sheets">'+cover+body+'</div></body></html>';
}
