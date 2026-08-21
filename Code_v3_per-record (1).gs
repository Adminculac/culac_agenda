/**
 * ระบบบริหารโครงการ CULAC — Backend (Google Apps Script)  v3  "per-record"
 * -------------------------------------------------------------------------
 * แก้ที่รากของปัญหาความช้า: เปลี่ยนจากเก็บ "ก้อน JSON เดียว" มาเป็น
 * "ฐานข้อมูลจริงแบบทีละรายการ" — โครงการละ 1 แถว, ค่าตั้งค่าละ 1 แถว
 * ทำให้หน้าเว็บส่ง/ดึงเฉพาะรายการที่เปลี่ยน (ไม่ใช่ทั้งก้อน) จึงเร็วขึ้นมาก
 * และหลายคนแก้คนละโครงการพร้อมกันได้โดยไม่ทับกัน
 *
 * แผ่นงานที่ใช้ (สร้างให้อัตโนมัติ)
 *   PROJ : A=id   B=rev  C=deleted(1/0)  D=json(ของโครงการนั้น)
 *   KV   : A=key  B=rev  C=json           (persons, roles, emails, pw, kpis,
 *                                           seq, intake, meetDates, meetings,
 *                                           years, fy, audit)
 *   SYS  : A1=rev(ตัวนับรุ่นข้อมูลรวม)
 *
 * API
 *   GET  ?action=rev                → {ok, rev}
 *   GET  ?action=pull&since=N       → {ok, rev, projects:[{id,value,deleted}], kv:{key:value}}
 *                                     (คืนเฉพาะรายการที่ rev > N ; since=0 = ทั้งหมด)
 *   POST body {changes:{projects:{id:obj|null}, kv:{key:obj}}}
 *                                     → {ok, rev}   (null = ลบโครงการนั้น)
 *
 * การย้ายข้อมูลจากรุ่นเก่า: ครั้งแรกที่รัน ถ้า PROJ ว่างและมีแผ่น 'DB' (ก้อน JSON เดิม)
 * ระบบจะแตกก้อนเดิมออกเป็นแถว ๆ ให้อัตโนมัติ — ข้อมูลเดิมไม่หาย
 *
 * ติดตั้ง: วางทับ Code.gs → ปรับใช้เป็นเว็บแอป (Execute as: Me, Who has access: Anyone)
 * แนะนำให้ deploy เป็น "การปรับใช้ใหม่" (URL ใหม่) เพื่อทดสอบคู่กับของเดิมก่อนสลับใช้จริง
 */

var META_KEYS = ['persons','roles','emails','pw','kpis','seq','intake','meetDates','meetings','years','fy','audit'];

/* ==== แจ้งเตือนทางอีเมล ====
 * เปลี่ยนอีเมลเลขานุการที่นี่ได้ตามต้องการ */
var SEC_EMAIL = 'parichart.sa@chula.ac.th';
var TH_MONTHS_ = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

/** จัดรูปแบบตัวเลขเงิน เช่น 85000 -> "85,000.00" */
function fmtMoney_(n){
  n = Number(n)||0;
  var neg = n<0; n = Math.abs(n);
  var parts = n.toFixed(2).split('.');
  var intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg?'-':'') + intPart + '.' + parts[1];
}
/** จัดรูปแบบวันที่ไทย เช่น "19 สิงหาคม 2570" จาก ISO string */
function fmtThaiDate_(iso){
  if(!iso) return '-';
  var d = new Date(iso);
  if(isNaN(d.getTime())) return '-';
  return d.getDate() + ' ' + TH_MONTHS_[d.getMonth()] + ' ' + (d.getFullYear()+543);
}

function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, headers){
  var ss = ss_(), sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); if(headers) sh.getRange(1,1,1,headers.length).setValues([headers]); sh.hideSheet(); }
  return sh;
}
function sys_(){ return sheet_('SYS'); }
function getRev_(){ var v = sys_().getRange('A1').getValue(); return (typeof v==='number' && !isNaN(v)) ? v : 0; }
function setRev_(n){ sys_().getRange('A1').setValue(n); }

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** เตรียมแผ่นงาน + ย้ายข้อมูลจากก้อนเดิมครั้งแรก */
function ensure_(){
  var proj = sheet_('PROJ', ['id','rev','deleted','json']);
  var kv   = sheet_('KV',   ['key','rev','json']);
  sys_();
  if(getRev_()===0 && proj.getLastRow() < 2){
    migrateFromBlob_(proj, kv);
  }
  return { proj: proj, kv: kv };
}

/** อ่านก้อน JSON เดิมจากแผ่น 'DB' (ถ้ามี) แล้วแตกเป็นแถว */
function migrateFromBlob_(proj, kv){
  var ss = ss_(), old = ss.getSheetByName('DB');
  if(!old) return;
  var raw = old.getRange('A1').getValue();
  if(!raw) return;
  var data;
  try{ data = JSON.parse(raw); }catch(e){ return; }
  if(!data || typeof data !== 'object') return;
  var rev = 0;
  var projRows = [];
  (data.projects||[]).forEach(function(p){
    if(!p || p.id==null) return;
    rev++; projRows.push([String(p.id), rev, 0, JSON.stringify(p)]);
  });
  if(projRows.length) proj.getRange(2,1,projRows.length,4).setValues(projRows);
  var kvRows = [];
  META_KEYS.forEach(function(k){
    var v = (k==='audit') ? (data.audit||[]) : (data[k]===undefined ? null : data[k]);
    rev++; kvRows.push([k, rev, JSON.stringify(v)]);
  });
  if(kvRows.length) kv.getRange(2,1,kvRows.length,3).setValues(kvRows);
  setRev_(rev);
}

/** อ่านรายการที่ rev > since (โครงการ + kv) */
function pullSince_(s, since){
  since = Number(since)||0;
  var out = { ok:true, rev:getRev_(), projects:[], kv:{} };
  var pv = s.proj.getLastRow()>=2 ? s.proj.getRange(2,1,s.proj.getLastRow()-1,4).getValues() : [];
  pv.forEach(function(r){
    if(!r[0] && r[0]!==0) return;
    if((Number(r[1])||0) > since){
      var rec = { id:String(r[0]), deleted: r[2]==1 || r[2]==='1' };
      if(!rec.deleted){ try{ rec.value = JSON.parse(r[3]||'null'); }catch(_){ rec.value=null; } }
      out.projects.push(rec);
    }
  });
  var kvv = s.kv.getLastRow()>=2 ? s.kv.getRange(2,1,s.kv.getLastRow()-1,3).getValues() : [];
  kvv.forEach(function(r){
    if(!r[0]) return;
    if((Number(r[1])||0) > since){ try{ out.kv[String(r[0])] = JSON.parse(r[2]||'null'); }catch(_){ out.kv[String(r[0])]=null; } }
  });
  return out;
}

function doGet(e){
  try{
    var p = (e && e.parameter) || {};
    var s = ensure_();
    if(p.action === 'rev'){ return jsonOut_({ ok:true, rev:getRev_() }); }
    if(p.action === 'pull'){ return jsonOut_(pullSince_(s, p.since)); }
    return jsonOut_({ ok:false, error:'unknown action' });
  }catch(err){ return jsonOut_({ ok:false, error:String(err) }); }
}

function doPost(e){
  try{
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var changes = body.changes || {};
    var since = (body.since==null) ? null : Number(body.since);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try{
      var s = ensure_();
      var rev = getRev_();

      // ---- projects ----
      var projChanges = changes.projects || {};
      var ids = Object.keys(projChanges);
      if(ids.length){
        var pRange = s.proj.getLastRow()>=2 ? s.proj.getRange(2,1,s.proj.getLastRow()-1,4) : null;
        var pVals = pRange ? pRange.getValues() : [];
        var rowOf = {};
        for(var i=0;i<pVals.length;i++){ rowOf[String(pVals[i][0])] = i; }
        var appends = [];
        var notify = [];
        ids.forEach(function(id){
          rev++;
          var v = projChanges[id];
          var del = (v===null || v===undefined) ? 1 : 0;
          var json = del ? '' : JSON.stringify(v);
          var oldStatus = null;
          if(rowOf[id]!==undefined){
            var oldJson = pVals[rowOf[id]][3];
            if(oldJson){ try{ oldStatus = JSON.parse(oldJson).status; }catch(_){ } }
          }
          if(!del && v && v.status){
            if(v.status==='submitted' && oldStatus!=='submitted'){ notify.push({type:'submitted', project:v}); }
            if(v.status==='running' && oldStatus!=='running'){ notify.push({type:'approved', project:v}); }
          }
          if(rowOf[id]!==undefined){
            var ri = rowOf[id];
            pVals[ri][1]=rev; pVals[ri][2]=del; pVals[ri][3]=json;
          }else{
            appends.push([String(id), rev, del, json]);
          }
        });
        if(pVals.length) s.proj.getRange(2,1,pVals.length,4).setValues(pVals);
        if(appends.length) s.proj.getRange(s.proj.getLastRow()+1,1,appends.length,4).setValues(appends);
      }

      // ---- kv ----
      var kvChanges = changes.kv || {};
      var keys = Object.keys(kvChanges);
      if(keys.length){
        var kRange = s.kv.getLastRow()>=2 ? s.kv.getRange(2,1,s.kv.getLastRow()-1,3) : null;
        var kVals = kRange ? kRange.getValues() : [];
        var krow = {};
        for(var m=0;m<kVals.length;m++){ krow[String(kVals[m][0])] = m; }
        var kApp = [];
        keys.forEach(function(k){
          rev++;
          var json = JSON.stringify(kvChanges[k]);
          if(krow[k]!==undefined){ kVals[krow[k]][1]=rev; kVals[krow[k]][2]=json; }
          else{ kApp.push([k, rev, json]); }
        });
        if(kVals.length) s.kv.getRange(2,1,kVals.length,3).setValues(kVals);
        if(kApp.length) s.kv.getRange(s.kv.getLastRow()+1,1,kApp.length,3).setValues(kApp);
      }

      setRev_(rev);
      var result = (since!=null) ? pullSince_(s, since) : { ok:true, rev:rev };
      var toNotify = (typeof notify!=='undefined') ? notify : [];
      lock.releaseLock();
      if(toNotify.length){ try{ sendNotifications_(toNotify); }catch(notifyErr){ /* ไม่ให้กระทบผลบันทึกหลัก */ } }
      return jsonOut_(result);
    } catch(errInner){ try{ lock.releaseLock(); }catch(_){ } throw errInner; }
  }catch(err){ return jsonOut_({ ok:false, error:String(err) }); }
}

/** อ่านค่า KV ปัจจุบันของ key เดียว (ใช้หาอีเมลของเจ้าของโครงการ) */
function readKvValue_(key){
  var kv = sheet_('KV', ['key','rev','json']);
  var last = kv.getLastRow();
  if(last<2) return null;
  var vals = kv.getRange(2,1,last-1,3).getValues();
  for(var i=0;i<vals.length;i++){
    if(String(vals[i][0])===key){ try{ return JSON.parse(vals[i][2]||'null'); }catch(_){ return null; } }
  }
  return null;
}

/** ส่งอีเมลแจ้งเตือน: โครงการใหม่ → เลขานุการ, โครงการที่อนุมัติ → เจ้าของโครงการ */
function sendNotifications_(list, overrideEmails){
  var emails = overrideEmails || (readKvValue_('emails') || {});
  list.forEach(function(n){
    try{
      var p = n.project;
      if(n.type==='submitted'){
        var textBody = 'มีโครงการเสนอเข้าระบบรอพิจารณา\n\n' +
              'รหัส: ' + (p.code||'') + '\n' +
              'ชื่อโครงการ: ' + (p.name||'') + '\n' +
              'ผู้เสนอ: ' + (p.owner||'') + '\n' +
              'งบที่ขอ: ' + fmtMoney_(p.budgetRequested) + ' บาท\n\n' +
              'กรุณาเข้าสู่ระบบเพื่อพิจารณาอนุมัติ';
        var htmlBody = emailShell_(
          '#14233f', '#c3d0e6', '📋',
          'มีโครงการเสนอเข้ามาใหม่',
          'รอการพิจารณาจากคณะกรรมการบริหาร',
          '' +
           emailRow_('รหัส', es_(p.code)) +
           emailRow_('ชื่อโครงการ', es_(p.name)) +
           emailRow_('ผู้เสนอ', es_(p.owner)) +
           emailRow_('งบที่ขอ', fmtMoney_(p.budgetRequested) + ' บาท', '#a8531a', '16px'),
          'กรุณาเข้าสู่ระบบเพื่อพิจารณาอนุมัติ',
          '#f4f6fa', '#8592a8'
        );
        MailApp.sendEmail({ to: SEC_EMAIL, subject: 'มีโครงการเสนอเข้ามาใหม่ · ' + (p.code||''), body: textBody, htmlBody: htmlBody });
      } else if(n.type==='approved'){
        var toEmail = emails[p.owner];
        if(toEmail){
          var meetingRef = p.meetingRef || '-';
          var dateThai = fmtThaiDate_(p.approvedAt);
          var budgetTxt = fmtMoney_(p.budgetApproved!=null?p.budgetApproved:p.budgetRequested);
          var textBody2 = 'โครงการของท่านได้รับการอนุมัติจากคณะกรรมการบริหาร ศูนย์สัตว์ทดลอง จุฬาฯ แล้ว\n' +
                'การประชุม' + meetingRef + ' วันที่ ' + dateThai + '\n\n' +
                'รหัส: ' + (p.code||'') + '\n' +
                'ชื่อโครงการ: ' + (p.name||'') + '\n' +
                'งบที่อนุมัติ: ' + budgetTxt + ' บาท\n\n' +
                'ท่านสามารถเข้าสู่ระบบเพื่อเริ่มบันทึกผลการดำเนินงานได้แล้ว';
          var htmlBody2 = emailShell_(
            '#e75480', '#fbdce8', '🎉',
            'โครงการของท่านได้รับการอนุมัติ<br>จากคณะกรรมการบริหาร ศูนย์สัตว์ทดลอง จุฬาฯ แล้ว',
            '',
            '' +
             '<tr><td colspan="2" style="padding:0 0 16px 0">' +
              '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td bgcolor="#fdf0f5" style="background-color:#fdf0f5;border:1px solid #f6c9dc;border-radius:10px;padding:11px 16px;text-align:center;color:#a13e64;font-weight:bold;font-size:13.5px;font-family:Tahoma,Arial,sans-serif">' +
               'การประชุม' + es_(meetingRef) + ' &middot; วันที่ ' + dateThai +
              '</td></tr></table>' +
             '</td></tr>' +
             emailRow_('รหัส', es_(p.code)) +
             emailRow_('ชื่อโครงการ', es_(p.name)) +
             emailRow_('งบที่อนุมัติ', budgetTxt + ' บาท', '#e75480', '17px'),
            'ท่านสามารถเข้าสู่ระบบเพื่อเริ่มบันทึกผลการดำเนินงานได้แล้ว',
            '#fdf0f5', '#c48b9f'
          );
          MailApp.sendEmail({ to: toEmail, subject: 'โครงการของท่านได้รับการอนุมัติแล้ว · ' + (p.code||''), body: textBody2, htmlBody: htmlBody2 });
        }
      }
    }catch(mailErr){ /* ไม่ให้ปัญหาการส่งอีเมลกระทบการบันทึกข้อมูลหลัก */ }
  });
}

/** โครงร่างอีเมล HTML แบบตาราง (table-based) — อีเมลไคลเอนต์ต่าง ๆ (โดยเฉพาะ Gmail) ตัด CSS gradient/background บน <div> ทิ้งเกือบทุกครั้ง
 *  จึงต้องใช้ <table> + attribute bgcolor เพื่อให้สีพื้นหลังติดแน่นอนในทุกอีเมลไคลเอนต์ */
function emailShell_(headColor, subTextColor, icon, title, subtitle, rowsHtml, footNote, footBg, footColor){
  return '' +
   '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#eef1f5;padding:24px 0">' +
    '<tr><td align="center">' +
     '<table role="presentation" width="520" cellpadding="0" cellspacing="0" style="max-width:520px;width:100%;background-color:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e3e7ee;font-family:Tahoma,Arial,sans-serif">' +
      '<tr><td bgcolor="' + headColor + '" style="background-color:' + headColor + ';padding:28px 26px;text-align:center">' +
       '<div style="font-size:34px;line-height:1.2;margin-bottom:8px">' + icon + '</div>' +
       '<div style="color:#ffffff;font-size:18.5px;font-weight:bold;line-height:1.5;font-family:Tahoma,Arial,sans-serif">' + title + '</div>' +
       (subtitle ? '<div style="color:' + subTextColor + ';font-size:13px;margin-top:6px;font-family:Tahoma,Arial,sans-serif">' + subtitle + '</div>' : '') +
      '</td></tr>' +
      '<tr><td style="padding:24px 26px">' +
       '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:14.5px;color:#333333;font-family:Tahoma,Arial,sans-serif">' +
        rowsHtml +
       '</table>' +
       '<div style="margin-top:20px;text-align:center;color:#666666;font-size:13px;line-height:1.6;font-family:Tahoma,Arial,sans-serif">' + footNote + '</div>' +
      '</td></tr>' +
      '<tr><td bgcolor="' + footBg + '" style="background-color:' + footBg + ';padding:12px;text-align:center;font-size:11.5px;color:' + footColor + ';font-family:Tahoma,Arial,sans-serif">ศูนย์สัตว์ทดลอง จุฬาลงกรณ์มหาวิทยาลัย</td></tr>' +
     '</table>' +
    '</td></tr>' +
   '</table>';
}
/** แถวข้อมูลแบบ label/value ในตารางอีเมล */
function emailRow_(label, value, valueColor, valueSize){
  var vStyle = 'padding:9px 0;font-weight:bold;color:' + (valueColor||'#222222') + ';font-size:' + (valueSize||'14.5px') + ';font-family:Tahoma,Arial,sans-serif';
  return '<tr>' +
   '<td style="padding:9px 0;color:#888888;width:105px;vertical-align:top;border-top:1px solid #f0f0f0;font-family:Tahoma,Arial,sans-serif">' + label + '</td>' +
   '<td style="' + vStyle + ';border-top:1px solid #f0f0f0">' + value + '</td>' +
  '</tr>';
}

function es_(v){
  v = (v==null?'':String(v));
  return v.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/* ============================================================
 * ฟังก์ชันทดสอบส่งอีเมล — ไม่กระทบข้อมูลจริงในชีตเลย
 * วิธีใช้: เปิดหน้านี้ใน Apps Script → เลือกชื่อฟังก์ชันที่แถบด้านบน
 *          (testSendSubmittedEmail หรือ testSendApprovedEmail) → กด "เรียกใช้" (Run)
 *          → เข้าไปเช็คอีเมลที่ระบุไว้ด้านล่าง
 * ============================================================ */

/** ทดสอบ: อีเมลแจ้ง "มีโครงการเสนอเข้ามาใหม่" — จะส่งไปที่ SEC_EMAIL ด้านบนของไฟล์ */
function testSendSubmittedEmail(){
  sendNotifications_([{ type:'submitted', project:{
    code:'CULAC-PRJ-099',
    name:'โครงการทดสอบระบบแจ้งเตือนทางอีเมล',
    owner:'ทดสอบ ระบบแจ้งเตือน',
    budgetRequested:12345
  }}]);
  Logger.log('ส่งอีเมลทดสอบ (โครงการใหม่) ไปที่ ' + SEC_EMAIL + ' แล้ว — ลองเช็คกล่องจดหมาย');
}

/** ทดสอบ: อีเมลแจ้ง "โครงการได้รับอนุมัติ" — แก้อีเมลผู้รับด้านล่างก่อนกดรัน */
function testSendApprovedEmail(){
  var TEST_RECEIVER = 'parichart.sa@chula.ac.th'; // <-- แก้เป็นอีเมลที่ต้องการทดสอบได้
  sendNotifications_([{ type:'approved', project:{
    code:'CULAC-PRJ-099',
    name:'โครงการทดสอบระบบแจ้งเตือนทางอีเมล',
    owner:'__test__',
    budgetApproved:85000,
    budgetRequested:85000,
    meetingRef:'ครั้งที่ 3/2570',
    approvedAt:new Date().toISOString()
  }}], { '__test__': TEST_RECEIVER });
  Logger.log('ส่งอีเมลทดสอบ (อนุมัติแล้ว) ไปที่ ' + TEST_RECEIVER + ' แล้ว — ลองเช็คกล่องจดหมาย');
}
