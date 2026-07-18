/*************************************************************
 * ระบบส่งเรื่องเข้าที่ประชุมคณะกรรมการบริหาร CULAC — Backend (GAS)
 * ---------------------------------------------------------
 * วิธีใช้ (ครั้งแรก):
 *  1) เปิด script.google.com > New project > วางโค้ดนี้ทั้งหมด
 *  2) กด Run เลือกฟังก์ชัน setup() หนึ่งครั้ง (อนุญาตสิทธิ์)
 *     -> จะสร้าง Google Sheet "CULAC-Agenda-DB" และโฟลเดอร์ Drive
 *        "CULAC-Agenda-Files" ให้อัตโนมัติ
 *  3) Deploy > New deployment > Type: Web app
 *        - Execute as: Me
 *        - Who has access: Anyone            <-- สำคัญมาก
 *     คัดลอกลิงก์ที่ลงท้าย /exec ไปวางในหน้า "ตั้งค่า" ของระบบ
 *  4) ทุกครั้งที่แก้โค้ด ให้ Deploy > Manage deployments > Edit > Version: New
 *************************************************************/

var P = PropertiesService.getScriptProperties();

/* ---------- ตั้งค่าเริ่มต้น (รันครั้งเดียว) ---------- */
function setup() {
  var ss = SpreadsheetApp.create('CULAC-Agenda-DB');
  var mSheet = ss.getSheets()[0].setName('Meetings');
  mSheet.appendRow(['id','round','year','date','time','venue','prev','status','createdAt']);
  var iSheet = ss.insertSheet('Items');
  iSheet.appendRow(['id','ref','meetingId','dept','proposer','contact','type','title','detail','status','order','category','urgency','amount','resolution','resolutionNote','owner','files','createdAt']);

  var rSheet = ss.insertSheet('Recipients');
  rSheet.appendRow(['email','name']);

  var folder = DriveApp.createFolder('CULAC-Agenda-Files');
  P.setProperty('SHEET_ID', ss.getId());
  P.setProperty('FOLDER_ID', folder.getId());
  Logger.log('เสร็จสิ้น: Sheet=' + ss.getUrl() + ' | Folder=' + folder.getUrl());
}

/* ---------- อัปเกรดฐานข้อมูลเดิม (รันครั้งเดียวหลังอัปเดตโค้ด) ----------
 * ใช้เมื่อเคย setup() ไปแล้วก่อนมีระบบจำแนกเรื่อง/รายงาน
 * จะเพิ่มคอลัมน์ใหม่ให้ชีต Items โดยไม่กระทบข้อมูลเดิม
 */
function migrate() {
  var sh = sh_('Items');
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var needed = ['category','urgency','amount','resolution','resolutionNote','owner'];
  var added = [];
  needed.forEach(function(col){
    if (head.indexOf(col) === -1) {
      // แทรกก่อนคอลัมน์ files เพื่อคงลำดับให้อ่านง่าย
      var filesIdx = head.indexOf('files');
      var insertAt = filesIdx === -1 ? sh.getLastColumn() + 1 : filesIdx + 1;
      sh.insertColumnBefore(insertAt);
      sh.getRange(1, insertAt).setValue(col);
      head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      added.push(col);
    }
  });
  // ตั้งค่าเริ่มต้นให้แถวเดิมที่ยังว่าง
  var last = sh.getLastRow();
  if (last > 1) {
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var uCol = head.indexOf('urgency') + 1, rCol = head.indexOf('resolution') + 1;
    for (var r = 2; r <= last; r++) {
      if (uCol > 0 && !sh.getRange(r, uCol).getValue()) sh.getRange(r, uCol).setValue('normal');
      if (rCol > 0 && !sh.getRange(r, rCol).getValue()) sh.getRange(r, rCol).setValue('none');
    }
  }
  // สร้างชีตรายชื่อผู้รับอีเมล ถ้ายังไม่มี
  if (!ss_().getSheetByName('Recipients')) {
    var rs = ss_().insertSheet('Recipients');
    rs.appendRow(['email','name']);
    added.push('ชีต Recipients');
  }
  Logger.log(added.length ? 'อัปเดต: ' + added.join(', ') : 'ฐานข้อมูลเป็นเวอร์ชันล่าสุดอยู่แล้ว');
}

/* ---------- helpers ---------- */
function ss_()   { return SpreadsheetApp.openById(P.getProperty('SHEET_ID')); }
function sh_(n)  { return ss_().getSheetByName(n); }
function folder_(){ return DriveApp.getFolderById(P.getProperty('FOLDER_ID')); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function readSheet_(name){
  var sh = sh_(name), vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var head = vals[0], out = [];
  for (var r=1; r<vals.length; r++){
    var o = {};
    for (var c=0; c<head.length; c++) o[head[c]] = vals[r][c];
    if (o.files){ try { o.files = JSON.parse(o.files); } catch(e){ o.files = []; } } else if (name==='Items') o.files = [];
    out.push(o);
  }
  return out;
}
function findRow_(name, id){
  var sh = sh_(name), ids = sh.getRange(1,1,sh.getLastRow(),1).getValues();
  for (var r=1; r<ids.length; r++) if (String(ids[r][0])===String(id)) return r+1;
  return -1;
}
function rowFrom_(name, obj){
  var head = sh_(name).getRange(1,1,1,sh_(name).getLastColumn()).getValues()[0];
  return head.map(function(k){
    if (k==='files') return JSON.stringify(obj.files||[]);
    return obj[k]!=null ? obj[k] : '';
  });
}

/* ---------- GET: ส่งข้อมูลทั้งหมด ---------- */
function doGet(e){
  try{
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action==='list'){
      var rcp = [];
      try { rcp = readSheet_('Recipients'); } catch (e2) { rcp = []; }
      return json_({ ok:true, meetings: readSheet_('Meetings'), items: readSheet_('Items'), recipients: rcp });
    }
    return json_({ ok:false, error:'unknown action' });
  }catch(err){ return json_({ ok:false, error:String(err) }); }
}

/* ---------- POST: บันทึก/แก้ไข/ลบ ---------- */
function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    var a = body.action;

    if (a==='submit'){
      var item = body.item;
      item.files = saveFiles_(body.files, item.ref || item.id);
      sh_('Items').appendRow(rowFrom_('Items', item));
      return json_({ ok:true, item:item });
    }

    if (a==='updateItem'){
      var it = body.item, row = findRow_('Items', it.id);
      if (row<0) return json_({ ok:false, error:'ไม่พบเรื่อง' });
      // คงไฟล์เดิมไว้ถ้า payload ไม่ได้ส่ง files ใหม่มา
      if (!it.files){
        var existing = readSheet_('Items').filter(function(x){return x.id===it.id;})[0];
        it.files = existing ? existing.files : [];
      }
      sh_('Items').getRange(row,1,1,sh_('Items').getLastColumn()).setValues([rowFrom_('Items', it)]);
      return json_({ ok:true });
    }

    if (a==='deleteItem'){
      var r = findRow_('Items', body.id);
      if (r>0) sh_('Items').deleteRow(r);
      return json_({ ok:true });
    }

    if (a==='saveMeeting'){
      var m = body.meeting, mr = findRow_('Meetings', m.id);
      if (mr<0) sh_('Meetings').appendRow(rowFrom_('Meetings', m));
      else sh_('Meetings').getRange(mr,1,1,sh_('Meetings').getLastColumn()).setValues([rowFrom_('Meetings', m)]);
      return json_({ ok:true });
    }

    if (a==='deleteMeeting'){
      var mrow = findRow_('Meetings', body.id);
      if (mrow>0) sh_('Meetings').deleteRow(mrow);
      // ลบเรื่องในรอบนี้ด้วย
      var items = readSheet_('Items');
      for (var i=0;i<items.length;i++) if (items[i].meetingId===body.id){
        var rr = findRow_('Items', items[i].id); if (rr>0) sh_('Items').deleteRow(rr);
      }
      return json_({ ok:true });
    }

    if (a==='reorder'){
      var orders = body.orders||[];
      var sh = sh_('Items'), data = sh.getDataRange().getValues(), head = data[0];
      var idCol = head.indexOf('id'), ordCol = head.indexOf('order');
      var map = {}; orders.forEach(function(o){ map[o.id]=o.order; });
      for (var r=1;r<data.length;r++){
        var id = data[r][idCol];
        if (map[id]!=null) sh.getRange(r+1, ordCol+1).setValue(map[id]);
      }
      return json_({ ok:true });
    }

    if (a==='saveRecipients'){
      var rs = ss_().getSheetByName('Recipients');
      if (!rs){ rs = ss_().insertSheet('Recipients'); rs.appendRow(['email','name']); }
      if (rs.getLastRow() > 1) rs.deleteRows(2, rs.getLastRow()-1);
      var list = body.recipients || [];
      list.forEach(function(r){ rs.appendRow([r.email||'', r.name||'']); });
      return json_({ ok:true, count:list.length });
    }

    if (a==='sendInvite'){
      return sendInvite_(body);
    }

    return json_({ ok:false, error:'unknown action: '+a });
  }catch(err){ return json_({ ok:false, error:String(err) }); }
}

/* ---------- อัปโหลดไฟล์เข้า Drive ---------- */
function saveFiles_(files, refPrefix){
  var out = [];
  if (!files || !files.length) return out;
  var folder = folder_();
  files.forEach(function(f){
    var bytes = Utilities.base64Decode(f.data);
    var blob  = Utilities.newBlob(bytes, f.mime || 'application/octet-stream', (refPrefix||'file')+'_'+f.name);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    out.push({ id:file.getId(), name:f.name, size:bytes.length, url:file.getUrl() });
  });
  return out;
}


/* ---------- ส่งอีเมลเชิญประชุม ----------
 * body = { meeting:{...}, link:'https://...', qr:'<base64 png>', items:[{no,title}], note:'' }
 */
function sendInvite_(body){
  var m = body.meeting || {};
  var rs = ss_().getSheetByName('Recipients');
  if (!rs || rs.getLastRow() < 2) return json_({ ok:false, error:'ยังไม่มีรายชื่อผู้รับอีเมล' });

  var rows = rs.getRange(2, 1, rs.getLastRow()-1, 2).getValues();
  var emails = rows.map(function(r){ return String(r[0]).trim(); })
                   .filter(function(e){ return e && e.indexOf('@') > 0; });
  if (!emails.length) return json_({ ok:false, error:'ไม่พบอีเมลที่ถูกต้องในรายชื่อ' });

  var subject = 'ขอเชิญประชุมคณะกรรมการบริหารศูนย์สัตว์ทดลอง ครั้งที่ ' + m.round + '/' + m.year;

  var agenda = '';
  if (body.items && body.items.length){
    agenda = '<p style="margin:18px 0 6px;font-weight:bold;color:#152744">วาระการประชุมโดยสรุป</p><ol style="margin:0;padding-left:20px;color:#232b36">';
    body.items.forEach(function(i){
      agenda += '<li style="margin:3px 0">' + escapeHtml_(i.title) + '</li>';
    });
    agenda += '</ol>';
  }

  var html =
    '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:15px;color:#232b36;max-width:600px">' +
      '<div style="background:#152744;color:#fff;padding:18px 22px;border-radius:10px 10px 0 0">' +
        '<div style="font-size:13px;color:#e5c477">ขอเชิญประชุม</div>' +
        '<div style="font-size:19px;font-weight:bold;margin-top:4px">คณะกรรมการบริหารศูนย์สัตว์ทดลอง</div>' +
        '<div style="font-size:14px;color:#c6d2e6">จุฬาลงกรณ์มหาวิทยาลัย</div>' +
      '</div>' +
      '<div style="border:1px solid #e3e6ec;border-top:none;border-radius:0 0 10px 10px;padding:22px">' +
        '<table style="width:100%;border-collapse:collapse;font-size:15px">' +
          row_('ครั้งที่', m.round + '/' + m.year) +
          row_('วัน–เวลา', (m.date || '-') + (m.time ? '  เวลา ' + m.time : '')) +
          row_('สถานที่', m.venue || '-') +
        '</table>' +
        (body.note ? '<p style="margin-top:14px">' + escapeHtml_(body.note) + '</p>' : '') +
        agenda +
        '<div style="margin-top:22px;padding-top:18px;border-top:1px solid #e3e6ec;text-align:center">' +
          '<p style="margin:0 0 10px;font-weight:bold;color:#152744">ดูวาระการประชุมและส่งเรื่องเข้าที่ประชุม</p>' +
          (body.qr ? '<img src="cid:qrcode" width="180" style="border:1px solid #e3e6ec;border-radius:8px"><br>' : '') +
          '<a href="' + (body.link || '#') + '" style="display:inline-block;margin-top:12px;background:#c8671a;color:#fff;' +
          'padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:bold">เปิดระบบวาระการประชุม</a>' +
          '<p style="font-size:12px;color:#69727f;margin-top:10px;word-break:break-all">' + (body.link || '') + '</p>' +
        '</div>' +
      '</div>' +
      '<p style="font-size:12px;color:#69727f;text-align:center;margin-top:14px">' +
        'อีเมลฉบับนี้ส่งจากระบบส่งเรื่องเข้าที่ประชุมคณะกรรมการบริหาร ศูนย์สัตว์ทดลอง จุฬาลงกรณ์มหาวิทยาลัย</p>' +
    '</div>';

  var opts = { htmlBody: html, name: 'ระบบวาระการประชุม CULAC' };
  if (body.qr){
    var blob = Utilities.newBlob(Utilities.base64Decode(body.qr), 'image/png', 'qr.png');
    opts.inlineImages = { qrcode: blob };
  }

  var sent = 0, failed = [];
  emails.forEach(function(e){
    try { MailApp.sendEmail(Object.assign({ to: e, subject: subject }, opts)); sent++; }
    catch (err) { failed.push(e); }
  });

  return json_({ ok:true, sent:sent, failed:failed, quotaLeft: MailApp.getRemainingDailyQuota() });
}

function row_(label, value){
  return '<tr>' +
    '<td style="padding:7px 0;color:#69727f;width:110px;vertical-align:top">' + label + '</td>' +
    '<td style="padding:7px 0;font-weight:bold;color:#152744">' + escapeHtml_(String(value)) + '</td></tr>';
}

function escapeHtml_(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
