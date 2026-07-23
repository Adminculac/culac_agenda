/*************************************************************
 * ระบบส่งเรื่องเข้าที่ประชุมคณะกรรมการบริหาร
 * ศูนย์สัตว์ทดลอง จุฬาลงกรณ์มหาวิทยาลัย
 * ---------------------------------------------------------
 * ฐานข้อมูลกลาง (Google Apps Script)
 *
 * ── ติดตั้งครั้งแรก ──────────────────────────────────────
 * 1) เปิด script.google.com → New project → วางโค้ดนี้ทั้งหมด
 * 2) เลือกฟังก์ชัน  setup  แล้วกด Run หนึ่งครั้ง
 *    (ระบบขออนุญาต → Review permissions → Advanced → Go to… → Allow)
 *    ผลลัพธ์: สร้าง Google Sheet "CULAC-Agenda-DB"
 *             และโฟลเดอร์ Drive "CULAC-Agenda-Files"
 * 3) Deploy → New deployment → เลือกชนิด Web app
 *       Execute as       : Me
 *       Who has access   : Anyone      ← สำคัญที่สุด ห้ามพลาด
 * 4) คัดลอกลิงก์ที่ลงท้าย /exec
 *    ไปวางในหน้าเว็บ → ตั้งค่าระบบ → GAS Web App URL → บันทึก & ทดสอบ
 *
 * ── กรณีเคยติดตั้งเวอร์ชันเก่ามาก่อน ─────────────────────
 *    เลือกฟังก์ชัน  migrate  แล้วกด Run หนึ่งครั้ง
 *    (เพิ่มคอลัมน์ใหม่ให้ครบ โดยข้อมูลเดิมไม่หาย)
 *
 * ── เมื่อแก้ไขโค้ดนี้ ทุกครั้ง ───────────────────────────
 *    Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy
 *    (ถ้าไม่ทำ ลิงก์เดิมจะยังใช้โค้ดเก่า)
 *
 * ── เมื่อมีปัญหา ─────────────────────────────────────────
 *    เลือกฟังก์ชัน  showConfig  แล้วกด Run เพื่อดูสถานะการตั้งค่า
 *************************************************************/

var P = PropertiesService.getScriptProperties();

/* ---------- ตั้งค่าเริ่มต้น (รันครั้งเดียว) ----------
 * มีระบบป้องกันการสร้างซ้ำ หากเคยติดตั้งแล้วจะไม่สร้างไฟล์ใหม่
 * เพื่อไม่ให้ข้อมูลเดิมหาย
 */
function setup() {
  var sid = P.getProperty('SHEET_ID');
  if (sid) {
    try {
      var old = SpreadsheetApp.openById(sid);
      Logger.log('⚠️ ติดตั้งไว้แล้ว ไม่ได้สร้างใหม่ เพื่อป้องกันข้อมูลเดิมหาย');
      Logger.log('Sheet ปัจจุบัน: ' + old.getName() + ' → ' + old.getUrl());
      Logger.log('หากต้องการตรวจสอบสถานะ ให้รันฟังก์ชัน showConfig');
      Logger.log('หากต้องการสร้างใหม่จริงๆ ให้รันฟังก์ชัน resetSetup ก่อน');
      return;
    } catch (e) {
      Logger.log('พบการตั้งค่าเดิมแต่เปิดไฟล์ไม่ได้ กำลังสร้างใหม่…');
    }
  }
  var ss = SpreadsheetApp.create('CULAC-Agenda-DB');
  var mSheet = ss.getSheets()[0].setName('Meetings');
  mSheet.appendRow(['id','round','year','date','time','venue','prev','status','createdAt']);
  var iSheet = ss.insertSheet('Items');
  iSheet.appendRow(['id','ref','meetingId','dept','proposer','contact','type','title','detail','status','order','category','urgency','amount','deadline','resolution','resolutionNote','owner','files','createdAt']);

  var folder = DriveApp.createFolder('CULAC-Agenda-Files');
  P.setProperty('SHEET_ID', ss.getId());
  P.setProperty('FOLDER_ID', folder.getId());
  Logger.log('เสร็จสิ้น: Sheet=' + ss.getUrl() + ' | Folder=' + folder.getUrl());
}

/* ---------- ล้างการตั้งค่า (ใช้เมื่อต้องการเริ่มใหม่จริงๆ) ----------
 * ไม่ได้ลบ Google Sheet หรือไฟล์เดิม เพียงยกเลิกการเชื่อมโยงเท่านั้น
 */
function resetSetup() {
  P.deleteProperty('SHEET_ID');
  P.deleteProperty('FOLDER_ID');
  Logger.log('ล้างการตั้งค่าแล้ว — รัน setup ได้อีกครั้ง (ไฟล์เดิมยังอยู่ใน Drive)');
}

/* ---------- อัปเกรดฐานข้อมูลเดิม (รันครั้งเดียวหลังอัปเดตโค้ด) ----------
 * ใช้เมื่อเคย setup() ไปแล้วก่อนมีระบบจำแนกเรื่อง/รายงาน
 * จะเพิ่มคอลัมน์ใหม่ให้ชีต Items โดยไม่กระทบข้อมูลเดิม
 */
function migrate() {
  var sh = sh_('Items');
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var needed = ['category','urgency','amount','deadline','resolution','resolutionNote','owner'];
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
  Logger.log(added.length ? 'อัปเดต: ' + added.join(', ') : 'ฐานข้อมูลเป็นเวอร์ชันล่าสุดอยู่แล้ว');
}

/* ---------- ตรวจสอบการตั้งค่า (รันเมื่อมีปัญหา) ----------
 * เลือกฟังก์ชันนี้แล้วกด Run จากนั้นดูผลที่ Execution log
 */
function showConfig() {
  var sid = P.getProperty('SHEET_ID');
  var fid = P.getProperty('FOLDER_ID');
  if (!sid || !fid) {
    Logger.log('❌ ยังไม่ได้ตั้งค่า — ให้รันฟังก์ชัน setup() หนึ่งครั้งก่อน');
    Logger.log('SHEET_ID = ' + sid + ' | FOLDER_ID = ' + fid);
    return;
  }
  try {
    var ss = SpreadsheetApp.openById(sid);
    var fd = DriveApp.getFolderById(fid);
    Logger.log('✅ ตั้งค่าเรียบร้อย');
    Logger.log('Sheet : ' + ss.getName() + ' → ' + ss.getUrl());
    Logger.log('Folder: ' + fd.getName() + ' → ' + fd.getUrl());
    Logger.log('ชีตที่มี: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
  } catch (e) {
    Logger.log('❌ เปิดไฟล์ไม่ได้: ' + e);
    Logger.log('ให้รัน setup() ใหม่ หรือใช้ useExisting() เพื่อชี้ไปยังไฟล์เดิม');
  }
}

/* ---------- เชื่อมกลับไปยัง Sheet เดิม (กรณีเคยมีข้อมูลอยู่แล้ว) ----------
 * ใช้เมื่อเคยรัน setup() ไปแล้วในโปรเจกต์เก่า และไม่อยากให้สร้าง Sheet ใหม่
 * วิธีใช้: แก้ค่า 2 บรรทัดด้านล่างให้เป็น ID ของไฟล์เดิม แล้วกด Run
 * (ID คือข้อความยาวๆ ใน URL ระหว่าง /d/ กับ /edit)
 */
function useExisting() {
  var SHEET_ID  = 'วาง_ID_ของ_Google_Sheet_ที่นี่';
  var FOLDER_ID = 'วาง_ID_ของโฟลเดอร์ที่นี่';

  P.setProperty('SHEET_ID', SHEET_ID);
  P.setProperty('FOLDER_ID', FOLDER_ID);
  Logger.log('บันทึกแล้ว — ให้รัน showConfig() เพื่อตรวจสอบ');
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
      return json_({ ok:true, meetings: readSheet_('Meetings'), items: readSheet_('Items'),
        notify: { mode: P.getProperty('NOTIFY_MODE') || 'urgent',
                  emails: P.getProperty('NOTIFY_EMAILS') || '',
                  site: P.getProperty('SITE_LINK') || '' } });
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
      var noti = notifyNewItem_(item);   // แจ้งเตือนอีเมล (ถ้าเปิดใช้)
      return json_({ ok:true, item:item, notified:noti });
    }

    if (a==='updateItem'){
      var it = body.item, row = findRow_('Items', it.id);
      if (row<0) return json_({ ok:false, error:'ไม่พบเรื่อง' });
      // คงไฟล์เดิมไว้เสมอ ถ้า payload ไม่ได้ส่งไฟล์ใหม่มา หรือส่งมาเป็นรายการว่าง
      // (ป้องกันไฟล์แนบหายจากการแก้ไขข้อมูลเรื่อง)
      if (!it.files || (Object.prototype.toString.call(it.files) === '[object Array]' && it.files.length === 0)) {
        var existing = readSheet_('Items').filter(function(x){return x.id===it.id;})[0];
        it.files = (existing && existing.files) ? existing.files : [];
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

    if (a==='saveNotify'){
      P.setProperty('NOTIFY_MODE', body.mode || 'urgent');
      P.setProperty('NOTIFY_EMAILS', body.emails || '');
      P.setProperty('SITE_LINK', body.site || '');
      return json_({ ok:true });
    }

    if (a==='testNotify'){
      var to = notifyList_();
      if (!to.length) return json_({ ok:false, error:'ยังไม่ได้ตั้งอีเมลผู้รับแจ้งเตือน' });
      MailApp.sendEmail({
        to: to.join(','),
        subject: 'ทดสอบระบบแจ้งเตือน — ระบบวาระการประชุม CULAC',
        htmlBody: '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:15px">' +
          '<p>นี่คืออีเมลทดสอบจากระบบส่งเรื่องเข้าที่ประชุมคณะกรรมการบริหาร</p>' +
          '<p>หากได้รับอีเมลฉบับนี้ แสดงว่าการตั้งค่าแจ้งเตือนถูกต้องแล้ว</p></div>',
        name: 'ระบบวาระการประชุม CULAC'
      });
      return json_({ ok:true, sent: to.length });
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


function row_(label, value){
  return '<tr>' +
    '<td style="padding:7px 0;color:#69727f;width:110px;vertical-align:top">' + label + '</td>' +
    '<td style="padding:7px 0;font-weight:bold;color:#152744">' + escapeHtml_(String(value)) + '</td></tr>';
}

function escapeHtml_(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}


/* ---------- แจ้งเตือนอีเมลเมื่อมีเรื่องส่งเข้ามา ---------- */
function notifyList_(){
  var raw = P.getProperty('NOTIFY_EMAILS') || '';
  return raw.split(/[,;\n]/)
            .map(function(e){ return e.trim().replace(/^\.+/, ''); })
            .filter(function(e){ return e && e.indexOf('@') > 0; });
}

function notifyNewItem_(item){
  try {
    var mode = P.getProperty('NOTIFY_MODE') || 'urgent';
    if (mode === 'off') return 'off';
    if (mode === 'urgent' && item.type !== 'circulate') return 'skip';

    var to = notifyList_();
    if (!to.length) return 'no-recipient';

    var urgent = (item.type === 'circulate');
    var meeting = '';
    try {
      var ms = readSheet_('Meetings');
      for (var i = 0; i < ms.length; i++)
        if (ms[i].id === item.meetingId) meeting = 'ครั้งที่ ' + ms[i].round + '/' + ms[i].year;
    } catch (e) {}

    var typeLabel = { inform:'แจ้งเพื่อทราบ', consider:'เสนอเพื่อพิจารณา',
                      other:'เรื่องอื่นๆ', circulate:'เรื่องด่วน — ขอเวียนมติ' }[item.type] || item.type;

    var files = '';
    if (item.files && item.files.length){
      files = item.files.map(function(f){
        return '<a href="' + f.url + '" style="color:#2c5aa0">' + escapeHtml_(f.name) + '</a>';
      }).join('<br>');
    } else {
      files = '<span style="color:#69727f">ไม่มีไฟล์แนบ</span>';
    }

    var site = P.getProperty('SITE_LINK') || '';
    var bar  = urgent ? '#b23b3b' : '#152744';
    var head = urgent ? 'เรื่องด่วน — ขอเวียนมติ' : 'มีเรื่องเสนอเข้าที่ประชุมใหม่';

    var html =
      '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:15px;color:#232b36;max-width:600px">' +
        '<div style="background:' + bar + ';color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">' +
          '<div style="font-size:18px;font-weight:bold">' + head + '</div>' +
          '<div style="font-size:13px;opacity:.85">ระบบวาระการประชุมคณะกรรมการบริหาร ศูนย์สัตว์ทดลอง</div>' +
        '</div>' +
        '<div style="border:1px solid #e3e6ec;border-top:none;border-radius:0 0 10px 10px;padding:20px">' +
          '<table style="width:100%;border-collapse:collapse;font-size:15px">' +
            row_('เรื่อง', item.title) +
            row_('งานที่เสนอ', item.dept) +
            row_('ผู้เสนอ', item.proposer || '-') +
            row_('ประเภท', typeLabel) +
            (meeting ? row_('การประชุม', meeting) : '') +
          '</table>' +
          (item.detail ? '<p style="margin-top:12px;color:#232b36">' + escapeHtml_(item.detail) + '</p>' : '') +
          '<p style="margin:16px 0 4px;font-weight:bold;color:#152744">เอกสารแนบ</p>' + files +
          (urgent ? '<p style="margin-top:18px;padding:12px;background:#fdeaea;border-radius:8px;color:#a33232">' +
             'เรื่องนี้เป็นเรื่องด่วนที่ขอเวียนมติ กรุณาพิจารณาดำเนินการโดยเร็ว</p>' : '') +
          (site ? '<div style="margin-top:18px;text-align:center">' +
             '<a href="' + site + '" style="display:inline-block;background:#c8671a;color:#fff;padding:10px 22px;' +
             'border-radius:8px;text-decoration:none;font-weight:bold">เปิดระบบวาระการประชุม</a></div>' : '') +
        '</div></div>';

    MailApp.sendEmail({
      to: to.join(','),
      subject: (urgent ? '[ด่วน] ' : '[เรื่องใหม่] ') + item.title,
      htmlBody: html,
      name: 'ระบบวาระการประชุม CULAC'
    });
    return 'sent';
  } catch (err) {
    return 'error: ' + err;   // ไม่ให้การแจ้งเตือนล้มเหลวกระทบการบันทึกเรื่อง
  }
}
