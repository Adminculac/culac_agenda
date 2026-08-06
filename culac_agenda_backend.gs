/*************************************************************************
 * CULAC — ระบบส่งเรื่องประชุมคณะกรรมการบริหาร  (Google Apps Script Backend)
 * เขียนใหม่ทั้งไฟล์ ให้เข้ากันกับ culac_agenda.html เวอร์ชันล่าสุด
 *
 * จุดเด่นของเวอร์ชันนี้ :
 *  - อ่าน/เขียนข้อมูล "ตามชื่อหัวคอลัมน์" (ไม่ยึดลำดับคอลัมน์) → ยืดหยุ่น ไม่พังง่าย
 *  - เพิ่มฟิลด์ใหม่ได้ทันทีแค่เติมชื่อใน ITEM_FIELDS แล้วรัน migrate (รองรับ proposerSig แล้ว)
 *  - ไฟล์แนบอัปโหลดขึ้น Google Drive แล้วเก็บเป็นลิงก์ (ไม่ทำให้ชีตบวม)
 *  - มี login / audit log ฝั่งเซิร์ฟเวอร์เตรียมไว้ (พร้อมต่อกับ HTML ภายหลัง)
 *
 * ---------- วิธีติดตั้ง (ทำครั้งเดียว) ----------
 *  1) เปิด script.google.com → New project → วางโค้ดนี้ทั้งหมด
 *  2) เมนู Run → เลือกฟังก์ชัน  setup  → กด Run (อนุญาตสิทธิ์ครั้งแรก)
 *     - จะสร้างสเปรดชีต "CULAC-Agenda-DB" + โฟลเดอร์ไฟล์แนบให้อัตโนมัติ
 *  3) Deploy → New deployment → ประเภท Web app
 *       Execute as = Me , Who has access = Anyone
 *  4) คัดลอก URL ที่ลงท้าย /exec ไปวางในหน้า "ตั้งค่าระบบ" ของเว็บ
 *
 *  * ถ้าอยากใช้สเปรดชีตเดิมที่มีอยู่แล้ว : ใส่รหัสชีตในตัวแปร SHEET_ID ด้านล่าง
 *    (โค้ดจะเติมคอลัมน์ที่ขาดให้เอง โดยไม่ลบข้อมูลเดิม)
 *************************************************************************/

/* ====================== ตั้งค่า ====================== */
var SHEET_ID   = '';              // เว้นว่าง = สร้าง/หาชีตชื่อ CULAC-Agenda-DB ให้อัตโนมัติ
var DB_NAME    = 'CULAC-Agenda-DB';
var FILES_FOLDER_NAME = 'CULAC-Agenda-Files';
var DEFAULT_ADMIN_PASS = '1234';
var DEFAULT_DIR_PASS   = 'culac1';

/* คอลัมน์ของแต่ละชีต — เพิ่มฟิลด์ใหม่ในอนาคตแค่เติมชื่อที่นี่แล้วรัน migrate */
var ITEM_FIELDS = ['id','ref','meetingId','dept','proposer','contact','type','title','detail',
  'status','order','category','urgency','amount','resolution','resolutionNote','owner','deadline',
  'memoNo','memoDate','proposerPos','proposerSig','memoTables',
  'approval','approvedAt','approvedNote','approvedSig','files','createdAt'];
var MEETING_FIELDS = ['id','round','year','date','time','venue','createdAt'];
var JSON_FIELDS = {items:['memoTables','files'], meetings:[]};   // ฟิลด์ที่เก็บเป็น JSON ในเซลล์
var NUM_FIELDS  = {items:['order'], meetings:[]};                // ฟิลด์ที่แปลงเป็นตัวเลข

/* ====================== Entry points ====================== */
function doGet(e){
  var action = (e && e.parameter && e.parameter.action) || 'list';
  try{
    if(action === 'list'){
      return json_({
        ok:true,
        meetings: readAll_('Meetings','meetings'),
        items:    readAll_('Items','items'),
        branding: getBranding_(),
        notify:   getNotify_()
      });
    }
    return json_({ok:false, error:'unknown action: '+action});
  }catch(err){ return json_({ok:false, error:String(err)}); }
}

function doPost(e){
  var p = {};
  try{ p = JSON.parse(e.postData.contents); }catch(err){ return json_({ok:false, error:'bad json'}); }
  var action = p.action || '';
  try{
    switch(action){
      case 'submit':        return json_(submit_(p));
      case 'updateItem':    return json_(updateItem_(p.item));
      case 'deleteItem':    return json_(deleteById_('Items', p.id));
      case 'saveMeeting':   return json_(saveMeeting_(p.meeting));
      case 'deleteMeeting': return json_(deleteMeeting_(p.id));
      case 'reorder':       return json_(reorder_(p.orders));
      case 'approveItem':   return json_(approveItem_(p));
      case 'saveBranding':  return json_(saveBranding_(p));
      case 'saveOrg':       return json_(saveBranding_(p));   // เก็บที่ Settings ชุดเดียวกัน
      case 'saveNotify':    return json_(saveNotify_(p));
      case 'testNotify':    return json_(testNotify_());
      case 'login':         return json_(login_(p));          // ตรวจรหัสฝั่งเซิร์ฟเวอร์ (hash)
      case 'changePass':    return json_(changePass_(p));     // เจ้าของเปลี่ยนรหัสตัวเอง
      case 'adminResetDir': return json_(adminResetDir_(p));  // ผู้ดูแลรีเซ็ตรหัส ผอ. กรณีลืม
      case 'logEvent':      return json_(logEvent_(p));        // เผื่อ audit log รวมศูนย์
      default:              return json_({ok:false, error:'unknown action: '+action});
    }
  }catch(err){ return json_({ok:false, error:String(err)}); }
}

/* ====================== การกระทำกับข้อมูล ====================== */
function submit_(p){
  var item = p.item || {};
  // อัปโหลดไฟล์แนบขึ้น Drive แล้วเก็บเป็นลิงก์
  var files = [];
  (p.files || []).forEach(function(f){
    var saved = saveFile_(f);
    if(saved) files.push(saved);
  });
  item.files = files;
  if(!item.createdAt) item.createdAt = new Date().toISOString();
  writeRow_('Items','items', item);
  logEvent_({actor: actorOf_(item), action:'ส่งเรื่อง', ref:item.ref, title:item.title, detail:item.dept||''});
  notifyNewItem_(item);
  return {ok:true, item:item};
}

function updateItem_(item){
  if(!item || !item.id) return {ok:false, error:'no id'};
  writeRow_('Items','items', item);
  return {ok:true};
}

function approveItem_(p){
  var sh = getSheet_('Items', ITEM_FIELDS);
  var row = findRow_(sh, 'id', p.id);
  if(row < 0) return {ok:false, error:'ไม่พบเรื่องนี้'};
  setCell_(sh, row, 'approval', p.approval || '');
  setCell_(sh, row, 'approvedAt', new Date().toISOString());
  setCell_(sh, row, 'approvedNote', p.note || '');
  setCell_(sh, row, 'approvedSig', (p.approval === 'approved') ? (p.sig || '') : '');
  logEvent_({actor:'ผู้อำนวยการ', action:(p.approval==='approved'?'อนุมัติ':'ไม่อนุมัติ'), ref:'', title:'', detail:p.note||''});
  return {ok:true};
}

function reorder_(orders){
  var sh = getSheet_('Items', ITEM_FIELDS);
  (orders || []).forEach(function(o){
    var row = findRow_(sh, 'id', o.id);
    if(row >= 0) setCell_(sh, row, 'order', o.order);
  });
  return {ok:true};
}

function saveMeeting_(m){
  if(!m || !m.id) return {ok:false, error:'no id'};
  if(!m.createdAt) m.createdAt = new Date().toISOString();
  writeRow_('Meetings','meetings', m);
  return {ok:true};
}

function deleteMeeting_(id){
  deleteById_('Meetings', id);
  // ลบเรื่องที่อยู่ในการประชุมนี้ทั้งหมด
  var sh = getSheet_('Items', ITEM_FIELDS);
  var data = sh.getDataRange().getValues();
  var head = data[0];
  var col = head.indexOf('meetingId');
  for(var r = data.length - 1; r >= 1; r--){
    if(col >= 0 && String(data[r][col]) === String(id)) sh.deleteRow(r + 1);
  }
  return {ok:true};
}

function deleteById_(sheetName, id){
  var sh = getSheet_(sheetName, sheetName === 'Items' ? ITEM_FIELDS : MEETING_FIELDS);
  var row = findRow_(sh, 'id', id);
  if(row >= 0) sh.deleteRow(row);
  return {ok:true};
}

/* ====================== Settings / Branding / Notify ====================== */
function saveBranding_(p){
  var keys = ['logo','emblem','dirName','dirPos','orgLine'];
  keys.forEach(function(k){ if(p[k] !== undefined) settingsSet_(k, p[k]); });
  return {ok:true};
}
function getBranding_(){
  return {
    logo:    settingsGet_('logo','') ,
    emblem:  settingsGet_('emblem',''),
    dirName: settingsGet_('dirName',''),
    dirPos:  settingsGet_('dirPos',''),
    orgLine: settingsGet_('orgLine','')
  };
}
function saveNotify_(p){
  settingsSet_('notifyMode',  p.mode  || 'off');
  settingsSet_('notifyEmails',p.emails|| '');
  settingsSet_('notifySite',  p.site  || '');
  return {ok:true};
}
function getNotify_(){
  return {
    mode:  settingsGet_('notifyMode','off'),
    emails:settingsGet_('notifyEmails',''),
    site:  settingsGet_('notifySite','')
  };
}
function testNotify_(){
  var to = settingsGet_('notifyEmails','');
  if(!to || to.indexOf('@') < 0) return {ok:false, error:'ยังไม่ได้ตั้งอีเมลผู้รับ'};
  MailApp.sendEmail(to, '[ทดสอบ] ระบบส่งเรื่องประชุม CULAC',
    'นี่คืออีเมลทดสอบจากระบบส่งเรื่องประชุมคณะกรรมการบริหาร CULAC\nหากได้รับแสดงว่าการแจ้งเตือนทำงานปกติ');
  return {ok:true};
}
function notifyNewItem_(item){
  try{
    if(settingsGet_('notifyMode','off') === 'off') return;
    var to = settingsGet_('notifyEmails','');
    if(!to || to.indexOf('@') < 0) return;
    var site = settingsGet_('notifySite','');
    MailApp.sendEmail(to, 'มีเรื่องใหม่เข้าระบบ : ' + (item.title || ''),
      'งาน : ' + (item.dept || '-') +
      '\nผู้เสนอ : ' + (item.proposer || '-') +
      '\nเรื่อง : ' + (item.title || '-') +
      (site ? ('\n\nเปิดระบบ : ' + site) : ''));
  }catch(err){ /* ไม่ให้การส่งเมลล้มเหลวไปกระทบการบันทึกข้อมูล */ }
}

/* ====================== Login / เปลี่ยนรหัส (เก็บเป็น hash เท่านั้น) ====================== */
/* salt ประจำการติดตั้ง — สุ่มครั้งเดียว เก็บใน Script Properties ทำให้ hash เดารื้อกลับไม่ได้ */
function salt_(){
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('PW_SALT');
  if(!s){ s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('PW_SALT', s); }
  return s;
}
function hash_(pw){
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + String(pw || ''));
  return raw.map(function(b){ return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function verify_(role, pw){
  var stored = settingsGet_(role === 'admin' ? 'adminHash' : 'dirHash', '');
  return stored !== '' && stored === hash_(pw);
}
function login_(p){
  var pass = (p.password || '').trim();
  if(pass && verify_('admin', pass))    return {ok:true, role:'admin'};
  if(pass && verify_('director', pass)) return {ok:true, role:'director'};
  return {ok:true, role:''};   // ไม่ตรง = ไม่มีสิทธิ์ (ไม่บอกว่ารหัสอะไรผิด)
}
/* เจ้าของรหัสเปลี่ยนรหัสตัวเอง : ต้องกรอกรหัสปัจจุบันให้ถูก */
function changePass_(p){
  var role = (p.role === 'admin') ? 'admin' : 'director';
  var current = (p.current || '').trim(), next = (p.next || '').trim();
  if(next.length < 4) return {ok:false, error:'รหัสใหม่ควรยาวอย่างน้อย 4 ตัวอักษร'};
  if(!verify_(role, current)) return {ok:false, error:'รหัสผ่านปัจจุบันไม่ถูกต้อง'};
  settingsSet_(role === 'admin' ? 'adminHash' : 'dirHash', hash_(next));
  logEvent_({actor: role === 'admin' ? 'ผู้ดูแลระบบ' : 'ผู้อำนวยการ', action:'เปลี่ยนรหัสผ่าน', ref:'', title:'', detail:''});
  return {ok:true};
}
/* ผู้ดูแลรีเซ็ตรหัส ผอ. กรณีลืม : ต้องยืนยันด้วยรหัสผู้ดูแล */
function adminResetDir_(p){
  var admin = (p.adminPassword || '').trim(), next = (p.next || '').trim();
  if(!verify_('admin', admin)) return {ok:false, error:'รหัสผู้ดูแลระบบไม่ถูกต้อง'};
  if(next.length < 4) return {ok:false, error:'รหัสใหม่ควรยาวอย่างน้อย 4 ตัวอักษร'};
  settingsSet_('dirHash', hash_(next));
  logEvent_({actor:'ผู้ดูแลระบบ', action:'รีเซ็ตรหัสผู้อำนวยการ', ref:'', title:'', detail:''});
  return {ok:true};
}
function logEvent_(ev){
  try{
    var sh = getSheet_('Audit', ['ts','actor','action','ref','title','detail']);
    sh.appendRow([ new Date().toISOString(), ev.actor||'', ev.action||'', ev.ref||'', ev.title||'', ev.detail||'' ]);
  }catch(err){}
  return {ok:true};
}
function actorOf_(item){ return (item && item.approval === 'none') ? 'ผู้เสนอ' : 'ผู้เสนอ'; }

/* ====================== ชั้นเข้าถึงสเปรดชีต (อ่าน/เขียนตามชื่อหัวคอลัมน์) ====================== */
function getSpreadsheet_(){
  var props = PropertiesService.getScriptProperties();
  var id = SHEET_ID || props.getProperty('SHEET_ID');
  if(id){ try{ return SpreadsheetApp.openById(id); }catch(e){} }
  // หาไฟล์เดิมตามชื่อ ถ้าไม่มีให้สร้างใหม่
  var it = DriveApp.getFilesByName(DB_NAME);
  var ss = it.hasNext() ? SpreadsheetApp.open(it.next()) : SpreadsheetApp.create(DB_NAME);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet_(name, fields){
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(fields);
    return sh;
  }
  // เติมหัวคอลัมน์ที่ขาด (migrate ในตัว) โดยไม่ลบของเดิม
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
  var added = false;
  (fields || []).forEach(function(f){
    if(head.indexOf(f) === -1){ head.push(f); added = true; }
  });
  if(added) sh.getRange(1,1,1,head.length).setValues([head]);
  return sh;
}

function headers_(sh){ return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String); }

function readAll_(name, kind){
  var sh = getSheet_(name, name === 'Items' ? ITEM_FIELDS : MEETING_FIELDS);
  var data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  var head = data[0].map(String);
  var jsonF = JSON_FIELDS[kind] || [];
  var numF  = NUM_FIELDS[kind] || [];
  var out = [];
  for(var r = 1; r < data.length; r++){
    var row = data[r];
    if(String(row[head.indexOf('id')] || '') === '') continue;   // ข้ามแถวว่าง
    var obj = {};
    for(var c = 0; c < head.length; c++){
      var key = head[c], val = row[c];
      if(jsonF.indexOf(key) >= 0){
        try{ obj[key] = val ? JSON.parse(val) : (key === 'files' || key === 'memoTables' ? [] : ''); }
        catch(e){ obj[key] = (key === 'files' || key === 'memoTables') ? [] : val; }
      }else if(numF.indexOf(key) >= 0){
        obj[key] = Number(val) || 0;
      }else{
        obj[key] = (val === null || val === undefined) ? '' : String(val);
      }
    }
    out.push(obj);
  }
  return out;
}

function writeRow_(name, kind, obj){
  var fields = (name === 'Items') ? ITEM_FIELDS : MEETING_FIELDS;
  var sh = getSheet_(name, fields);
  var head = headers_(sh);
  var jsonF = JSON_FIELDS[kind] || [];
  var rowArr = head.map(function(key){
    var v = obj[key];
    if(v === undefined || v === null) return '';
    if(jsonF.indexOf(key) >= 0) return JSON.stringify(v);
    return v;
  });
  var existing = findRow_(sh, 'id', obj.id);
  if(existing >= 0){
    sh.getRange(existing, 1, 1, head.length).setValues([rowArr]);
  }else{
    sh.appendRow(rowArr);
  }
  return sh;
}

/* คืนเลขแถวจริง (1-based รวมหัว) ที่ตรง id ; -1 ถ้าไม่พบ */
function findRow_(sh, idField, id){
  var data = sh.getDataRange().getValues();
  if(data.length < 2) return -1;
  var col = data[0].map(String).indexOf(idField);
  if(col < 0) return -1;
  for(var r = 1; r < data.length; r++){
    if(String(data[r][col]) === String(id)) return r + 1;
  }
  return -1;
}

function setCell_(sh, row, field, value){
  var col = headers_(sh).indexOf(field);
  if(col < 0){                                   // ยังไม่มีคอลัมน์นี้ ให้เพิ่ม
    col = headers_(sh).length;
    sh.getRange(1, col + 1).setValue(field);
  }
  sh.getRange(row, col + 1).setValue(value);
}

/* ====================== Settings key/value ====================== */
function settingsSheet_(){ return getSheet_('Settings', ['key','value']); }
function settingsGet_(key, dflt){
  var sh = settingsSheet_();
  var data = sh.getDataRange().getValues();
  for(var r = 1; r < data.length; r++){ if(String(data[r][0]) === key) return String(data[r][1]); }
  return dflt;
}
function settingsSet_(key, value){
  var sh = settingsSheet_();
  var data = sh.getDataRange().getValues();
  for(var r = 1; r < data.length; r++){
    if(String(data[r][0]) === key){ sh.getRange(r + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

/* ====================== ไฟล์แนบ → Google Drive ====================== */
function filesFolder_(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FILES_FOLDER_ID');
  if(id){ try{ return DriveApp.getFolderById(id); }catch(e){} }
  var it = DriveApp.getFoldersByName(FILES_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(FILES_FOLDER_NAME);
  props.setProperty('FILES_FOLDER_ID', folder.getId());
  return folder;
}
function saveFile_(f){
  try{
    if(!f || !f.data) return null;
    var bytes = Utilities.base64Decode(f.data);
    var blob  = Utilities.newBlob(bytes, f.mime || 'application/octet-stream', f.name || 'file');
    var file  = filesFolder_().createFile(blob);
    try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
    return { id:file.getId(), name:f.name || file.getName(), size:file.getSize(), mime:f.mime || '', url:file.getUrl() };
  }catch(err){ return null; }
}

/* ====================== Utils ====================== */
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ====================== ติดตั้ง / ปรับปรุงโครงสร้าง (รันมือ) ====================== */
function setup(){ migrate(); }   // เรียกครั้งแรกหลังวางโค้ด

function migrate(){
  getSheet_('Meetings', MEETING_FIELDS);
  getSheet_('Items',    ITEM_FIELDS);
  getSheet_('Settings', ['key','value']);
  getSheet_('Audit',    ['ts','actor','action','ref','title','detail']);
  // ตั้งรหัสเริ่มต้น (เก็บเป็น hash) ถ้ายังไม่เคยตั้ง
  salt_();
  if(settingsGet_('adminHash','') === '') settingsSet_('adminHash', hash_(DEFAULT_ADMIN_PASS));
  if(settingsGet_('dirHash','')   === '') settingsSet_('dirHash',   hash_(DEFAULT_DIR_PASS));
  var ss = getSpreadsheet_();
  Logger.log('พร้อมใช้งาน · Spreadsheet ID = ' + ss.getId() + '  · URL = ' + ss.getUrl());
}
