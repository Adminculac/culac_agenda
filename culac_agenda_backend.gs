/*************************************************************
 * ระบบส่งเรื่องเข้าที่ประชุมคณะกรรมการบริหาร
 * ศูนย์สัตว์ทดลอง จุฬาลงกรณ์มหาวิทยาลัย
 * ─────────────────────────────────────────────────────────
 * ฐานข้อมูลกลาง (Google Apps Script)
 *
 *  ╔═══════════════════════════════════════════════════════╗
 *  ║  ติดตั้งครั้งแรก                                       ║
 *  ╚═══════════════════════════════════════════════════════╝
 *  1) เลือกฟังก์ชัน  setup  แล้วกด Run หนึ่งครั้ง
 *     (ระบบขออนุญาต → Review permissions → Advanced → Allow)
 *     จะได้ Google Sheet "CULAC-Agenda-DB"
 *     และโฟลเดอร์ Drive "CULAC-Agenda-Files"
 *
 *  2) Deploy → New deployment → เลือกชนิด Web app
 *        Execute as      : Me
 *        Who has access  : Anyone      ← สำคัญที่สุด ห้ามพลาด
 *
 *  3) คัดลอกลิงก์ที่ลงท้าย /exec
 *     ไปวางในหน้าเว็บ → ตั้งค่าระบบ → GAS Web App URL
 *
 *  ╔═══════════════════════════════════════════════════════╗
 *  ║  อัปเกรดจากเวอร์ชันเก่า (มีข้อมูลอยู่แล้ว)             ║
 *  ╚═══════════════════════════════════════════════════════╝
 *     เลือกฟังก์ชัน  migrate  แล้วกด Run หนึ่งครั้ง
 *     (เพิ่มคอลัมน์และชีตที่ขาด โดยข้อมูลเดิมไม่หาย)
 *
 *  ╔═══════════════════════════════════════════════════════╗
 *  ║  ทุกครั้งที่แก้โค้ดนี้                                  ║
 *  ╚═══════════════════════════════════════════════════════╝
 *     Deploy → Manage deployments → ✏️ Edit
 *            → Version: New version → Deploy
 *     (ถ้าไม่ทำ ลิงก์เดิมจะยังใช้โค้ดเก่า)
 *
 *  ╔═══════════════════════════════════════════════════════╗
 *  ║  เมื่อมีปัญหา                                          ║
 *  ╚═══════════════════════════════════════════════════════╝
 *     เลือกฟังก์ชัน  showConfig  แล้วกด Run
 *     เพื่อดูว่าเชื่อมกับไฟล์ใดอยู่ และมีชีตครบหรือไม่
 *************************************************************/

var P = PropertiesService.getScriptProperties();

/* คอลัมน์มาตรฐานของแต่ละชีต */
var COL_MEETINGS = ['id','round','year','date','time','venue','prev','status','createdAt'];
var COL_ITEMS    = ['id','ref','meetingId','dept','proposer','contact','type','title','detail',
                    'status','order','category','urgency','amount','deadline',
                    'resolution','resolutionNote','owner','files','createdAt'];


/* ═══════════════════════════════════════════════════════════
   ส่วนที่ 1 · ติดตั้งและดูแลระบบ  (รันจากหน้า Apps Script)
   ═══════════════════════════════════════════════════════════ */

/** ติดตั้งครั้งแรก — มีระบบกันสร้างซ้ำเพื่อไม่ให้ข้อมูลเดิมหาย */
function setup() {
  var sid = P.getProperty('SHEET_ID');
  if (sid) {
    try {
      var exist = SpreadsheetApp.openById(sid);
      Logger.log('⚠️ ติดตั้งไว้แล้ว ไม่ได้สร้างใหม่ เพื่อป้องกันข้อมูลเดิมหาย');
      Logger.log('Sheet ปัจจุบัน: ' + exist.getName() + ' → ' + exist.getUrl());
      Logger.log('• ต้องการตรวจสอบสถานะ → รันฟังก์ชัน showConfig');
      Logger.log('• ต้องการเพิ่มชีต/คอลัมน์ที่ขาด → รันฟังก์ชัน migrate');
      return;
    } catch (e) {
      Logger.log('พบการตั้งค่าเดิมแต่เปิดไฟล์ไม่ได้ กำลังสร้างใหม่…');
    }
  }

  var ss = SpreadsheetApp.create('CULAC-Agenda-DB');
  ss.getSheets()[0].setName('Meetings').appendRow(COL_MEETINGS);
  ss.insertSheet('Items').appendRow(COL_ITEMS);
  ss.insertSheet('Settings').appendRow(['key','value']);

  var folder = DriveApp.createFolder('CULAC-Agenda-Files');

  P.setProperty('SHEET_ID',  ss.getId());
  P.setProperty('FOLDER_ID', folder.getId());

  Logger.log('✅ ติดตั้งเรียบร้อย');
  Logger.log('Sheet : ' + ss.getUrl());
  Logger.log('Folder: ' + folder.getUrl());
  Logger.log('ขั้นต่อไป → Deploy เป็น Web app แล้วตั้ง Who has access เป็น Anyone');
}

/** อัปเกรดฐานข้อมูลเดิม — เพิ่มชีตและคอลัมน์ที่ขาด ข้อมูลเดิมไม่หาย */
function migrate() {
  var err = configError_();
  if (err) { Logger.log('❌ ' + err); return; }

  var added = [];
  var ss = ss_();

  /* ชีตที่ต้องมี */
  if (!ss.getSheetByName('Settings')) {
    ss.insertSheet('Settings').appendRow(['key','value']);
    added.push('ชีต Settings (เก็บโลโก้ส่วนกลาง)');
  }
  if (!ss.getSheetByName('Meetings')) {
    ss.insertSheet('Meetings').appendRow(COL_MEETINGS);
    added.push('ชีต Meetings');
  }
  if (!ss.getSheetByName('Items')) {
    ss.insertSheet('Items').appendRow(COL_ITEMS);
    added.push('ชีต Items');
  }

  /* คอลัมน์ที่ต้องมีในชีต Items */
  var sh   = ss.getSheetByName('Items');
  var head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var need = ['category','urgency','amount','deadline','resolution','resolutionNote','owner'];

  need.forEach(function (col) {
    if (head.indexOf(col) !== -1) return;
    var filesIdx = head.indexOf('files');
    var at = (filesIdx === -1) ? sh.getLastColumn() + 1 : filesIdx + 1;
    sh.insertColumnBefore(at);
    sh.getRange(1, at).setValue(col);
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    added.push('คอลัมน์ ' + col);
  });

  /* ใส่ค่าเริ่มต้นให้แถวเดิมที่ยังว่าง */
  var last = sh.getLastRow();
  if (last > 1) {
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var uCol = head.indexOf('urgency') + 1;
    var rCol = head.indexOf('resolution') + 1;
    for (var r = 2; r <= last; r++) {
      if (uCol > 0 && !sh.getRange(r, uCol).getValue()) sh.getRange(r, uCol).setValue('normal');
      if (rCol > 0 && !sh.getRange(r, rCol).getValue()) sh.getRange(r, rCol).setValue('none');
    }
  }

  Logger.log(added.length ? '✅ อัปเดตแล้ว: ' + added.join(', ')
                          : '✅ ฐานข้อมูลเป็นเวอร์ชันล่าสุดอยู่แล้ว');
  Logger.log('อย่าลืม Deploy เวอร์ชันใหม่ด้วย');
}

/** ตรวจสถานะการตั้งค่า — ใช้เมื่อมีปัญหา */
function showConfig() {
  var sid = P.getProperty('SHEET_ID');
  var fid = P.getProperty('FOLDER_ID');

  if (!sid || !fid) {
    Logger.log('❌ ยังไม่ได้ตั้งค่าในโปรเจกต์นี้ — ให้รันฟังก์ชัน setup ก่อน');
    Logger.log('SHEET_ID = ' + sid + ' | FOLDER_ID = ' + fid);
    return;
  }
  try {
    var ss = SpreadsheetApp.openById(sid);
    var fd = DriveApp.getFolderById(fid);
    var names = ss.getSheets().map(function (s) { return s.getName(); });

    Logger.log('✅ ตั้งค่าเรียบร้อย');
    Logger.log('Sheet : ' + ss.getName() + ' → ' + ss.getUrl());
    Logger.log('Folder: ' + fd.getName() + ' → ' + fd.getUrl());
    Logger.log('ชีตที่มี: ' + names.join(', '));

    if (names.indexOf('Settings') === -1)
      Logger.log('⚠️ ยังไม่มีชีต Settings → ให้รันฟังก์ชัน migrate');

    Logger.log('การประชุม ' + Math.max(ss.getSheetByName('Meetings').getLastRow() - 1, 0) + ' รอบ · ' +
               'เรื่องที่เสนอ ' + Math.max(ss.getSheetByName('Items').getLastRow() - 1, 0) + ' รายการ');
  } catch (e) {
    Logger.log('❌ เปิดไฟล์ไม่ได้ (อาจถูกลบหรือย้าย): ' + e);
    Logger.log('แก้ได้โดยรัน resetSetup แล้วรัน setup ใหม่ หรือใช้ useExisting ชี้ไปไฟล์เดิม');
  }
}

/** ล้างการเชื่อมโยง (ไม่ได้ลบไฟล์ใน Drive) */
function resetSetup() {
  P.deleteProperty('SHEET_ID');
  P.deleteProperty('FOLDER_ID');
  Logger.log('ล้างการตั้งค่าแล้ว — รัน setup ได้อีกครั้ง (ไฟล์เดิมยังอยู่ใน Drive)');
}

/** ชี้กลับไปยังไฟล์เดิมที่มีข้อมูลอยู่แล้ว
 *  วิธีหา ID: เปิดไฟล์ใน Drive แล้วดู URL ส่วนที่อยู่ระหว่าง /d/ กับ /edit */
function useExisting() {
  var SHEET_ID  = 'วาง_ID_ของ_Google_Sheet_ที่นี่';
  var FOLDER_ID = 'วาง_ID_ของโฟลเดอร์ที่นี่';

  P.setProperty('SHEET_ID',  SHEET_ID);
  P.setProperty('FOLDER_ID', FOLDER_ID);
  Logger.log('บันทึกแล้ว — รัน showConfig เพื่อตรวจสอบ');
}


/* ═══════════════════════════════════════════════════════════
   ส่วนที่ 2 · เครื่องมือภายใน
   ═══════════════════════════════════════════════════════════ */

function ss_()      { return SpreadsheetApp.openById(P.getProperty('SHEET_ID')); }
function sh_(name)  { return ss_().getSheetByName(name); }
function folder_()  { return DriveApp.getFolderById(P.getProperty('FOLDER_ID')); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/** ตรวจว่าตั้งค่าแล้วหรือยัง คืนข้อความภาษาไทยแทน error ดิบ */
function configError_() {
  var sid = P.getProperty('SHEET_ID');
  if (!sid) {
    return 'ยังไม่ได้ตั้งค่าฐานข้อมูลในโปรเจกต์ Apps Script นี้ — ' +
           'ให้เปิด Apps Script เลือกฟังก์ชัน setup แล้วกด Run หนึ่งครั้ง ' +
           'จากนั้น Deploy > Manage deployments > Edit > Version: New version';
  }
  try {
    SpreadsheetApp.openById(sid);
  } catch (e) {
    return 'เปิดไฟล์ฐานข้อมูลไม่ได้ (อาจถูกลบหรือย้าย) — ' +
           'ให้รันฟังก์ชัน showConfig เพื่อตรวจสอบ';
  }
  return null;
}

/** อ่านทั้งชีตออกมาเป็นรายการ object */
function readSheet_(name) {
  var sh = sh_(name);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  var head = vals[0], out = [];
  for (var r = 1; r < vals.length; r++) {
    var o = {};
    for (var c = 0; c < head.length; c++) o[head[c]] = vals[r][c];
    if (name === 'Items') {
      try { o.files = o.files ? JSON.parse(o.files) : []; } catch (e) { o.files = []; }
    }
    out.push(o);
  }
  return out;
}

/** หาแถวจาก id (คืนเลขแถวจริง ไม่พบคืน -1) */
function findRow_(name, id) {
  var sh = sh_(name);
  if (!sh || sh.getLastRow() < 2) return -1;
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var r = 1; r < ids.length; r++) {
    if (String(ids[r][0]) === String(id)) return r + 1;
  }
  return -1;
}

/** แปลง object เป็นแถวตามลำดับหัวคอลัมน์จริงของชีต */
function rowFrom_(name, obj) {
  var sh = sh_(name);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return head.map(function (k) {
    if (k === 'files') return JSON.stringify(obj.files || []);
    return (obj[k] !== null && obj[k] !== undefined) ? obj[k] : '';
  });
}

/* ---- ตั้งค่าส่วนกลาง เช่น โลโก้ (เก็บในชีต Settings) ---- */
function settingsSheet_() {
  var st = sh_('Settings');
  if (!st) { st = ss_().insertSheet('Settings'); st.appendRow(['key','value']); }
  return st;
}
function getSetting_(key) {
  var st = settingsSheet_(), last = st.getLastRow();
  if (last < 2) return '';
  var rows = st.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) return String(rows[i][1] || '');
  }
  return '';
}
function setSetting_(key, value) {
  var st = settingsSheet_(), last = st.getLastRow();
  if (last >= 2) {
    var keys = st.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === key) { st.getRange(i + 2, 2).setValue(value); return; }
    }
  }
  st.appendRow([key, value]);
}

/** บันทึกไฟล์แนบลง Drive แล้วคืนรายการลิงก์ */
function saveFiles_(files, prefix) {
  var out = [];
  if (!files || !files.length) return out;

  var folder = folder_();
  files.forEach(function (f) {
    var bytes = Utilities.base64Decode(f.data);
    var blob  = Utilities.newBlob(bytes, f.mime || 'application/octet-stream',
                                  (prefix || 'file') + '_' + f.name);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    out.push({ id: file.getId(), name: f.name, size: bytes.length, url: file.getUrl() });
  });
  return out;
}


/* ═══════════════════════════════════════════════════════════
   ส่วนที่ 3 · รับคำสั่งจากหน้าเว็บ
   ═══════════════════════════════════════════════════════════ */

/** อ่านข้อมูลทั้งหมด */
function doGet(e) {
  try {
    var err = configError_();
    if (err) return json_({ ok: false, error: err });

    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action !== 'list') return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + action });

    var branding = {};
    try { branding = { logo: getSetting_('logo'), emblem: getSetting_('emblem') }; } catch (e2) {}

    return json_({
      ok: true,
      meetings: readSheet_('Meetings'),
      items:    readSheet_('Items'),
      branding: branding,
      notify: {
        mode:   P.getProperty('NOTIFY_MODE')   || 'urgent',
        emails: P.getProperty('NOTIFY_EMAILS') || '',
        site:   P.getProperty('SITE_LINK')     || ''
      }
    });
  } catch (err2) {
    return json_({ ok: false, error: String(err2) });
  }
}

/** บันทึก แก้ไข ลบ และตั้งค่า */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var a = body.action;

    var err = configError_();
    if (err) return json_({ ok: false, error: err });

    /* ---- ส่งเรื่องเข้าที่ประชุม ---- */
    if (a === 'submit') {
      var item = body.item;
      item.files = saveFiles_(body.files, item.ref || item.id);
      sh_('Items').appendRow(rowFrom_('Items', item));
      return json_({ ok: true, item: item, notified: notifyNewItem_(item) });
    }

    /* ---- แก้ไขเรื่อง ---- */
    if (a === 'updateItem') {
      var it  = body.item;
      var row = findRow_('Items', it.id);
      if (row < 0) return json_({ ok: false, error: 'ไม่พบเรื่องที่ต้องการแก้ไข' });

      /* คงไฟล์แนบเดิมไว้เสมอ ถ้าไม่ได้ส่งไฟล์ใหม่มา
         (ป้องกันไฟล์แนบหายจากการแก้ไขข้อมูล) */
      var isEmptyArr = (Object.prototype.toString.call(it.files) === '[object Array]' && it.files.length === 0);
      if (!it.files || isEmptyArr) {
        var old = readSheet_('Items').filter(function (x) { return x.id === it.id; })[0];
        it.files = (old && old.files) ? old.files : [];
      }

      sh_('Items').getRange(row, 1, 1, sh_('Items').getLastColumn())
                  .setValues([rowFrom_('Items', it)]);
      return json_({ ok: true });
    }

    /* ---- ลบเรื่อง ---- */
    if (a === 'deleteItem') {
      var r1 = findRow_('Items', body.id);
      if (r1 > 0) sh_('Items').deleteRow(r1);
      return json_({ ok: true });
    }

    /* ---- บันทึก/แก้ไขการประชุม ---- */
    if (a === 'saveMeeting') {
      var m  = body.meeting;
      var mr = findRow_('Meetings', m.id);
      if (mr < 0) sh_('Meetings').appendRow(rowFrom_('Meetings', m));
      else sh_('Meetings').getRange(mr, 1, 1, sh_('Meetings').getLastColumn())
                          .setValues([rowFrom_('Meetings', m)]);
      return json_({ ok: true });
    }

    /* ---- ลบการประชุม พร้อมเรื่องในรอบนั้น ---- */
    if (a === 'deleteMeeting') {
      var mrow = findRow_('Meetings', body.id);
      if (mrow > 0) sh_('Meetings').deleteRow(mrow);

      var items = readSheet_('Items');
      for (var i = 0; i < items.length; i++) {
        if (items[i].meetingId === body.id) {
          var rr = findRow_('Items', items[i].id);
          if (rr > 0) sh_('Items').deleteRow(rr);
        }
      }
      return json_({ ok: true });
    }

    /* ---- จัดลำดับวาระ ---- */
    if (a === 'reorder') {
      var sh = sh_('Items'), data = sh.getDataRange().getValues(), head = data[0];
      var idCol = head.indexOf('id'), ordCol = head.indexOf('order');
      var map = {};
      (body.orders || []).forEach(function (o) { map[o.id] = o.order; });

      for (var r2 = 1; r2 < data.length; r2++) {
        var id = data[r2][idCol];
        if (map[id] !== null && map[id] !== undefined) sh.getRange(r2 + 1, ordCol + 1).setValue(map[id]);
      }
      return json_({ ok: true });
    }

    /* ---- โลโก้ส่วนกลาง (ทุกเครื่องเห็นเหมือนกัน) ---- */
    if (a === 'saveBranding') {
      var MAXLEN = 45000;   // ขีดจำกัดของช่องใน Google Sheet
      if (body.logo !== undefined) {
        if (String(body.logo).length > MAXLEN)
          return json_({ ok: false, error: 'ไฟล์โลโก้ใหญ่เกินไป กรุณาใช้รูปขนาดเล็กลง' });
        setSetting_('logo', body.logo);
      }
      if (body.emblem !== undefined) {
        if (String(body.emblem).length > MAXLEN)
          return json_({ ok: false, error: 'ไฟล์ตราสัญลักษณ์ใหญ่เกินไป กรุณาใช้รูปขนาดเล็กลง' });
        setSetting_('emblem', body.emblem);
      }
      return json_({ ok: true });
    }

    /* ---- ตั้งค่าการแจ้งเตือน ---- */
    if (a === 'saveNotify') {
      P.setProperty('NOTIFY_MODE',   body.mode   || 'urgent');
      P.setProperty('NOTIFY_EMAILS', body.emails || '');
      P.setProperty('SITE_LINK',     body.site   || '');
      return json_({ ok: true });
    }

    /* ---- ส่งอีเมลทดสอบ ---- */
    if (a === 'testNotify') {
      var to = notifyList_();
      if (!to.length) return json_({ ok: false, error: 'ยังไม่ได้ตั้งอีเมลผู้รับแจ้งเตือน' });

      MailApp.sendEmail({
        to: to.join(','),
        subject: 'ทดสอบระบบแจ้งเตือน — ระบบวาระการประชุม CULAC',
        htmlBody: '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:15px">' +
                  '<p>นี่คืออีเมลทดสอบจากระบบส่งเรื่องเข้าที่ประชุมคณะกรรมการบริหาร</p>' +
                  '<p>หากได้รับอีเมลฉบับนี้ แสดงว่าการตั้งค่าแจ้งเตือนถูกต้องแล้ว</p></div>',
        name: 'ระบบวาระการประชุม CULAC'
      });
      return json_({ ok: true, sent: to.length });
    }

    return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + a });

  } catch (err3) {
    return json_({ ok: false, error: String(err3) });
  }
}


/* ═══════════════════════════════════════════════════════════
   ส่วนที่ 4 · แจ้งเตือนทางอีเมล
   ═══════════════════════════════════════════════════════════ */

function notifyList_() {
  var raw = P.getProperty('NOTIFY_EMAILS') || '';
  return raw.split(/[,;\n]/)
            .map(function (x) { return x.trim().replace(/^\.+/, ''); })
            .filter(function (x) { return x && x.indexOf('@') > 0; });
}

/** ส่งอีเมลเมื่อมีเรื่องเข้ามา (ไม่ให้ล้มเหลวกระทบการบันทึกเรื่อง) */
function notifyNewItem_(item) {
  try {
    var mode = P.getProperty('NOTIFY_MODE') || 'urgent';
    if (mode === 'off') return 'off';
    if (mode === 'urgent' && item.type !== 'circulate') return 'skip';

    var to = notifyList_();
    if (!to.length) return 'no-recipient';

    var urgent  = (item.type === 'circulate');
    var meeting = '';
    try {
      readSheet_('Meetings').forEach(function (m) {
        if (m.id === item.meetingId) meeting = 'ครั้งที่ ' + m.round + '/' + m.year;
      });
    } catch (e) {}

    var typeLabel = ({
      inform:    'แจ้งเพื่อทราบ',
      consider:  'เสนอเพื่อพิจารณา',
      other:     'เรื่องอื่นๆ',
      circulate: 'เรื่องด่วน — ขอเวียนมติ'
    })[item.type] || item.type;

    var files = (item.files && item.files.length)
      ? item.files.map(function (f) {
          return '<a href="' + f.url + '" style="color:#2c5aa0">' + escapeHtml_(f.name) + '</a>';
        }).join('<br>')
      : '<span style="color:#69727f">ไม่มีไฟล์แนบ</span>';

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
            row_('เรื่อง',      item.title) +
            row_('งานที่เสนอ',  item.dept) +
            row_('ผู้เสนอ',     item.proposer || '-') +
            row_('ประเภท',      typeLabel) +
            (meeting ? row_('การประชุม', meeting) : '') +
          '</table>' +
          (item.detail ? '<p style="margin-top:12px">' + escapeHtml_(item.detail) + '</p>' : '') +
          '<p style="margin:16px 0 4px;font-weight:bold;color:#152744">เอกสารแนบ</p>' + files +
          (urgent
            ? '<p style="margin-top:18px;padding:12px;background:#fdeaea;border-radius:8px;color:#a33232">' +
              'เรื่องนี้เป็นเรื่องด่วนที่ขอเวียนมติ กรุณาพิจารณาดำเนินการโดยเร็ว</p>'
            : '') +
          (site
            ? '<div style="margin-top:18px;text-align:center">' +
              '<a href="' + site + '" style="display:inline-block;background:#c8671a;color:#fff;' +
              'padding:10px 22px;border-radius:8px;text-decoration:none;font-weight:bold">' +
              'เปิดระบบวาระการประชุม</a></div>'
            : '') +
        '</div></div>';

    MailApp.sendEmail({
      to: to.join(','),
      subject: (urgent ? '[ด่วน] ' : '[เรื่องใหม่] ') + item.title,
      htmlBody: html,
      name: 'ระบบวาระการประชุม CULAC'
    });
    return 'sent';

  } catch (err) {
    return 'error: ' + err;
  }
}

function row_(label, value) {
  return '<tr>' +
    '<td style="padding:7px 0;color:#69727f;width:110px;vertical-align:top">' + label + '</td>' +
    '<td style="padding:7px 0;font-weight:bold;color:#152744">' + escapeHtml_(String(value)) + '</td>' +
  '</tr>';
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
