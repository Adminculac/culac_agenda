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
        ids.forEach(function(id){
          rev++;
          var v = projChanges[id];
          var del = (v===null || v===undefined) ? 1 : 0;
          var json = del ? '' : JSON.stringify(v);
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
      // คืนการเปลี่ยนแปลงตั้งแต่ since (รวมของคนอื่นที่ไคลเอนต์ยังไม่เห็น) เพื่อรวบ pull+push เป็นรอบเดียว
      if(since!=null){ return jsonOut_(pullSince_(s, since)); }
      return jsonOut_({ ok:true, rev:rev });
    } finally { lock.releaseLock(); }
  }catch(err){ return jsonOut_({ ok:false, error:String(err) }); }
}
