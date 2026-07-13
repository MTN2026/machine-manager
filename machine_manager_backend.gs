/**
 * ระบบจัดการเครื่องจักร — Backend (Google Apps Script)
 * ==========================================================================
 * ไฟล์นี้ทำหน้าที่ 3 อย่าง:
 *
 *  1) เป็น "ฐานข้อมูลกลาง" ให้เว็บแอปทุกอุปกรณ์อ่าน-เขียนข้อมูลชุดเดียวกันจริงๆ
 *     (ไม่ใช่แค่ export ทางเดียวเหมือนเดิม) โดยเก็บเป็น key-value ไว้ในชีทชื่อ
 *     "AppData" — คีย์เดียวกับที่แอปใช้เก็บใน localStorage เดิม (machines_list,
 *     cranes_list, parts_catalog, machmaint_xxx ฯลฯ)
 *
 *  2) เก็บรูปภาพเครื่องจักร/เครน/ยานพาหนะ/เครื่องมือ ไว้ใน Google Drive (โฟลเดอร์
 *     "MachineManagerPhotos") แทนการยัด base64 ลงเซลล์ Sheet ที่มีขีดจำกัดขนาด
 *     แล้วเก็บแค่ลิงก์รูปไว้ใน AppData ทำให้รูปภาพซิงก์ข้ามอุปกรณ์ได้ด้วยเช่นกัน
 *     ค่าอื่นๆ ที่ยาวเกินขีดจำกัดเซลล์ (เช่น รายการไฟล์แนบเอกสาร) ก็จะถูกย้ายไป
 *     เก็บเป็นไฟล์ข้อความใน Drive (โฟลเดอร์ "MachineManagerBlobs") ให้อัตโนมัติ
 *     เช่นกัน โดยฝั่งแอปไม่ต้องรู้เรื่องนี้เลย เรียก get/set ตามปกติ
 *
 *  3) มิเรอร์ข้อมูลไปยังชีทที่มนุษย์อ่านง่าย (Machines / Cranes / Vehicles /
 *     Tools / Parts / PartsTransactions) ทุกครั้งที่มีการแก้ไข เพื่อให้เปิดดู
 *     ตรงๆ ใน Google Sheet ได้ และให้สคริปต์แจ้งเตือนอีเมลใช้งานต่อได้
 *
 *  4) ส่งอีเมลแจ้งเตือนกำหนดบำรุงรักษาที่ใกล้ถึง (ต้องตั้ง Trigger เอง — ดูข้อ 8)
 *
 * หมายเหตุ: ถ้าคุณมีไฟล์ pm_email_reminder.gs เดิมอยู่แล้วที่ปรับแต่งไว้เฉพาะ
 * ของคุณ (เช่น ข้อความอีเมล, เงื่อนไขวันแจ้งเตือน) ให้เทียบเนื้อหากับไฟล์นี้
 * และย้ายส่วนที่ปรับแต่งไว้มาใส่ในฟังก์ชัน sendPmReminderEmails() ด้านล่าง
 * เพราะไฟล์นี้เป็นการเขียนขึ้นใหม่ทั้งหมดโดยไม่มีต้นฉบับเดิมให้อ้างอิง
 *
 * ===== วิธีติดตั้ง =====
 * 1. เปิด Google Sheet ที่จะใช้เป็นฐานข้อมูล (ไฟล์ใหม่ก็ได้ หรือไฟล์เดิมที่มีอยู่)
 * 2. เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3. ลบโค้ดเดิมในไฟล์ Code.gs ทั้งหมด แล้ววางโค้ดทั้งไฟล์นี้ทับ จากนั้นกดบันทึก
 * 4. ไปที่ Project Settings (รูปเฟืองด้านซ้าย) > Script Properties
 *    กด "Add script property" ใส่:
 *      Property: SYNC_TOKEN
 *      Value:    (ตั้งรหัสลับของคุณเอง ยิ่งยาวและสุ่มยิ่งปลอดภัย เช่น 20+ ตัวอักษร)
 * 5. กด Deploy (มุมขวาบน) > New deployment
 *      - Select type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    กด Deploy แล้วกด Authorize เพื่ออนุญาตสิทธิ์ (ครั้งแรกจะมีคำเตือนความปลอดภัย
 *    ของ Google เพราะสคริปต์ยังไม่ได้ verify — กด Advanced > ไปที่หน้านี้ (ไม่ปลอดภัย)
 *    ได้ตามปกติ เพราะเป็นสคริปต์ของคุณเอง) ต้องกด Allow ให้สิทธิ์เข้าถึง Google Drive
 *    ด้วย (ใช้เก็บไฟล์รูปภาพ) ไม่งั้นฟีเจอร์อัปโหลดรูปจะใช้งานไม่ได้
 *    คัดลอก "Web app URL" ที่ได้ (รูปแบบ https://script.google.com/macros/s/xxxx/exec)
 * 6. เปิดเว็บแอประบบจัดการเครื่องจักร กดปุ่ม 🔄 มุมขวาบน วาง Web App URL และ
 *    SYNC_TOKEN ที่ตั้งไว้ในข้อ 4 ให้ตรงกัน ทำแบบเดียวกันในทุกอุปกรณ์ที่ต้องการ
 *    ให้เห็นข้อมูลชุดเดียวกัน
 * 7. อุปกรณ์ที่มีข้อมูลอยู่แล้ว (เช่น เครื่องที่มี 39 เครื่องจักร) ให้กดปุ่ม
 *    "📤 อัปโหลดข้อมูลเครื่องนี้ขึ้นคลาวด์" หนึ่งครั้งหลังตั้งค่า URL/รหัสลับเสร็จ
 *    จากนั้นอุปกรณ์อื่นๆ กด "⬇ ดึงข้อมูลล่าสุดจากคลาวด์" ก็จะเห็นข้อมูลตรงกันทันที
 * 8. (ไม่บังคับ) เปิดใช้อีเมลแจ้งเตือน: ไปที่ Triggers (รูปนาฬิกาด้านซ้าย) >
 *    Add Trigger > เลือกฟังก์ชัน sendPmReminderEmails > Time-driven >
 *    Day timer > เลือกช่วงเวลาที่ต้องการ (เช่น 07:00-08:00) > Save
 *
 * ถ้าอัปเดตไฟล์นี้ในอนาคต ไม่ต้อง Deploy ใหม่ทั้งหมด — ใช้ Deploy > Manage
 * deployments > กดไอคอนดินสอที่ deployment เดิม > เปลี่ยน Version เป็น
 * "New version" > Deploy ก็พอ (URL เดิมยังใช้ได้ ไม่ต้องไปตั้งค่าใหม่ในแอป)
 * ==========================================================================
 */

const SYNC_TOKEN = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN') || '';
const KV_SHEET_NAME = 'AppData';

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------
function doGet(e){
  try{
    const p = (e && e.parameter) || {};
    if(!checkToken(p.token)) return json({error:'unauthorized'});
    const action = p.action || 'get';
    if(action === 'get')    return json(kvGet(p.key, p.shared === '1'));
    if(action === 'list')   return json(kvList(p.prefix || '', p.shared === '1'));
    if(action === 'export') return json({data: kvExportAll()});
    return json({error:'unknown_action'});
  }catch(err){
    return json({error:String(err)});
  }
}

function doPost(e){
  try{
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if(!checkToken(body.token)) return json({error:'unauthorized'});
    const action = body.action;

    // คำสั่งอ่านข้อมูล — แอปยิงมาทาง POST ทั้งหมด (ไม่ใช้ GET) เพื่อเลี่ยงปัญหา
    // เบราว์เซอร์บล็อกการตอบกลับของ GET ด้วยกลไก CORB
    if(action === 'get')    return json(kvGet(body.key, !!body.shared));
    if(action === 'list')   return json(kvList(body.prefix || '', !!body.shared));
    if(action === 'export') return json({data: kvExportAll()});

    if(action === 'set'){ kvSet(body.key, body.value, !!body.shared); return json({ok:true}); }
    if(action === 'delete'){ kvDelete(body.key, !!body.shared); return json({ok:true}); }
    if(action === 'bulk_set'){
      (body.entries || []).forEach(en => kvSet(en.key, en.value, !!en.shared));
      return json({ok:true, count:(body.entries || []).length});
    }
    if(action === 'set_photo'){
      const url = savePhotoToDrive(body.key, body.value);
      return json({ok:true, url});
    }
    if(action === 'delete_photo'){
      deletePhotoFromDrive(body.key);
      return json({ok:true});
    }
    // ไม่มี action ระบุ แต่มี machines/parts/tools/cranes/vehicles = payload แบบเดิม
    // จากปุ่ม "ซิงก์ชีทที่อ่านง่ายเดี๋ยวนี้" -> เขียนลงชีทที่มนุษย์อ่านได้เท่านั้น
    if(!action && (body.machines || body.parts || body.tools || body.cranes || body.vehicles)){
      publishReadableSheets(body);
      return json({ok:true});
    }
    return json({error:'unknown_action'});
  }catch(err){
    return json({error:String(err)});
  }
}

function checkToken(t){ return !!SYNC_TOKEN && t === SYNC_TOKEN; }
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// KV store — ชีท "AppData" (key, shared, value, updated_at)
// นี่คือฐานข้อมูลจริงที่ทำให้ทุกอุปกรณ์เห็นข้อมูลตรงกัน
// ---------------------------------------------------------------------------
function kvSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(KV_SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(KV_SHEET_NAME);
    sh.appendRow(['key','shared','value','updated_at']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function kvFindRow(sh, key, shared){
  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(data[i][0] === key && !!data[i][1] === !!shared) return i + 1; // 1-indexed row
  }
  return -1;
}
function kvGet(key, shared){
  const sh = kvSheet();
  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(data[i][0] === key && !!data[i][1] === !!shared){
      return {found:true, key, value:resolveBlobValue(data[i][2]), updated_at:data[i][3]};
    }
  }
  return {found:false, key};
}
// เซลล์ของ Google Sheet เก็บตัวอักษรได้ไม่เกิน ~50,000 ตัว ค่าที่ยาวเกินขนาดนี้
// (เช่น รายการไฟล์แนบที่มี base64 ของเอกสารฝังอยู่) จะถูกเก็บเป็นไฟล์ข้อความ
// ใน Google Drive แทนโดยอัตโนมัติ แล้วเก็บแค่ตัวชี้ (blob:<fileId>) ไว้ในเซลล์
// ทำให้ set()/get()/export() ของฝั่งแอปใช้งานได้เหมือนเดิมทุกจุดโดยไม่ต้องรู้เรื่องนี้เลย
const BLOB_MARKER = 'blob:';
const MAX_CELL_CHARS = 35000; // เผื่อระยะห่างจากลิมิตจริงของ Sheets (~50,000)

function kvSet(key, value, shared){
  const sh = kvSheet();
  const idx = kvFindRow(sh, key, shared);
  const now = new Date().toISOString();
  let storeValue = value;
  if(typeof value === 'string' && value.length > MAX_CELL_CHARS){
    deleteValueBlob(key); // ลบไฟล์เก่าคีย์เดียวกันก่อน กันสะสมไฟล์ซ้ำ
    const fileId = saveValueBlob(key, value);
    storeValue = BLOB_MARKER + fileId;
  }
  if(idx === -1) sh.appendRow([key, !!shared, storeValue, now]);
  else sh.getRange(idx, 1, 1, 4).setValues([[key, !!shared, storeValue, now]]);
}
function kvDelete(key, shared){
  const sh = kvSheet();
  const idx = kvFindRow(sh, key, shared);
  if(idx !== -1){
    const raw = sh.getRange(idx, 3).getValue();
    if(typeof raw === 'string' && raw.indexOf(BLOB_MARKER) === 0) deleteValueBlob(key);
    sh.deleteRow(idx);
  }
}
function kvList(prefix, shared){
  const sh = kvSheet();
  const data = sh.getDataRange().getValues();
  const keys = [];
  for(let i = 1; i < data.length; i++){
    if(!!data[i][1] === !!shared && (!prefix || String(data[i][0]).indexOf(prefix) === 0)) keys.push(data[i][0]);
  }
  return {keys, prefix};
}
function kvExportAll(){
  const sh = kvSheet();
  const data = sh.getDataRange().getValues();
  const out = {};
  for(let i = 1; i < data.length; i++){
    out[data[i][0]] = {value:resolveBlobValue(data[i][2]), shared:!!data[i][1], updated_at:data[i][3]};
  }
  return out;
}
function resolveBlobValue(raw){
  if(typeof raw === 'string' && raw.indexOf(BLOB_MARKER) === 0){
    const fileId = raw.substring(BLOB_MARKER.length);
    try{ return DriveApp.getFileById(fileId).getBlob().getDataAsString(); }
    catch(e){ return raw; } // ไฟล์หายหรืออ่านไม่ได้ - คืนค่า marker ดิบไปเผื่อ debug ได้
  }
  return raw;
}
function sanitizeFileBase(key){
  return String(key).replace(/[^a-zA-Z0-9_-]/g, '_');
}
function getBlobFolder(){
  const FOLDER_NAME = 'MachineManagerBlobs';
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(FOLDER_NAME);
}
function saveValueBlob(key, value){
  const base = sanitizeFileBase(key);
  const folder = getBlobFolder();
  const blob = Utilities.newBlob(value, 'text/plain;charset=utf-8', base + '.txt');
  const file = folder.createFile(blob);
  return file.getId();
}
function deleteValueBlob(key){
  const base = sanitizeFileBase(key);
  const folder = getBlobFolder();
  const it = folder.getFiles();
  while(it.hasNext()){
    const f = it.next();
    if(f.getName().indexOf(base + '.') === 0) f.setTrashed(true);
  }
}

// ---------------------------------------------------------------------------
// รูปภาพ — เก็บไฟล์จริงไว้ใน Google Drive แทนการยัด base64 ลงเซลล์ Sheet
// (เซลล์มีขีดจำกัดขนาดตัวอักษร ~50,000 ตัว ซึ่งรูปภาพเกินได้ง่ายมาก)
// ใน AppData จะเก็บแค่ "URL" ของรูปแทน ซึ่งสั้นและโหลดเร็วกว่ามาก
// ---------------------------------------------------------------------------
const PHOTO_FOLDER_NAME = 'MachineManagerPhotos';

function getPhotoFolder(){
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function savePhotoToDrive(key, dataUrl){
  const m = /^data:(.+?);base64,(.+)$/.exec(dataUrl || '');
  if(!m) throw new Error('invalid_data_url');
  const mime = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  const ext = (mime.split('/')[1] || 'jpg').split('+')[0];
  const base = sanitizeFileBase(key);
  const blob = Utilities.newBlob(bytes, mime, base + '.' + ext);

  const folder = getPhotoFolder();
  // ลบไฟล์เก่าคีย์เดียวกันก่อน กันสะสมไฟล์ซ้ำทุกครั้งที่แก้ไขรูป
  const existing = folder.getFiles();
  while(existing.hasNext()){
    const f = existing.next();
    if(f.getName().indexOf(base + '.') === 0) f.setTrashed(true);
  }

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://lh3.googleusercontent.com/d/' + file.getId();
  kvSet(key, url, false); // เก็บ URL ไว้ใน AppData ด้วย เพื่อให้ export/pullAll ของอุปกรณ์อื่นดึงไปได้
  return url;
}

function deletePhotoFromDrive(key){
  const base = sanitizeFileBase(key);
  const folder = getPhotoFolder();
  const it = folder.getFiles();
  while(it.hasNext()){
    const f = it.next();
    if(f.getName().indexOf(base + '.') === 0) f.setTrashed(true);
  }
  kvDelete(key, false);
}

// ---------------------------------------------------------------------------
// ชีทที่มนุษย์อ่านง่าย — มิเรอร์ไว้เผื่อเปิดดูตรงๆ ใน Sheet และให้สคริปต์อีเมลใช้
// (เขียนทับทั้งชีททุกครั้ง เพื่อให้ตรงกับข้อมูลล่าสุดเสมอ)
// ---------------------------------------------------------------------------
function publishReadableSheets(payload){
  writeTable('Machines',
    ['รหัส','ชื่อ','ประเภท','สถานะ','ระดับความสำคัญ','ยี่ห้อ','รุ่น','ปีที่ผลิต','ปีที่ติดตั้ง','ซีเรียล','สถานที่','ผู้จำหน่าย','ผู้รับผิดชอบดูแล','อีเมลผู้รับผิดชอบ','วันที่จัดซื้อ','ราคา','วันหมดประกัน','รอบบำรุงรักษา(วัน)','บำรุงรักษาล่าสุด','กำหนดครั้งถัดไป','หมายเหตุ'],
    (payload.machines || []).map(m => [m.code,m.name,m.type,m.status,m.priority,m.brand,m.model,m.manufactureYear,m.installYear,m.serial,m.location,m.supplier,m.owner,m.email,m.purchaseDate,m.price,m.warrantyDate,m.cycle,m.lastMaint,m.nextMaint,m.notes]));

  writeTable('Cranes',
    ['รหัส','ชื่อ','ประเภท','สถานะ','ระดับความสำคัญ','SWL','หน่วยSWL','เลขที่ปจ.1','เลขที่ปจ.2','ผู้ตรวจสอบ','เลขที่ใบอนุญาตผู้ตรวจสอบ','ยี่ห้อ','รุ่น','ซีเรียล','สถานที่','ผู้จำหน่าย','ผู้รับผิดชอบดูแล','อีเมลผู้รับผิดชอบ','วันที่จัดซื้อ','ราคา','วันหมดประกัน','รอบตรวจสอบ(วัน)','ตรวจสอบล่าสุด','กำหนดครั้งถัดไป','หมายเหตุ'],
    (payload.cranes || []).map(c => [c.code,c.name,c.type,c.status,c.priority,c.swl,c.swlUnit,c.pj1,c.pj2,c.inspector,c.inspectorLicense,c.brand,c.model,c.serial,c.location,c.supplier,c.owner,c.email,c.purchaseDate,c.price,c.warrantyDate,c.cycle,c.lastMaint,c.nextMaint,c.notes]));

  writeTable('Vehicles',
    ['รหัส','ชื่อ','ประเภท','สถานะ','ระดับความสำคัญ','ทะเบียนรถ','ยี่ห้อ','รุ่น','สถานที่','ผู้จำหน่าย','ผู้รับผิดชอบดูแล','อีเมลผู้รับผิดชอบ','วันที่จัดซื้อ','ราคา','วันหมดประกัน','รอบบำรุงรักษา(วัน)','บำรุงรักษาล่าสุด','กำหนดครั้งถัดไป','หมายเหตุ'],
    (payload.vehicles || []).map(v => [v.code,v.name,v.type,v.status,v.priority,v.plate,v.brand,v.model,v.location,v.supplier,v.owner,v.email,v.purchaseDate,v.price,v.warrantyDate,v.cycle,v.lastMaint,v.nextMaint,v.notes]));

  writeTable('Tools',
    ['รหัส','ชื่อ','ประเภท','สถานะ','ระดับความสำคัญ','ยี่ห้อ','รุ่น','ซีเรียล','สถานที่','ผู้จำหน่าย','ผู้รับผิดชอบดูแล','อีเมลผู้รับผิดชอบ','วันที่จัดซื้อ','ราคา','วันหมดประกัน','รอบบำรุงรักษา(วัน)','บำรุงรักษาล่าสุด','กำหนดครั้งถัดไป','หมายเหตุ'],
    (payload.tools || []).map(t => [t.code,t.name,t.type,t.status,t.priority,t.brand,t.model,t.serial,t.location,t.supplier,t.owner,t.email,t.purchaseDate,t.price,t.warrantyDate,t.cycle,t.lastMaint,t.nextMaint,t.notes]));

  writeTable('Parts',
    ['รหัส','ชื่ออะไหล่','รุ่น','หมวดหมู่','คงเหลือ','ขั้นต่ำ','หน่วย','ราคา/หน่วย','ที่จัดเก็บ','เลขที่ล็อต','วันหมดอายุ','ผู้จำหน่าย'],
    (payload.parts || []).map(p => [p.code,p.name,p.model,p.category,p.qty,p.min,p.unit,p.cost,p.location,p.lot,p.expiry,p.supplier]));

  writeTable('PartsTransactions',
    ['วันที่','ประเภท','อะไหล่','จำนวน','เครื่องจักร','เลขที่เอกสาร','ผู้ดำเนินการ','หมายเหตุ'],
    (payload.partsTx || []).map(t => [t.date,t.type,t.partName,t.qty,t.machineCode,t.docRef,t.person,t.note]));
}
function writeTable(sheetName, headers, rows){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if(!sh) sh = ss.insertSheet(sheetName);
  sh.clearContents();
  sh.appendRow(headers);
  if(rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// อีเมลแจ้งเตือนกำหนดบำรุงรักษา — ต้องตั้ง Trigger เอง (ดูขั้นตอนที่ 8 ด้านบน)
// อ่านข้อมูลตรงจากฐานข้อมูล AppData (คีย์ machines_list) จึงเห็นข้อมูลล่าสุดเสมอ
// ---------------------------------------------------------------------------
function sendPmReminderEmails(){
  const data = kvExportAll();
  const raw = data['machines_list'] ? JSON.parse(data['machines_list'].value) : [];
  const today = new Date(); today.setHours(0,0,0,0);
  const SOON_DAYS = 3; // แจ้งเตือนล่วงหน้ากี่วัน ปรับได้ตามต้องการ

  const dueSoon = raw.filter(m => {
    if(!m.nextMaint || m.status === 'retired') return false;
    const d = new Date(m.nextMaint + 'T00:00:00');
    if(isNaN(d.getTime())) return false;
    const diffDays = Math.round((d - today) / 86400000);
    return diffDays >= 0 && diffDays <= SOON_DAYS;
  });
  if(!dueSoon.length) return;

  const byOwner = {};
  dueSoon.forEach(m => {
    const email = m.ownerEmail || m.email;
    if(!email) return;
    (byOwner[email] = byOwner[email] || []).push(m);
  });

  Object.keys(byOwner).forEach(email => {
    const list = byOwner[email];
    const lines = list.map(m => '- ' + m.code + ' ' + m.name + ' (กำหนด ' + m.nextMaint + ')');
    const body = 'แจ้งเตือนกำหนดบำรุงรักษาเครื่องจักรที่ใกล้ถึงกำหนด (ภายใน ' + SOON_DAYS + ' วัน):\n\n' + lines.join('\n') + '\n\nระบบจัดการเครื่องจักร (ส่งอัตโนมัติ)';
    try{ MailApp.sendEmail(email, 'แจ้งเตือนกำหนดบำรุงรักษาเครื่องจักร (' + list.length + ' รายการ)', body); }
    catch(e){ /* ข้ามอีเมลที่ส่งไม่สำเร็จ ไม่ให้กระทบรายการอื่น */ }
  });
}
