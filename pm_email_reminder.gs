/**
 * ==========================================================
 * ระบบแจ้งเตือนบำรุงรักษาเครื่องจักร (PM) ทางอีเมล - อัตโนมัติ
 * ==========================================================
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *
 * 1. เปิด Google Sheets สร้างชีตใหม่ ตั้งชื่อชีตแรก (แท็บด้านล่าง) ว่า "Machines"
 *
 * 2. ในชีตนี้ ไปที่เมนู ส่วนขยาย (Extensions) > Apps Script
 *    ลบโค้ดเดิมทั้งหมดในไฟล์ที่เปิดขึ้นมา แล้ววางโค้ดทั้งไฟล์นี้แทน
 *
 * 3. แก้ไขค่าในส่วน CONFIG ด้านล่างให้ตรงกับของคุณ (อย่างน้อยต้องเปลี่ยน
 *    NOTIFY_EMAIL เป็นอีเมลของคุณเอง)
 *
 * 4. ที่แถบด้านบนของ Apps Script เลือกฟังก์ชัน "setupTrigger" จาก dropdown
 *    แล้วกดปุ่ม Run (รูปสามเหลี่ยม) ครั้งแรกระบบจะขอสิทธิ์เข้าถึง Gmail/ชีต
 *    ให้กด "อนุญาต" (Allow) ได้เลย เป็นสิทธิ์ในบัญชี Google ของคุณเอง
 *    ไม่มีข้อมูลใดถูกส่งไปที่อื่นนอกจากอีเมลที่คุณกำหนด
 *
 * 5. เสร็จแล้ว! ระบบจะรันอัตโนมัติทุกวันตามเวลาที่ตั้งไว้ (ค่าเริ่มต้น 07:00 น.)
 *    และส่งอีเมลสรุปเฉพาะวันที่มีเครื่องจักรเลยกำหนด/ใกล้ถึงกำหนดบำรุงรักษาเท่านั้น
 *
 * ==========================================================
 * เปิดใช้การซิงก์อัตโนมัติจากแอประบบจัดการเครื่องจักร (แนะนำ - ไม่ต้อง export/import CSV เอง)
 * ==========================================================
 *
 * 6. ในหน้า Apps Script (มุมขวาบน) กดปุ่ม "Deploy" > "New deployment"
 * 7. ที่ช่อง "Select type" กดไอคอนรูปเฟือง เลือก "Web app"
 * 8. ตั้งค่า:
 *      - Execute as: Me (บัญชีของคุณ)
 *      - Who has access: Anyone
 *    (ไม่ต้องกังวลเรื่องความปลอดภัย เพราะสคริปต์นี้แค่รับข้อมูลมาบันทึกลงชีตของคุณเอง
 *     ไม่ได้เปิดให้ใครอ่านข้อมูลออกไปได้)
 * 9. กด Deploy แล้วกด "Authorize access" อนุญาตสิทธิ์ตามที่ระบบขอ
 * 10. คัดลอก "Web app URL" ที่ได้ (จะขึ้นต้นด้วย https://script.google.com/macros/s/.../exec)
 * 11. เปิดแอประบบจัดการเครื่องจักร กดปุ่ม "🔄 ซิงก์ Google Sheet" แล้ววาง URL นี้ลงไป
 *     ติ๊ก "ซิงก์อัตโนมัติทุกครั้งที่มีการเพิ่ม/แก้ไข/ลบเครื่องจักร" แล้วกดบันทึกการตั้งค่า
 *
 * จากนี้ไป ทุกครั้งที่มีการเพิ่ม/แก้ไข/ลบเครื่องจักร หรือบันทึกงานบำรุงรักษาในแอป
 * ข้อมูลจะถูกส่งไปอัปเดตในชีต "Machines" นี้โดยอัตโนมัติ ไม่ต้อง export/import CSV เองอีกต่อไป
 *
 * หมายเหตุ: ถ้าคุณแก้ไขโค้ดสคริปต์นี้ใหม่ในภายหลัง ต้องกด Deploy > Manage deployments
 * > แก้ไข (ไอคอนดินสอ) > เลือกเวอร์ชันใหม่ > Deploy อีกครั้ง เพื่อให้ลิงก์เดิมใช้โค้ดล่าสุด
 *
 * ทดสอบด้วยตัวเอง:
 *    เลือกฟังก์ชัน "checkAndNotify" แล้วกด Run เพื่อสั่งให้ตรวจสอบและส่งอีเมลทันที
 *    (จะส่งก็ต่อเมื่อมีรายการเลยกำหนด/ใกล้ถึงกำหนดจริงเท่านั้น)
 * ==========================================================
 */

const CONFIG = {
  SHEET_NAME: 'Machines',      // ชื่อชีตที่เก็บข้อมูลเครื่องจักร
  NOTIFY_EMAIL: 'you@example.com', // อีเมลหลักที่จะรับสรุปการแจ้งเตือนทุกวัน (แก้เป็นของคุณ)
  DAYS_AHEAD: 7,                // แจ้งเตือนล่วงหน้ากี่วันก่อนถึงกำหนด
  RUN_HOUR: 7                   // ชั่วโมงที่จะให้รันทุกวัน (0-23)
};

/**
 * อ่านข้อมูลเครื่องจักรจากชีต โดยอ้างอิงหัวคอลัมน์ภาษาไทย
 * (ไม่ยึดตำแหน่งคอลัมน์ตายตัว รองรับแม้สลับลำดับคอลัมน์)
 */
function getMachines_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "' + CONFIG.SHEET_NAME + '" กรุณาตรวจสอบชื่อชีตให้ตรงกับ CONFIG.SHEET_NAME');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const idx = (name) => headers.indexOf(name);
  const col = {
    code: idx('รหัส'),
    name: idx('ชื่อ'),
    status: idx('สถานะ'),
    location: idx('สถานที่'),
    owner: idx('ผู้รับผิดชอบ'),
    next: idx('กำหนดครั้งถัดไป'),
    email: idx('อีเมลผู้รับผิดชอบ') // คอลัมน์เสริม ไม่บังคับต้องมี
  };

  if (col.code === -1 || col.next === -1) {
    throw new Error('ไม่พบคอลัมน์ "รหัส" หรือ "กำหนดครั้งถัดไป" ในชีต กรุณาตรวจสอบหัวตาราง');
  }

  return data.slice(1)
    .filter(r => r[col.code])
    .map(r => ({
      code: r[col.code],
      name: col.name >= 0 ? r[col.name] : '',
      status: col.status >= 0 ? r[col.status] : '',
      location: col.location >= 0 ? r[col.location] : '',
      owner: col.owner >= 0 ? r[col.owner] : '',
      next: r[col.next],
      email: col.email >= 0 ? r[col.email] : ''
    }));
}

function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function daysUntil_(date) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const d = new Date(date); d.setHours(0, 0, 0, 0);
  return Math.round((d - today) / 86400000);
}

/**
 * ฟังก์ชันหลัก: ตรวจสอบกำหนดบำรุงรักษาทั้งหมด และส่งอีเมลสรุป
 * จะรันทุกวันอัตโนมัติผ่าน trigger ที่ตั้งไว้ใน setupTrigger()
 */
function checkAndNotify() {
  const machines = getMachines_();
  const overdue = [], soon = [];

  machines.forEach(m => {
    if (m.status === 'ปลดระวาง') return; // ข้ามเครื่องที่ปลดระวางแล้ว
    const d = parseDate_(m.next);
    if (!d) return;
    const days = daysUntil_(d);
    if (days < 0) overdue.push(Object.assign({}, m, { days: days }));
    else if (days <= CONFIG.DAYS_AHEAD) soon.push(Object.assign({}, m, { days: days }));
  });

  if (overdue.length === 0 && soon.length === 0) {
    Logger.log('ไม่มีรายการที่ต้องแจ้งเตือนวันนี้');
    return;
  }

  let html = '<div style="font-family:sans-serif;">';
  html += '<h2>สรุปการแจ้งเตือนบำรุงรักษาเครื่องจักร (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy') + ')</h2>';

  if (overdue.length) {
    html += '<h3 style="color:#c0392b;">🔴 เลยกำหนดบำรุงรักษา (' + overdue.length + ' รายการ)</h3><ul>';
    overdue.forEach(m => {
      html += '<li><b>' + m.code + '</b> ' + m.name + ' — เลยกำหนดมาแล้ว ' + Math.abs(m.days) + ' วัน '
        + '(สถานที่: ' + (m.location || '-') + ', ผู้รับผิดชอบ: ' + (m.owner || '-') + ')</li>';
    });
    html += '</ul>';
  }

  if (soon.length) {
    html += '<h3 style="color:#b8860b;">🟡 ใกล้ถึงกำหนดบำรุงรักษา (' + soon.length + ' รายการ)</h3><ul>';
    soon.forEach(m => {
      html += '<li><b>' + m.code + '</b> ' + m.name + ' — อีก ' + m.days + ' วัน '
        + '(สถานที่: ' + (m.location || '-') + ', ผู้รับผิดชอบ: ' + (m.owner || '-') + ')</li>';
    });
    html += '</ul>';
  }
  html += '</div>';

  // ส่งอีเมลสรุปรวมถึงอีเมลหลัก
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    subject: 'แจ้งเตือนบำรุงรักษาเครื่องจักร: เลยกำหนด ' + overdue.length + ' รายการ / ใกล้ถึงกำหนด ' + soon.length + ' รายการ',
    htmlBody: html
  });
  Logger.log('ส่งอีเมลสรุปไปที่ ' + CONFIG.NOTIFY_EMAIL);

  // ถ้ามีคอลัมน์ "อีเมลผู้รับผิดชอบ" ให้ส่งแจ้งเตือนแยกถึงแต่ละคนด้วย
  overdue.concat(soon).forEach(m => {
    if (m.email) {
      const statusText = m.days < 0
        ? 'เลยกำหนดมาแล้ว ' + Math.abs(m.days) + ' วัน'
        : 'อีก ' + m.days + ' วันถึงกำหนดบำรุงรักษา';
      MailApp.sendEmail({
        to: m.email,
        subject: 'แจ้งเตือนบำรุงรักษา: ' + m.code + ' ' + m.name,
        body: m.name + ' (' + m.code + ')\n' + statusText + '\nสถานที่: ' + (m.location || '-')
      });
    }
  });
}

/**
 * เรียกใช้ครั้งเดียวตอนติดตั้ง เพื่อสร้างตัวตั้งเวลาให้รันทุกวันอัตโนมัติ
 * ถ้ารันซ้ำจะลบตัวตั้งเวลาเก่าก่อนแล้วสร้างใหม่ (ไม่ทำให้ซ้ำซ้อน)
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAndNotify') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkAndNotify')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.RUN_HOUR)
    .create();
  Logger.log('ตั้งเวลาทำงานอัตโนมัติทุกวัน เวลา ' + CONFIG.RUN_HOUR + ':00 น. เรียบร้อย');
  checkAndNotify(); // ทดสอบรันทันทีหนึ่งครั้ง
}

/**
 * รับข้อมูลเครื่องจักรที่ส่งมาจากแอประบบจัดการเครื่องจักร (ผ่านปุ่ม "ซิงก์ Google Sheet")
 * แล้วเขียนทับข้อมูลทั้งหมดในชีต "Machines" ให้เป็นข้อมูลล่าสุดโดยอัตโนมัติ
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const rows = payload.machines || [];

    const headers = ['รหัส','ชื่อ','ประเภท','สถานะ','ยี่ห้อ','รุ่น','ซีเรียล','สถานที่','ผู้จำหน่าย',
      'ผู้รับผิดชอบ','อีเมลผู้รับผิดชอบ','วันที่จัดซื้อ','ราคา','วันหมดประกัน','รอบบำรุงรักษา(วัน)',
      'บำรุงรักษาล่าสุด','กำหนดครั้งถัดไป','หมายเหตุ'];

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.SHEET_NAME);
    sheet.clearContents();

    const data = [headers].concat(rows.map(m => [
      m.code, m.name, m.type, m.status, m.brand, m.model, m.serial, m.location, m.supplier,
      m.owner, m.email || '', m.purchaseDate, m.price, m.warrantyDate, m.cycle, m.lastMaint, m.nextMaint, m.notes
    ]));
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);

    return ContentService.createTextOutput(JSON.stringify({ ok: true, count: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ใช้สำหรับทดสอบว่า Web App deploy สำเร็จหรือไม่
 * เปิดลิงก์ Web App URL ในเบราว์เซอร์โดยตรง ถ้าเห็นข้อความนี้แปลว่าใช้งานได้
 */
function doGet(e) {
  return ContentService.createTextOutput('PM Email Reminder Web App is running. ส่งข้อมูลด้วย POST เท่านั้น');
}
