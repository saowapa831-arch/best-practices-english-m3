/**************************************************************************************************
 * ระบบส่งผลงาน Best Practices — เวอร์ชัน "หน้าเว็บของเราเอง" (ไม่ใช้ Google Form)
 * สพป.อุดรธานี เขต 1
 *
 * ส่วนนี้คือ "หลังบ้าน" (Backend) — Google Apps Script Web App
 * รับข้อมูลจากหน้าเว็บ index.html → เก็บไฟล์ WORD/PDF ลง Google Drive → เขียนแถวลง Google Sheet
 *
 * ★ ขั้นตอนใช้งาน (ทำครั้งเดียว):
 *   1) เปิด https://script.google.com → โครงการใหม่ → วางโค้ดนี้ทั้งหมด → บันทึก
 *   2) เลือกฟังก์ชัน "ติดตั้งระบบ" → กด Run (▶) → อนุญาตสิทธิ์
 *      → ดูลิงก์ Google Sheet และโฟลเดอร์ไฟล์ได้ที่ Execution log
 *   3) เมนู Deploy → New deployment → เลือกชนิด "Web app"
 *        - Execute as: Me (ตัวคุณครูเอง)
 *        - Who has access: Anyone  (ทุกคน)
 *      → Deploy → คัดลอก "Web app URL"
 *   4) นำ Web app URL ไปวางในไฟล์ index.html ตรงบรรทัด WEBAPP_URL
 **************************************************************************************************/

var ชื่อไฟล์ชีต   = 'ฐานข้อมูลผลงาน Best Practices — สพป.อุดรธานี เขต 1';
var ชื่อโฟลเดอร์   = 'ไฟล์ผลงาน Best Practices (ม.3)';
var สังกัดค่าเริ่มต้น = 'สพป.อุดรธานี เขต 1';

// น้ำหนักคะแนน 11 ตัวชี้วัดย่อย: [1, 2, 3.1, 3.2, 3.3, 4.1, 4.2, 5, 6, 7, 8] รวม×3 = 90
var น้ำหนัก = [3, 1, 2, 4, 3, 4, 4, 3, 2, 2, 2];

/* ============================ 1) ติดตั้งระบบ (รันครั้งเดียว) ============================ */
function ติดตั้งระบบ() {
  var props = PropertiesService.getScriptProperties();

  // สร้าง (หรือใช้ซ้ำ) Spreadsheet
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(ชื่อไฟล์ชีต);
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  // สร้าง (หรือใช้ซ้ำ) โฟลเดอร์เก็บไฟล์ใน Drive
  var folderId = props.getProperty('FOLDER_ID');
  var folder = null;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder(ชื่อโฟลเดอร์);
    props.setProperty('FOLDER_ID', folder.getId());
  }

  สร้างแผ่นการส่งผลงาน(ss);
  สร้างแผ่นเกณฑ์(ss);
  สร้างแผ่นให้คะแนน(ss);
  สร้างแผ่นสรุปผล(ss);

  // ลบแผ่นเริ่มต้นที่ว่าง
  var d = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1');
  if (d && ss.getSheets().length > 1) { try { ss.deleteSheet(d); } catch (e) {} }

  จัดลำดับแผ่น(ss, ['การส่งผลงาน', 'ให้คะแนน', 'สรุปผล', 'เกณฑ์การให้คะแนน']);

  var msg =
    '✅ ติดตั้งฐานข้อมูลเรียบร้อย!\n\n' +
    '📊 Google Sheet ฐานข้อมูล:\n' + ss.getUrl() + '\n\n' +
    '📁 โฟลเดอร์เก็บไฟล์ผลงาน:\n' + folder.getUrl() + '\n\n' +
    'ขั้นต่อไป: เมนู Deploy → New deployment → Web app (Execute as: Me, Access: Anyone)\n' +
    'แล้วนำ Web app URL ไปวางใน index.html';
  Logger.log(msg);
  return msg;
}

/* ============================ 2) รับข้อมูลจากหน้าเว็บ ============================ */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var props = PropertiesService.getScriptProperties();
    var ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
    var folder = DriveApp.getFolderById(props.getProperty('FOLDER_ID'));
    var sheet = ss.getSheetByName('การส่งผลงาน');

    var ชื่อผลงาน = (data.chuePhonngan || '').toString().trim();
    var ผู้เสนอ   = (data.phuSanoe || '').toString().trim();
    if (!ชื่อผลงาน || !ผู้เสนอ) {
      return jsonOut({ ok: false, error: 'กรุณากรอกชื่อผลงานและชื่อผู้เสนอผลงาน' });
    }

    // ตั้งชื่อไฟล์ให้สื่อความ: ชื่อผลงาน_ผู้เสนอ
    var ฐานชื่อ = (ชื่อผลงาน + '_' + ผู้เสนอ).replace(/[\\\/:*?"<>|]/g, ' ').substring(0, 80);

    var wordUrl = บันทึกไฟล์(folder, data.wordName, data.wordMime, data.wordData, ฐานชื่อ + '_WORD');
    var pdfUrl  = บันทึกไฟล์(folder, data.pdfName, data.pdfMime, data.pdfData, ฐานชื่อ + '_PDF');

    sheet.appendRow([
      new Date(),
      ชื่อผลงาน,
      ผู้เสนอ,
      (data.tamnaeng || '').toString().trim(),
      (data.rongrian || '').toString().trim(),
      (data.sangkat || สังกัดค่าเริ่มต้น).toString().trim(),
      (data.tel || '').toString().trim(),
      (data.email || '').toString().trim(),
      wordUrl || '(ไม่มีไฟล์)',
      pdfUrl || '(ไม่มีไฟล์)',
      (data.videoLink || '').toString().trim(),
      (data.extraLink || '').toString().trim(),
      data.confirm ? 'ยืนยันแล้ว' : ''
    ]);

    return jsonOut({ ok: true, message: 'บันทึกผลงานเรียบร้อยแล้ว' });
  } catch (err) {
    return jsonOut({ ok: false, error: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;padding:40px;text-align:center">' +
    '<h2>✅ ระบบรับผลงาน Best Practices พร้อมทำงาน</h2>' +
    '<p>นี่คือหลังบ้าน (Web App) — ให้ครูส่งผลงานผ่านหน้าเว็บฟอร์ม (index.html)</p>' +
    '</body></html>');
}

function บันทึกไฟล์(folder, name, mime, base64, ตั้งชื่อ) {
  if (!base64) return '';
  var นามสกุล = '';
  if (name && name.indexOf('.') > -1) นามสกุล = name.substring(name.lastIndexOf('.'));
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), mime || 'application/octet-stream', ตั้งชื่อ + นามสกุล);
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file.getUrl();
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ แผ่น "การส่งผลงาน" (ฐานข้อมูล) ============================ */
function สร้างแผ่นการส่งผลงาน(ss) {
  var sh = ss.getSheetByName('การส่งผลงาน');
  if (!sh) sh = ss.insertSheet('การส่งผลงาน', 0);
  if (sh.getLastRow() === 0) {
    var หัว = ['เวลาบันทึก', 'ชื่อผลงาน', 'ผู้เสนอผลงาน', 'ตำแหน่ง', 'โรงเรียน', 'สังกัด',
              'เบอร์โทร', 'อีเมล', 'ไฟล์ WORD', 'ไฟล์ PDF',
              'ลิงก์คลิปวิดีโอการสอน', 'ลิงก์สื่อ/ผลงานเพิ่มเติม', 'ยืนยันเงื่อนไข'];
    sh.getRange(1, 1, 1, หัว.length).setValues([หัว]);
    sh.getRange(1, 1, 1, หัว.length).setFontWeight('bold')
      .setBackground('#1a73e8').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(2, 240);
    sh.setColumnWidth(3, 160);
    sh.setColumnWidth(9, 220);
    sh.setColumnWidth(10, 220);
  }
}

/* ============================ แผ่น "ให้คะแนน" ============================ */
function สร้างแผ่นให้คะแนน(ss) {
  if (ss.getSheetByName('ให้คะแนน')) return; // มีอยู่แล้ว ไม่สร้างซ้ำ
  var sh = ss.insertSheet('ให้คะแนน');

  var หัว = [
    'ลำดับ', 'ชื่อผลงาน', 'ผู้เสนอผลงาน', 'โรงเรียน', 'กรรมการผู้ประเมิน',
    '1.ความเป็นมาฯ (1-3)', '2.วัตถุประสงค์ (1-3)', '3.1 ออกแบบแผน (1-3)',
    '3.2 กระบวนการเรียนรู้ (1-3)', '3.3 สื่อ/AI สอดคล้อง (1-3)', '4.1 ผลบรรลุกิจกรรม (1-3)',
    '4.2 ประโยชน์ที่ได้รับ (1-3)', '5.บทเรียนที่ได้รับ (1-3)', '6.ปัจจัยความสำเร็จ (1-3)',
    '7.การเผยแพร่ (1-3)', '8.ภาคผนวก (1-3)',
    'คะแนนองค์ประกอบ (90)', 'คะแนนแผนการสอน (0-10)', 'คะแนนรวม (100)', 'ระดับคุณภาพ'
  ];
  sh.getRange(1, 1, 1, หัว.length).setValues([หัว]);

  var แถวสูงสุด = 200;
  for (var r = 2; r <= แถวสูงสุด + 1; r++) {
    var i = r - 1;
    // ดึงข้อมูลจากแผ่น "การส่งผลงาน": B=ชื่อผลงาน, C=ผู้เสนอ, E=โรงเรียน
    sh.getRange(r, 1).setFormula('=IF(\'การส่งผลงาน\'!B' + r + '="","",' + i + ')');
    sh.getRange(r, 2).setFormula('=IF(\'การส่งผลงาน\'!B' + r + '="","",\'การส่งผลงาน\'!B' + r + ')');
    sh.getRange(r, 3).setFormula('=IF(\'การส่งผลงาน\'!C' + r + '="","",\'การส่งผลงาน\'!C' + r + ')');
    sh.getRange(r, 4).setFormula('=IF(\'การส่งผลงาน\'!E' + r + '="","",\'การส่งผลงาน\'!E' + r + ')');
    sh.getRange(r, 17).setFormula(
      '=IF(COUNT(F' + r + ':P' + r + ')=0,"",' +
      'F' + r + '*3+G' + r + '*1+H' + r + '*2+I' + r + '*4+J' + r + '*3+' +
      'K' + r + '*4+L' + r + '*4+M' + r + '*3+N' + r + '*2+O' + r + '*2+P' + r + '*2)'
    );
    sh.getRange(r, 19).setFormula('=IF(AND(Q' + r + '="",R' + r + '=""),"",N(Q' + r + ')+N(R' + r + '))');
    sh.getRange(r, 20).setFormula(
      '=IF(S' + r + '="","",IF(S' + r + '>=90,"ดีเยี่ยม",IF(S' + r + '>=80,"ดีมาก",' +
      'IF(S' + r + '>=70,"ดี",IF(S' + r + '>=50,"เข้าร่วม","ต่ำกว่าเกณฑ์")))))'
    );
  }

  var ruleระดับ = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 3)
    .setAllowInvalid(false).setHelpText('กรอกระดับ 1, 2 หรือ 3 (3=ครบ 3 รายการ, 2=มี 2, 1=มี 1)').build();
  sh.getRange(2, 6, แถวสูงสุด, 11).setDataValidation(ruleระดับ);
  var ruleแผน = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 10)
    .setAllowInvalid(false).setHelpText('คะแนนแผนการสอน 0-10 (ดูรายการ 10 ข้อในแผ่นเกณฑ์)').build();
  sh.getRange(2, 18, แถวสูงสุด, 1).setDataValidation(ruleแผน);

  sh.getRange(1, 1, 1, หัว.length).setFontWeight('bold').setBackground('#1a73e8')
    .setFontColor('#ffffff').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.getRange(1, 17, แถวสูงสุด + 1, 4).setBackground('#e8f0fe');
  sh.getRange(2, 19, แถวสูงสุด, 1).setFontWeight('bold');
  sh.setColumnWidths(6, 11, 90);
  sh.setColumnWidth(2, 220);

  var rangeT = sh.getRange(2, 20, แถวสูงสุด, 1);
  var mk = function (txt, bg) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt)
      .setBackground(bg).setRanges([rangeT]).build();
  };
  sh.setConditionalFormatRules([
    mk('ดีเยี่ยม', '#b7e1cd'), mk('ดีมาก', '#d9ead3'), mk('ดี', '#fff2cc'),
    mk('เข้าร่วม', '#fce5cd'), mk('ต่ำกว่าเกณฑ์', '#f4cccc')
  ]);
}

/* ============================ แผ่น "สรุปผล" ============================ */
function สร้างแผ่นสรุปผล(ss) {
  if (ss.getSheetByName('สรุปผล')) return;
  var sh = ss.insertSheet('สรุปผล');
  sh.getRange('A1').setValue('สรุปผลการประกวด Best Practices').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('สพป.อุดรธานี เขต 1').setFontColor('#555555');

  var ข้อมูล = [
    ['รายการ', 'จำนวน / ค่า'],
    ['จำนวนผลงานที่ส่งเข้ามา', '=COUNTIF(\'การส่งผลงาน\'!B2:B,"<>")'],
    ['จำนวนผลงานที่ตรวจแล้ว', '=COUNT(\'ให้คะแนน\'!S2:S)'],
    ['ระดับ ดีเยี่ยม (90 ขึ้นไป)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ดีเยี่ยม")'],
    ['ระดับ ดีมาก (80-89.99)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ดีมาก")'],
    ['ระดับ ดี (70-79.99)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ดี")'],
    ['ระดับ เข้าร่วม (50-69.99)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"เข้าร่วม")'],
    ['คะแนนสูงสุด', '=IFERROR(MAX(\'ให้คะแนน\'!S2:S),"-")'],
    ['คะแนนเฉลี่ย', '=IFERROR(ROUND(AVERAGE(\'ให้คะแนน\'!S2:S),2),"-")'],
    ['ผลงานชนะเลิศ', '=IFERROR(INDEX(\'ให้คะแนน\'!B2:B,MATCH(MAX(\'ให้คะแนน\'!S2:S),\'ให้คะแนน\'!S2:S,0)),"-")'],
    ['ผู้เสนอผลงานชนะเลิศ', '=IFERROR(INDEX(\'ให้คะแนน\'!C2:C,MATCH(MAX(\'ให้คะแนน\'!S2:S),\'ให้คะแนน\'!S2:S,0)),"-")'],
    ['โรงเรียนชนะเลิศ', '=IFERROR(INDEX(\'ให้คะแนน\'!D2:D,MATCH(MAX(\'ให้คะแนน\'!S2:S),\'ให้คะแนน\'!S2:S,0)),"-")']
  ];
  sh.getRange(4, 1, ข้อมูล.length, 2).setValues(ข้อมูล);
  sh.getRange(4, 1, 1, 2).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange(4, 1, ข้อมูล.length, 2).setBorder(true, true, true, true, true, true);
  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 300);
}

/* ============================ แผ่น "เกณฑ์การให้คะแนน" ============================ */
function สร้างแผ่นเกณฑ์(ss) {
  if (ss.getSheetByName('เกณฑ์การให้คะแนน')) return;
  var sh = ss.insertSheet('เกณฑ์การให้คะแนน');
  sh.getRange('A1').setValue('เกณฑ์การคัดเลือกผลงาน Best Practices (รวม 100 คะแนน)').setFontSize(14).setFontWeight('bold');

  var ตาราง = [
    ['องค์ประกอบ', 'น้ำหนัก(×1-3)', 'เต็ม', 'กรอบการพิจารณาโดยย่อ'],
    ['1. ความเป็นมาและความสำคัญ', 3, 9, 'ที่มา/ความสำคัญของปัญหา หลักการ แนวคิด ทฤษฎี CLT พัฒนาการสื่อสารตามกรอบ CEFR'],
    ['2. วัตถุประสงค์และเป้าหมาย', 1, 3, 'วัตถุประสงค์สอดคล้องปัญหา + เป้าหมายเชิงปริมาณและคุณภาพ'],
    ['3.1 การออกแบบแผนการเรียนรู้ (CLT+AI)', 2, 6, 'จุดประสงค์/ตัวชี้วัด สะท้อนความรู้ ทักษะ สมรรถนะผู้เรียน เป็นขั้นตอน'],
    ['3.2 กระบวนการจัดการเรียนรู้ (CLT+AI)', 4, 12, 'กิจกรรมสอดคล้องจุดประสงค์ + เทคนิค CLT บูรณาการ AI + วัดผลเชิงประจักษ์หลากหลาย'],
    ['3.3 ความสอดคล้องของสื่อ AI กับแผน', 3, 9, 'สื่อ/นวัตกรรม AI สอดคล้องมาตรฐาน ตัวชี้วัด กระบวนการ + ประเมินความพึงพอใจ'],
    ['4.1 ผลที่เกิดขึ้นบรรลุตามกิจกรรม', 4, 12, 'ผลสอดคล้องวัตถุประสงค์ มีหลักฐาน อ้างผลสัมฤทธิ์ O-NET ปี 2567-2568'],
    ['4.2 ประโยชน์ที่ได้รับ', 4, 12, 'สื่อ/นวัตกรรมก่อประสบการณ์เรียนรู้ทั้งในและนอกห้องเรียน/รายบุคคล'],
    ['5. บทเรียนที่ได้รับ (Lesson Learned)', 3, 9, 'หลักการที่ได้เรียนรู้ ข้อสรุป ข้อเสนอแนะการนำไปประยุกต์ใช้'],
    ['6. ปัจจัยความสำเร็จ', 2, 6, 'ปัจจัย/บุคคล/หน่วยงาน + วิธีการ + ร่องรอยหลักฐานที่ช่วยให้สำเร็จ'],
    ['7. การเผยแพร่/การยอมรับ/รางวัล', 2, 6, 'ร่องรอยการเผยแพร่ การได้รับการยอมรับ และรางวัลระดับต่างๆ'],
    ['8. ภาคผนวก', 2, 6, 'แผนการสอนอย่างน้อย 1 ชม. + ร่องรอยหลักฐาน หนังสือราชการ เกียรติบัตร ฯลฯ'],
    ['รวมองค์ประกอบ', '', 90, ''],
    ['คะแนนแผนการจัดการเรียนรู้', '', 10, '10 ข้อ ข้อละ 1 คะแนน'],
    ['รวมทั้งสิ้น', '', 100, '']
  ];
  sh.getRange(3, 1, ตาราง.length, 4).setValues(ตาราง);
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange(3, 1, ตาราง.length, 4).setBorder(true, true, true, true, true, true).setWrap(true);

  var b = 3 + ตาราง.length + 1;
  sh.getRange(b, 1).setValue('ระดับคุณภาพ: ดีเยี่ยม 90↑ · ดีมาก 80-89.99 · ดี 70-79.99 · เข้าร่วม 50-69.99')
    .setFontWeight('bold');

  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 470);
}

function จัดลำดับแผ่น(ss, ลำดับ) {
  for (var i = 0; i < ลำดับ.length; i++) {
    var sh = ss.getSheetByName(ลำดับ[i]);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  }
}
