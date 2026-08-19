/**************************************************************************************************
 * ระบบส่งผลงาน Best Practices — สพป.อุดรธานี เขต 1
 * นวัตกรรมการจัดการเรียนรู้ภาษาอังกฤษเพื่อการสื่อสาร บูรณาการใช้ปัญญาประดิษฐ์ (ม.3)
 *
 * สคริปต์นี้จะสร้างให้อัตโนมัติ เมื่อกดรัน 1 ครั้ง:
 *   1) Google Form  ให้ครูส่งผลงาน + แนบไฟล์ WORD และ PDF ผ่านระบบได้เลย
 *   2) Google Sheet เป็นฐานข้อมูล มี 4 แผ่นงาน:
 *        - "การส่งผลงาน"  : ข้อมูลที่ครูส่งเข้ามา (เชื่อมกับฟอร์มอัตโนมัติ)
 *        - "ให้คะแนน"     : หน้าให้กรรมการกรอกคะแนน 8 องค์ประกอบ (90) + แผนการสอน (10) รวมอัตโนมัติ 100
 *        - "สรุปผล"       : สรุปจำนวนผลงาน/ระดับคุณภาพ/ผู้ชนะเลิศ อัตโนมัติ
 *        - "เกณฑ์การให้คะแนน" : ตารางเกณฑ์อ้างอิงสำหรับกรรมการ
 *
 * วิธีใช้:  เปิด https://script.google.com  →  โปรเจกต์ใหม่  →  วางโค้ดนี้ทั้งหมด
 *          →  เลือกฟังก์ชัน "สร้างระบบBest"  →  กด Run (▶)  →  อนุญาตสิทธิ์
 *          →  ดูลิงก์ Form และ Sheet ได้ที่หน้าต่าง Execution log
 **************************************************************************************************/

// ====== ตั้งค่าหลัก (แก้ไขได้) ======
var ชื่องาน   = 'ประกวดผลงาน Best Practices ภาษาอังกฤษเพื่อการสื่อสาร บูรณาการ AI (ม.3)';
var หน่วยงาน  = 'สำนักงานเขตพื้นที่การศึกษาประถมศึกษาอุดรธานี เขต 1';
var สังกัดค่าเริ่มต้น = 'สพป.อุดรธานี เขต 1';
var กำหนดส่ง  = 'ภายในวันที่ 27 สิงหาคม 2569';
var เว็บเผยแพร่ = 'https://web.udn1.go.th/';

// น้ำหนักคะแนนของ 11 ตัวชี้วัดย่อย (กรรมการกรอกระดับ 1-3 แล้วคูณน้ำหนัก, รวมสูงสุด 90)
// ลำดับตรงกับหัวตาราง: [1, 2, 3.1, 3.2, 3.3, 4.1, 4.2, 5, 6, 7, 8]
var น้ำหนัก = [3, 1, 2, 4, 3, 4, 4, 3, 2, 2, 2]; // รวม = 30  → ×ระดับสูงสุด 3 = 90

function สร้างระบบBest() {
  // ---------- 1) สร้าง Spreadsheet ฐานข้อมูล ----------
  var ss = SpreadsheetApp.create('ฐานข้อมูลผลงาน Best Practices — ' + สังกัดค่าเริ่มต้น);

  // ---------- 2) สร้าง Google Form ----------
  var form = FormApp.create('แบบส่งผลงาน Best Practices ภาษาอังกฤษ (ม.3) — ' + สังกัดค่าเริ่มต้น);
  form.setDescription(
    'แบบเสนอผลงานรูปแบบ/วิธีปฏิบัติที่เป็นเลิศ (Best Practices)\n' +
    'นวัตกรรมการจัดการเรียนรู้ภาษาอังกฤษเพื่อการสื่อสาร บูรณาการใช้ปัญญาประดิษฐ์\n' +
    'ที่ส่งผลต่อทักษะการสื่อสารภาษาอังกฤษของนักเรียนชั้น ม.3\n' +
    'โรงเรียนในสังกัด ' + หน่วยงาน + '\n\n' +
    '• เขียนรายงานตามแบบฟอร์ม ความยาวไม่เกิน 30 หน้า (ไม่นับปก คำนำ สารบัญ ภาคผนวก)\n' +
    '• แนบไฟล์ทั้ง WORD และ PDF ผ่านแบบฟอร์มนี้\n' +
    '• กำหนดส่ง ' + กำหนดส่ง + '\n' +
    '• เผยแพร่ผลงานผ่าน ' + เว็บเผยแพร่
  );
  form.setCollectEmail(true);          // เก็บอีเมลผู้ส่ง
  form.setAllowResponseEdits(true);    // ให้ครูแก้คำตอบภายหลังได้
  form.setProgressBar(true);

  // --- คำถามข้อมูลผลงาน ---
  form.addTextItem().setTitle('ชื่อผลงาน').setRequired(true);
  form.addTextItem().setTitle('ผู้เสนอผลงาน (ชื่อ-นามสกุล)').setRequired(true);
  form.addTextItem().setTitle('ตำแหน่ง').setRequired(true);
  form.addTextItem().setTitle('โรงเรียน').setRequired(true);
  form.addTextItem().setTitle('สังกัด').setRequired(true);
  form.addTextItem().setTitle('เบอร์โทรศัพท์ที่ติดต่อได้').setRequired(true);

  // --- ช่องแนบไฟล์ (แนบผ่านระบบได้เลย) ---
  var แนบไฟล์สำเร็จ = true;
  try {
    var wordUp = form.addFileUploadItem();
    wordUp.setTitle('แนบไฟล์รายงานฉบับ WORD (.doc / .docx)').setRequired(true);
    wordUp.setHelpText('อัปโหลดไฟล์รายงาน Best Practices ฉบับ Word');
    try { wordUp.setNumberOfFiles(1); } catch (e0) {}

    var pdfUp = form.addFileUploadItem();
    pdfUp.setTitle('แนบไฟล์รายงานฉบับ PDF').setRequired(true);
    pdfUp.setHelpText('อัปโหลดไฟล์รายงาน Best Practices ฉบับ PDF');
    try { pdfUp.setNumberOfFiles(1); } catch (e1) {}
  } catch (eFile) {
    // บางบัญชี (เช่น Gmail ส่วนตัวบางกรณี) อาจแนบไฟล์ในฟอร์มไม่ได้ → ใช้ช่องวางลิงก์แทน
    แนบไฟล์สำเร็จ = false;
    form.addParagraphTextItem()
        .setTitle('ลิงก์ไฟล์รายงาน WORD และ PDF (Google Drive)')
        .setHelpText('อัปโหลดไฟล์ขึ้น Google Drive แล้วตั้งค่าแชร์ "ผู้ที่มีลิงก์" จากนั้นวางลิงก์ทั้ง WORD และ PDF ที่นี่')
        .setRequired(true);
  }

  // ---------- 3) เชื่อมฟอร์มเข้ากับ Spreadsheet ----------
  var ชื่อเดิม = ss.getSheets().map(function (s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();

  // เปลี่ยนชื่อแผ่นคำตอบเป็น "การส่งผลงาน"
  var ssRe = SpreadsheetApp.openById(ss.getId());
  var แผ่นคำตอบ = null;
  ssRe.getSheets().forEach(function (s) {
    if (ชื่อเดิม.indexOf(s.getName()) === -1) แผ่นคำตอบ = s;
  });
  if (แผ่นคำตอบ) แผ่นคำตอบ.setName('การส่งผลงาน');

  // ลบแผ่นเปล่า Sheet1 เดิม (ถ้ามี) ทีหลัง หลังสร้างแผ่นอื่นครบแล้ว
  // ---------- 4) สร้างแผ่น "เกณฑ์การให้คะแนน" ----------
  สร้างแผ่นเกณฑ์(ssRe);

  // ---------- 5) สร้างแผ่น "ให้คะแนน" ----------
  สร้างแผ่นให้คะแนน(ssRe);

  // ---------- 6) สร้างแผ่น "สรุปผล" ----------
  สร้างแผ่นสรุปผล(ssRe);

  // ลบแผ่นค่าเริ่มต้นที่ว่าง (Sheet1 / แผ่น1)
  var ทุกแผ่น = ssRe.getSheets();
  if (ทุกแผ่น.length > 1) {
    ชื่อเดิม.forEach(function (nm) {
      var sh = ssRe.getSheetByName(nm);
      if (sh && sh.getLastRow() === 0 && ssRe.getSheets().length > 1) {
        try { ssRe.deleteSheet(sh); } catch (e) {}
      }
    });
  }

  // จัดลำดับแผ่นงาน
  จัดลำดับแผ่น(ssRe, ['การส่งผลงาน', 'ให้คะแนน', 'สรุปผล', 'เกณฑ์การให้คะแนน']);

  // ---------- 7) แสดงผลลัพธ์ ----------
  var สรุป =
    '✅ สร้างระบบเสร็จแล้ว!\n\n' +
    '📝 ลิงก์ฟอร์มให้ครูส่งผลงาน (ส่งลิงก์นี้ให้ครู):\n' + form.getPublishedUrl() + '\n\n' +
    '🛠️ ลิงก์แก้ไขฟอร์ม:\n' + form.getEditUrl() + '\n\n' +
    '📊 ลิงก์ Google Sheet ฐานข้อมูล + ให้คะแนน:\n' + ssRe.getUrl() + '\n\n' +
    (แนบไฟล์สำเร็จ
      ? '📎 ครูสามารถแนบไฟล์ WORD และ PDF ผ่านฟอร์มได้โดยตรง'
      : '📎 หมายเหตุ: บัญชีนี้แนบไฟล์ในฟอร์มไม่ได้ ระบบจึงใช้ช่อง "วางลิงก์ไฟล์" แทน');
  Logger.log(สรุป);
  try { SpreadsheetApp.getActive(); } catch (e) {}
  return สรุป;
}

/* ---------------------------------- แผ่น "ให้คะแนน" ---------------------------------- */
function สร้างแผ่นให้คะแนน(ss) {
  var sh = ss.getSheetByName('ให้คะแนน') || ss.insertSheet('ให้คะแนน');
  sh.clear();

  var หัว = [
    'ลำดับ', 'ชื่อผลงาน', 'ผู้เสนอผลงาน', 'โรงเรียน', 'กรรมการผู้ประเมิน',
    '1.ความเป็นมาฯ (ระดับ1-3)',
    '2.วัตถุประสงค์ (ระดับ1-3)',
    '3.1 ออกแบบแผน (ระดับ1-3)',
    '3.2 กระบวนการเรียนรู้ (ระดับ1-3)',
    '3.3 สื่อ/AI สอดคล้อง (ระดับ1-3)',
    '4.1 ผลบรรลุกิจกรรม (ระดับ1-3)',
    '4.2 ประโยชน์ที่ได้รับ (ระดับ1-3)',
    '5.บทเรียนที่ได้รับ (ระดับ1-3)',
    '6.ปัจจัยความสำเร็จ (ระดับ1-3)',
    '7.การเผยแพร่ (ระดับ1-3)',
    '8.ภาคผนวก (ระดับ1-3)',
    'คะแนนองค์ประกอบ (90)',
    'คะแนนแผนการสอน (0-10)',
    'คะแนนรวม (100)',
    'ระดับคุณภาพ'
  ];
  sh.getRange(1, 1, 1, หัว.length).setValues([หัว]);

  var แถวสูงสุด = 200; // เตรียมสูตรไว้ 200 แถว
  // คอลัมน์: A ลำดับ | B ชื่อผลงาน | C ผู้เสนอ | D โรงเรียน | E กรรมการ
  //          F..P = 11 ตัวชี้วัด (ระดับ 1-3) | Q คะแนน90 | R แผน10 | S รวม100 | T ระดับ
  for (var r = 2; r <= แถวสูงสุด + 1; r++) {
    var i = r - 1;
    // ดึงข้อมูลผลงานจากแผ่น "การส่งผลงาน" อัตโนมัติ (คอลัมน์ C=ชื่อผลงาน, D=ผู้เสนอ, F=โรงเรียน)
    sh.getRange(r, 1).setFormula('=IF(\'การส่งผลงาน\'!C' + r + '="","",' + i + ')');
    sh.getRange(r, 2).setFormula('=IF(\'การส่งผลงาน\'!C' + r + '="","",\'การส่งผลงาน\'!C' + r + ')');
    sh.getRange(r, 3).setFormula('=IF(\'การส่งผลงาน\'!D' + r + '="","",\'การส่งผลงาน\'!D' + r + ')');
    sh.getRange(r, 4).setFormula('=IF(\'การส่งผลงาน\'!F' + r + '="","",\'การส่งผลงาน\'!F' + r + ')');
    // Q = คะแนนองค์ประกอบ (90) = SUMPRODUCT(ระดับ F:P, น้ำหนัก)
    sh.getRange(r, 17).setFormula(
      '=IF(COUNT(F' + r + ':P' + r + ')=0,"",' +
      'F' + r + '*3+G' + r + '*1+H' + r + '*2+I' + r + '*4+J' + r + '*3+' +
      'K' + r + '*4+L' + r + '*4+M' + r + '*3+N' + r + '*2+O' + r + '*2+P' + r + '*2)'
    );
    // S = รวม (100) = Q + R
    sh.getRange(r, 19).setFormula(
      '=IF(AND(Q' + r + '="",R' + r + '=""),"",N(Q' + r + ')+N(R' + r + '))'
    );
    // T = ระดับคุณภาพ
    sh.getRange(r, 20).setFormula(
      '=IF(S' + r + '="","",' +
      'IF(S' + r + '>=90,"ดีเยี่ยม",' +
      'IF(S' + r + '>=80,"ดีมาก",' +
      'IF(S' + r + '>=70,"ดี",' +
      'IF(S' + r + '>=50,"เข้าร่วม","ต่ำกว่าเกณฑ์")))))'
    );
  }

  // Data validation: ระดับ 1-3 (F:P), แผน 0-10 (R)
  var ruleระดับ = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 3).setAllowInvalid(false)
    .setHelpText('กรอกระดับ 1, 2 หรือ 3 เท่านั้น (3=มีครบ 3 รายการ, 2=มี 2 รายการ, 1=มี 1 รายการ)').build();
  sh.getRange(2, 6, แถวสูงสุด, 11).setDataValidation(ruleระดับ);
  var ruleแผน = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 10).setAllowInvalid(false)
    .setHelpText('คะแนนแผนการจัดการเรียนรู้ 0-10 (ดูรายการ 10 ข้อในแผ่น "เกณฑ์การให้คะแนน")').build();
  sh.getRange(2, 18, แถวสูงสุด, 1).setDataValidation(ruleแผน);

  // จัดรูปแบบหัวตาราง
  var head = sh.getRange(1, 1, 1, หัว.length);
  head.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff')
      .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.getRange(1, 17, แถวสูงสุด + 1, 4).setBackground('#e8f0fe');
  sh.getRange(2, 19, แถวสูงสุด, 1).setFontWeight('bold');
  sh.setColumnWidths(6, 11, 90);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(20, 110);

  // สีตามระดับคุณภาพ
  var rangeT = sh.getRange(2, 20, แถวสูงสุด, 1);
  var rules = [];
  var mk = function (txt, bg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(txt).setBackground(bg).setRanges([rangeT]).build();
  };
  rules.push(mk('ดีเยี่ยม', '#b7e1cd'));
  rules.push(mk('ดีมาก', '#d9ead3'));
  rules.push(mk('ดี', '#fff2cc'));
  rules.push(mk('เข้าร่วม', '#fce5cd'));
  rules.push(mk('ต่ำกว่าเกณฑ์', '#f4cccc'));
  sh.setConditionalFormatRules(rules);
}

/* ---------------------------------- แผ่น "สรุปผล" ---------------------------------- */
function สร้างแผ่นสรุปผล(ss) {
  var sh = ss.getSheetByName('สรุปผล') || ss.insertSheet('สรุปผล');
  sh.clear();
  sh.getRange('A1').setValue('สรุปผลการประกวด Best Practices').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue(หน่วยงาน).setFontColor('#555555');

  var ข้อมูล = [
    ['รายการ', 'จำนวน / ค่า'],
    ['จำนวนผลงานที่ส่งเข้ามา', '=COUNTIF(\'การส่งผลงาน\'!C2:C,"<>")'],
    ['จำนวนผลงานที่ตรวจแล้ว', '=COUNT(\'ให้คะแนน\'!S2:S)'],
    ['ระดับ ดีเยี่ยม (90 ขึ้นไป)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ดีเยี่ยม")'],
    ['ระดับ ดีมาก (80-89.99)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ดีมาก")'],
    ['ระดับ ดี (70-79.99)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ดี")'],
    ['ระดับ เข้าร่วม (50-69.99)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"เข้าร่วม")'],
    ['ต่ำกว่าเกณฑ์ (ต่ำกว่า 50)', '=COUNTIF(\'ให้คะแนน\'!T2:T,"ต่ำกว่าเกณฑ์")'],
    ['คะแนนสูงสุด', '=IFERROR(MAX(\'ให้คะแนน\'!S2:S),"-")'],
    ['คะแนนเฉลี่ย', '=IFERROR(ROUND(AVERAGE(\'ให้คะแนน\'!S2:S),2),"-")'],
    ['ผลงานชนะเลิศ (คะแนนสูงสุด)',
      '=IFERROR(INDEX(\'ให้คะแนน\'!B2:B,MATCH(MAX(\'ให้คะแนน\'!S2:S),\'ให้คะแนน\'!S2:S,0)),"-")'],
    ['ผู้เสนอผลงานชนะเลิศ',
      '=IFERROR(INDEX(\'ให้คะแนน\'!C2:C,MATCH(MAX(\'ให้คะแนน\'!S2:S),\'ให้คะแนน\'!S2:S,0)),"-")'],
    ['โรงเรียนชนะเลิศ',
      '=IFERROR(INDEX(\'ให้คะแนน\'!D2:D,MATCH(MAX(\'ให้คะแนน\'!S2:S),\'ให้คะแนน\'!S2:S,0)),"-")']
  ];
  sh.getRange(4, 1, ข้อมูล.length, 2).setValues(ข้อมูล);
  sh.getRange(4, 1, 1, 2).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange(4, 1, ข้อมูล.length, 2).setBorder(true, true, true, true, true, true);
  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 260);
  sh.getRange(5, 1, ข้อมูล.length - 1, 1).setFontWeight('normal');
}

/* ---------------------------------- แผ่น "เกณฑ์การให้คะแนน" ---------------------------------- */
function สร้างแผ่นเกณฑ์(ss) {
  var sh = ss.getSheetByName('เกณฑ์การให้คะแนน') || ss.insertSheet('เกณฑ์การให้คะแนน');
  sh.clear();
  sh.getRange('A1').setValue('เกณฑ์การคัดเลือกผลงาน Best Practices (รวม 100 คะแนน)')
    .setFontSize(14).setFontWeight('bold');

  var ตาราง = [
    ['องค์ประกอบ', 'น้ำหนัก(×ระดับ1-3)', 'คะแนนเต็ม', 'กรอบการพิจารณาโดยย่อ'],
    ['1. ความเป็นมาและความสำคัญ', 3, 9, 'ที่มา/ความสำคัญของปัญหา หลักการ แนวคิด ทฤษฎี CLT เพื่อพัฒนาการสื่อสารตามกรอบ CEFR'],
    ['2. วัตถุประสงค์และเป้าหมาย', 1, 3, 'วัตถุประสงค์สอดคล้องปัญหา + เป้าหมายเชิงปริมาณและคุณภาพ'],
    ['3.1 การออกแบบแผนการเรียนรู้ (CLT+AI)', 2, 6, 'จุดประสงค์/ตัวชี้วัด สะท้อนความรู้ ทักษะ สมรรถนะผู้เรียน เป็นขั้นตอน'],
    ['3.2 กระบวนการจัดการเรียนรู้ (CLT+AI)', 4, 12, 'กิจกรรมสอดคล้องจุดประสงค์ + เทคนิค CLT บูรณาการ AI + วัดผลเชิงประจักษ์หลากหลาย'],
    ['3.3 ความสอดคล้องของสื่อ AI กับแผน', 3, 9, 'สื่อ/นวัตกรรม AI สอดคล้องมาตรฐาน ตัวชี้วัด และกระบวนการ + ประเมินความพึงพอใจ'],
    ['4.1 ผลที่เกิดขึ้นบรรลุตามกิจกรรม', 4, 12, 'ผลสอดคล้องวัตถุประสงค์ มีหลักฐาน อ้างผลสัมฤทธิ์ O-NET ปี 2567-2568'],
    ['4.2 ประโยชน์ที่ได้รับ', 4, 12, 'สื่อ/นวัตกรรมก่อประสบการณ์เรียนรู้ทั้งในและนอกห้องเรียน/รายบุคคล'],
    ['5. บทเรียนที่ได้รับ (Lesson Learned)', 3, 9, 'หลักการที่ได้เรียนรู้ ข้อสรุป ข้อเสนอแนะการนำไปประยุกต์ใช้'],
    ['6. ปัจจัยความสำเร็จ', 2, 6, 'ปัจจัย/บุคคล/หน่วยงาน + วิธีการ + ร่องรอยหลักฐานที่ช่วยให้สำเร็จ'],
    ['7. การเผยแพร่/การยอมรับ/รางวัล', 2, 6, 'ร่องรอยการเผยแพร่ การได้รับการยอมรับ และรางวัลระดับต่างๆ'],
    ['8. ภาคผนวก', 2, 6, 'แผนการสอนอย่างน้อย 1 ชม. + ร่องรอยหลักฐาน หนังสือราชการ เกียรติบัตร ฯลฯ'],
    ['รวมองค์ประกอบ', '', 90, ''],
    ['คะแนนแผนการจัดการเรียนรู้', '', 10, 'ดูรายการ 10 ข้อด้านล่าง (ข้อละ 1 คะแนน)'],
    ['รวมทั้งสิ้น', '', 100, '']
  ];
  sh.getRange(3, 1, ตาราง.length, 4).setValues(ตาราง);
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange(3, 1, ตาราง.length, 4).setBorder(true, true, true, true, true, true).setWrap(true);
  sh.getRange(3 + ตาราง.length - 3, 1, 3, 4).setFontWeight('bold').setBackground('#e8f0fe');

  var เริ่มแผน = 3 + ตาราง.length + 2;
  sh.getRange(เริ่มแผน, 1).setValue('เกณฑ์การให้คะแนนแผนการจัดการเรียนรู้ (10 คะแนน — มี=1, ไม่มี=0)')
    .setFontWeight('bold').setFontSize(12);
  var แผน = [
    ['ข้อ', 'รายการ'],
    ['1', 'ระบุมาตรฐาน/ตัวชี้วัด'],
    ['2', 'ระบุสาระสำคัญ/ความคิดรวบยอด'],
    ['3', 'ระบุจุดประสงค์การเรียนรู้/ผลการเรียนรู้'],
    ['4', 'ระบุสาระการเรียนรู้'],
    ['5', 'ระบุกิจกรรมการเรียนรู้ CLT บูรณาการ AI ที่ส่งผลต่อทักษะการสื่อสาร'],
    ['6', 'ระบุสื่อและแหล่งการเรียนรู้'],
    ['7', 'ระบุวิธีการวัดและประเมินผลการเรียนรู้'],
    ['8', 'มีบันทึกหลังสอน/ปัญหา/ข้อเสนอแนะ + ผู้บริหารรับรองแผน'],
    ['9', 'ภาคผนวก: ร่องรอยหลักฐาน เช่น ภาพกิจกรรม ใบงาน/แบบประเมิน เกณฑ์การประเมิน'],
    ['10', 'ภาคผนวก: คลิปวิดีโอการสอน และ infographic แผนการจัดการเรียนรู้']
  ];
  sh.getRange(เริ่มแผน + 1, 1, แผน.length, 2).setValues(แผน);
  sh.getRange(เริ่มแผน + 1, 1, 1, 2).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  sh.getRange(เริ่มแผน + 1, 1, แผน.length, 2).setBorder(true, true, true, true, true, true).setWrap(true);

  var เริ่มระดับ = เริ่มแผน + แผน.length + 3;
  sh.getRange(เริ่มระดับ, 1).setValue('เกณฑ์ระดับคุณภาพ (จากคะแนนรวม 100)').setFontWeight('bold').setFontSize(12);
  var ระดับ = [
    ['ระดับคุณภาพ', 'ช่วงคะแนน'],
    ['ดีเยี่ยม', '90.00 ขึ้นไป'],
    ['ดีมาก', '80.00 - 89.99'],
    ['ดี', '70.00 - 79.99'],
    ['เข้าร่วม', '50.00 - 69.99']
  ];
  sh.getRange(เริ่มระดับ + 1, 1, ระดับ.length, 2).setValues(ระดับ);
  sh.getRange(เริ่มระดับ + 1, 1, 1, 2).setFontWeight('bold').setBackground('#fbbc04');
  sh.getRange(เริ่มระดับ + 1, 1, ระดับ.length, 2).setBorder(true, true, true, true, true, true);

  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 460);
}

/* ---------------------------------- ช่วยจัดลำดับแผ่น ---------------------------------- */
function จัดลำดับแผ่น(ss, ลำดับ) {
  for (var i = 0; i < ลำดับ.length; i++) {
    var sh = ss.getSheetByName(ลำดับ[i]);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  }
  var first = ss.getSheetByName('การส่งผลงาน');
  if (first) ss.setActiveSheet(first);
}
