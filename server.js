const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// Data file paths
const DB_PATH = path.join(__dirname, 'data', 'database.json');
const UPLOADS_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

function readDb() {
  try {
    const raw = fs.readFileSync(DB_PATH, 'utf-8');
    return JSON.parse(raw);
  } catch (err) {
    console.error('Error reading DB:', err);
    return { settings: {}, applications: [], notificationLogs: [] };
  }
}

function writeDb(data) {
  try {
    fs.writeFileSync(DB_PATH, JSON.stringify(data, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error writing DB:', err);
  }
}

// Thai Number Helper
function toThaiNum(num) {
  if (num === null || num === undefined) return '';
  const thaiDigits = ['๐', '๑', '๒', '๓', '๔', '๕', '๖', '๗', '๘', '๙'];
  return String(num).replace(/[0-9]/g, d => thaiDigits[d]);
}

// Thai Months
const THAI_MONTHS = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// Citizen ID Validation
function validateThaiCitizenId(id) {
  if (!id || id.length !== 13 || !/^\d{13}$/.test(id)) return false;
  let sum = 0;
  for (let i = 0; i < 12; i++) {
    sum += parseInt(id.charAt(i), 10) * (13 - i);
  }
  const checkDigit = (11 - (sum % 11)) % 10;
  return checkDigit === parseInt(id.charAt(12), 10);
}

// Multer Config
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, UPLOADS_DIR);
  },
  filename: function (req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, file.fieldname + '-' + uniqueSuffix + ext);
  }
});

const fileFilter = (req, file, cb) => {
  const allowedTypes = ['image/jpeg', 'image/png', 'image/jpg', 'application/pdf'];
  if (allowedTypes.includes(file.mimetype)) {
    cb(null, true);
  } else {
    cb(new Error('รองรับเฉพาะไฟล์ JPG, PNG และ PDF เท่านั้น'), false);
  }
};

const upload = multer({
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
  fileFilter: fileFilter
});

// Notifications
async function sendNotification(message, type = 'general') {
  const db = readDb();
  const token = db.settings.lineNotifyToken;
  let status = 'SENT_LOCAL_LOG';

  if (token && token.trim().length > 0) {
    try {
      await axios.post('https://notify-api.line.me/api/notify', 
        `message=${encodeURIComponent(message)}`, {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Authorization': `Bearer ${token}`
          }
        }
      );
      status = 'SENT_LINE_NOTIFY';
    } catch (err) {
      status = 'LINE_NOTIFY_FAILED: ' + err.message;
    }
  }

  const newLog = {
    id: 'LOG-' + Date.now(),
    timestamp: new Date().toISOString(),
    type,
    message,
    status
  };
  db.notificationLogs.unshift(newLog);
  if (db.notificationLogs.length > 100) {
    db.notificationLogs = db.notificationLogs.slice(0, 100);
  }
  writeDb(db);
  return newLog;
}

// ==========================================
// 1. APPLICATION SUBMISSION & STATUS
// ==========================================

app.post('/api/apply', upload.fields([
  { name: 'photo', maxCount: 1 },
  { name: 'transcript', maxCount: 1 },
  { name: 'idCardFile', maxCount: 1 },
  { name: 'certFile', maxCount: 1 }
]), async (req, res) => {
  try {
    const {
      quotaType, // "ปวส. ( สายตรง )", "ปวส. ( ม.๖/ปวช. ต่างสาขา )", "ปวช."
      fullName,
      idCard,
      birthDay,
      birthMonth,
      birthYear,
      // บิดา
      fatherName,
      fatherAge,
      fatherJob,
      // มารดา
      motherName,
      motherAge,
      motherJob,
      // ที่อยู่ปัจจุบัน
      houseNo,
      villageNo,
      soi,
      road,
      subdistrict,
      district,
      province,
      postcode,
      phone,
      // โรงเรียนเดิม
      currentGrade,
      prevSchool,
      prevSchoolHouseNo,
      prevSchoolSubdistrict,
      prevSchoolDistrict,
      prevSchoolProvince,
      gpax,
      prevMajor,
      // สาขาที่ต้องการสมัคร
      majorPws,
      majorPwc1,
      majorPwc2
    } = req.body;

    if (!quotaType || !fullName || !idCard || !phone) {
      return res.status(400).json({ success: false, message: 'กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วน' });
    }

    const cleanIdCard = idCard.replace(/[-\s]/g, '');
    const cleanPhone = phone.replace(/[-\s]/g, '');

    const db = readDb();
    const academicYear = db.settings.academicYear || '๒๕๖๘';
    const countThisYear = db.applications.length + 1;
    const appId = `QUOTA-${new Date().getFullYear() + 543}-${String(countThisYear).padStart(4, '0')}`;

    // Assigned Department/Major
    let targetMajor = '';
    if (quotaType === 'ปวช.') {
      targetMajor = majorPwc1 || 'ไม่ระบุ';
    } else {
      targetMajor = majorPws || 'ไม่ระบุ';
    }

    const files = req.files || {};
    const filePaths = {
      photo: files['photo'] ? `/uploads/${files['photo'][0].filename}` : '',
      transcript: files['transcript'] ? `/uploads/${files['transcript'][0].filename}` : '',
      idCardFile: files['idCardFile'] ? `/uploads/${files['idCardFile'][0].filename}` : '',
      certFile: files['certFile'] ? `/uploads/${files['certFile'][0].filename}` : ''
    };

    const newApp = {
      id: appId,
      quotaType,
      fullName: fullName.trim(),
      idCard: cleanIdCard,
      birthDay: birthDay || '',
      birthMonth: birthMonth || '',
      birthYear: birthYear || '',
      fatherName: fatherName ? fatherName.trim() : '-',
      fatherAge: fatherAge || '-',
      fatherJob: fatherJob ? fatherJob.trim() : '-',
      motherName: motherName ? motherName.trim() : '-',
      motherAge: motherAge || '-',
      motherJob: motherJob ? motherJob.trim() : '-',
      houseNo: houseNo || '-',
      villageNo: villageNo || '-',
      soi: soi || '-',
      road: road || '-',
      subdistrict: subdistrict || '-',
      district: district || '-',
      province: province || '-',
      postcode: postcode || '-',
      phone: cleanPhone,
      currentGrade: currentGrade || 'ม.3',
      prevSchool: prevSchool ? prevSchool.trim() : '-',
      prevSchoolHouseNo: prevSchoolHouseNo || '-',
      prevSchoolSubdistrict: prevSchoolSubdistrict || '-',
      prevSchoolDistrict: prevSchoolDistrict || '-',
      prevSchoolProvince: prevSchoolProvince || '-',
      gpax: gpax || '0.00',
      prevMajor: prevMajor || '-',
      majorPws: majorPws || '-',
      majorPwc1: majorPwc1 || '-',
      majorPwc2: majorPwc2 || '-',
      targetMajor,
      status: 'pending', // pending, verified, rejected, enrolled
      statusNote: 'ส่งใบสมัครโควตาสำเร็จ รอเจ้าหน้าที่งานรับสมัครตรวจสอบเอกสาร',
      studentId: '',
      enrollDate: '',
      files: filePaths,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    db.applications.push(newApp);
    writeDb(db);

    const notifMsg = `🔔 มีผู้สมัครโควตาใหม่!\n` +
      `วิทยาลัยเทคนิคท่าหลวงซิเมนต์ไทยอนุสรณ์\n` +
      `รหัส: ${appId}\n` +
      `ชื่อ-สกุล: ${fullName}\n` +
      `ระดับ: ${quotaType}\n` +
      `สาขาวิชา: ${targetMajor}\n` +
      `โทร: ${cleanPhone}`;

    sendNotification(notifMsg, 'new_quota_application');

    res.json({
      success: true,
      message: 'บันทึกใบสมัครเข้ารับคัดเลือกเรียบร้อยแล้ว',
      applicationId: appId,
      data: newApp
    });

  } catch (err) {
    console.error('Apply error:', err);
    res.status(500).json({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
});

// Search applicant status
app.get('/api/status/:idCard', (req, res) => {
  const cleanId = req.params.idCard.replace(/[-\s]/g, '');
  const db = readDb();
  const appFound = db.applications.find(a => a.idCard === cleanId || a.id === req.params.idCard);

  if (!appFound) {
    return res.status(404).json({ 
      success: false, 
      message: 'ไม่พบข้อมูลใบสมัครด้วยเลขประจำตัวประชาชนหรือรหัสใบสมัครนี้' 
    });
  }

  res.json({ success: true, data: appFound });
});

// ==========================================
// 2. OFFICIAL PDF REPORT (แบบฟอร์ม 1:1 ตามรูปภาพ)
// ==========================================

app.get('/api/application/:id/pdf', (req, res) => {
  const appId = req.params.id;
  const db = readDb();
  const a = db.applications.find(x => x.id === appId);

  if (!a) {
    return res.status(404).send('ไม่พบข้อมูลใบสมัคร');
  }

  const isPwsDirect = a.quotaType && a.quotaType.includes('สายตรง');
  const isPwsOther = a.quotaType && a.quotaType.includes('ต่างสาขา');
  const isPwc = a.quotaType && (a.quotaType === 'ระดับ ปวช.' || a.quotaType === 'ปวช.');

  const createdDate = new Date(a.createdAt);
  const signDay = createdDate.getDate();
  const signMonth = THAI_MONTHS[createdDate.getMonth()];
  const signYear = createdDate.getFullYear() + 543;

  const html = `
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <title>ใบสมัครเข้ารับคัดเลือกเพื่อศึกษาต่อเป็นกรณีพิเศษ (โควตา) - ${a.id}</title>
  <style>
    @page {
      size: A4 portrait;
      margin: 12mm 15mm 12mm 15mm;
    }
    body {
      font-family: 'Sarabun', 'TH Sarabun New', Tahoma, sans-serif;
      font-size: 13.5pt;
      line-height: 1.5;
      color: #000;
      background: #fff;
      margin: 0;
      padding: 0;
    }
    .print-bar {
      background: #1e293b;
      color: #fff;
      padding: 12px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 100;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .btn-print {
      background: #2563eb;
      color: white;
      border: none;
      padding: 8px 18px;
      font-size: 14px;
      font-weight: bold;
      border-radius: 6px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .btn-print:hover { background: #1d4ed8; }
    .sheet {
      width: 100%;
      max-width: 780px;
      margin: 0 auto;
      padding: 10px 0;
      position: relative;
    }
    .header-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 5px;
    }
    .emblem-cell {
      width: 90px;
      vertical-align: top;
      text-align: left;
    }
    .title-cell {
      vertical-align: top;
      text-align: center;
      padding: 0 5px;
    }
    .photo-cell {
      width: 110px;
      vertical-align: top;
      text-align: right;
    }
    .photo-box {
      width: 105px;
      height: 135px;
      border: 1px solid #000;
      display: inline-block;
      text-align: center;
      vertical-align: middle;
      font-size: 11pt;
      padding-top: 45px;
      box-sizing: border-box;
      line-height: 1.3;
      background: #fafafa;
    }
    .photo-box img {
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
      margin-top: -45px;
    }
    .main-title {
      font-size: 15.5pt;
      font-weight: bold;
      margin-bottom: 3px;
    }
    .college-title {
      font-size: 14.5pt;
      font-weight: bold;
      margin-bottom: 12px;
    }
    .checkbox-group {
      text-align: left;
      margin-left: 200px;
      margin-bottom: 10px;
      font-size: 13pt;
      line-height: 1.7;
    }
    .check-box {
      display: inline-block;
      width: 14px;
      height: 14px;
      border: 1px solid #000;
      margin-right: 8px;
      vertical-align: middle;
      text-align: center;
      line-height: 13px;
      font-weight: bold;
      font-size: 11pt;
    }
    .dotted-line {
      border-bottom: 1px dotted #333;
      display: inline-block;
      padding: 0 4px;
      font-weight: 600;
    }
    .form-row {
      margin-bottom: 4px;
      text-align: justify;
      text-justify: inter-word;
    }
    .num-title {
      font-weight: bold;
      margin-right: 4px;
    }
    .evidence-section {
      margin-top: 15px;
      font-size: 13pt;
    }
    .evidence-item {
      margin-left: 175px;
      margin-bottom: 3px;
    }
    .sign-section {
      margin-top: 25px;
      float: right;
      width: 320px;
      text-align: center;
      font-size: 13pt;
      line-height: 1.6;
    }
    @media print {
      .print-bar { display: none !important; }
      body { margin: 0; background: #fff; }
      .sheet { max-width: 100%; }
    }
  </style>
</head>
<body>

  <div class="print-bar">
    <div style="font-size: 14px;">
      📄 <strong>ใบสมัครรายงานตัวเป็นกรณีพิเศษ (โควตา)</strong> - วิทยาลัยเทคนิคท่าหลวงซิเมนต์ไทยอนุสรณ์ [${a.id}]
    </div>
    <div style="display: flex; gap: 10px;">
      <button class="btn-print" onclick="window.print()">🖨️ สั่งพิมพ์ / บันทึกเป็น PDF</button>
    </div>
  </div>

  <div class="sheet">
    <table class="header-table">
      <tr>
        <td class="emblem-cell">
          <!-- ตราสัญลักษณ์ สอศ. แบบเวกเตอร์ทางการ -->
          <svg width="85" height="85" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="50" cy="50" r="46" stroke="#000" stroke-width="2.5" fill="#f8fafc"/>
            <circle cx="50" cy="50" r="41" stroke="#000" stroke-width="1.2" stroke-dasharray="2 2"/>
            <circle cx="50" cy="50" r="32" stroke="#000" stroke-width="1.5"/>
            <!-- ลายกนกจำลอง / ตราเสมาวิษณุกรรม -->
            <path d="M50 20 L58 35 L50 32 L42 35 Z" fill="#000"/>
            <path d="M35 45 C35 38 65 38 65 45 C65 58 50 62 50 72" stroke="#000" stroke-width="2" fill="none"/>
            <circle cx="50" cy="74" r="3" fill="#000"/>
            <path d="M30 75 Q50 82 70 75" stroke="#000" stroke-width="2" fill="none"/>
            <text x="50" y="93" font-size="6.5" font-weight="bold" text-anchor="middle" fill="#000" font-family="Sarabun">สอศ. อาชีวศึกษา</text>
          </svg>
        </td>
        <td class="title-cell">
          <div class="main-title">ใบสมัครเข้ารับคัดเลือกเพื่อศึกษาต่อเป็นกรณีพิเศษ <u>(โควตา)</u></div>
          <div class="college-title">วิทยาลัยเทคนิคท่าหลวงซิเมนต์ไทยอนุสรณ์ ประจำปีการศึกษา ${toThaiNum(db.settings.academicYear || '๒๕๖๘')}</div>
          
          <div class="checkbox-group">
            <div><span class="check-box">${isPwsDirect ? '✓' : ''}</span> ระดับ ปวส. ( สายตรง )</div>
            <div><span class="check-box">${isPwsOther ? '✓' : ''}</span> ระดับ ปวส. ( ม.๖/ปวช. ต่างสาขา )</div>
            <div><span class="check-box">${isPwc ? '✓' : ''}</span> ระดับ ปวช.</div>
          </div>
        </td>
        <td class="photo-cell">
          <div class="photo-box">
            ${a.files && a.files.photo && (a.files.photo.endsWith('.jpg') || a.files.photo.endsWith('.png'))
              ? `<img src="${a.files.photo}" alt="รูปถ่าย 1 นิ้ว">`
              : `รูปถ่าย<br>ขนาด ๑ นิ้ว`
            }
          </div>
        </td>
      </tr>
    </table>

    <!-- ข้อ ๑ -->
    <div class="form-row" style="margin-top: 10px;">
      <span class="num-title">๑. ชื่อ-นามสกุล</span>
      <span class="dotted-line" style="min-width: 280px;">${a.fullName || '....................................................................'}</span>
      เกิดวันที่ <span class="dotted-line" style="min-width: 45px; text-align: center;">${toThaiNum(a.birthDay) || '..........'}</span>
      เดือน <span class="dotted-line" style="min-width: 95px; text-align: center;">${a.birthMonth || '........................'}</span>
      พ.ศ. <span class="dotted-line" style="min-width: 65px; text-align: center;">${toThaiNum(a.birthYear) || '................'}</span>
    </div>

    <!-- ข้อ ๒ -->
    <div class="form-row">
      <span class="num-title">๒. ชื่อ-สกุลบิดา</span>
      <span class="dotted-line" style="min-width: 320px;">${a.fatherName || '................................................................................'}</span>
      อายุ <span class="dotted-line" style="min-width: 45px; text-align: center;">${toThaiNum(a.fatherAge) || '..........'}</span> ปี
      อาชีพ <span class="dotted-line" style="min-width: 175px;">${a.fatherJob || '................................................'}</span>
    </div>

    <!-- ข้อ ๓ -->
    <div class="form-row">
      <span class="num-title">๓. ชื่อ-สกุลมารดา</span>
      <span class="dotted-line" style="min-width: 305px;">${a.motherName || '................................................................................'}</span>
      อายุ <span class="dotted-line" style="min-width: 45px; text-align: center;">${toThaiNum(a.motherAge) || '..........'}</span> ปี
      อาชีพ <span class="dotted-line" style="min-width: 175px;">${a.motherJob || '................................................'}</span>
    </div>

    <!-- ข้อ ๔ -->
    <div class="form-row">
      <span class="num-title">๔. ที่อยู่ปัจจุบันที่สามารถติดต่อได้</span>
      บ้านเลขที่ <span class="dotted-line" style="min-width: 65px; text-align: center;">${toThaiNum(a.houseNo) || '................'}</span>
      หมู่ <span class="dotted-line" style="min-width: 45px; text-align: center;">${toThaiNum(a.villageNo) || '..........'}</span>
      ซอย <span class="dotted-line" style="min-width: 100px;">${a.soi || '..............................'}</span>
      ถนน <span class="dotted-line" style="min-width: 110px;">${a.road || '..............................'}</span>
    </div>
    <div class="form-row" style="margin-left: 20px;">
      ตำบล <span class="dotted-line" style="min-width: 120px;">${a.subdistrict || '....................................'}</span>
      อำเภอ <span class="dotted-line" style="min-width: 125px;">${a.district || '....................................'}</span>
      จังหวัด <span class="dotted-line" style="min-width: 135px;">${a.province || '........................................'}</span>
    </div>
    <div class="form-row" style="margin-left: 20px;">
      รหัสไปรษณีย์ <span class="dotted-line" style="min-width: 100px; text-align: center;">${toThaiNum(a.postcode) || '............................'}</span>
      หมายเลขโทรศัพท์ <span class="dotted-line" style="min-width: 250px;">${toThaiNum(a.phone) || '................................................................'}</span>
    </div>

    <!-- ข้อ ๕ -->
    <div class="form-row">
      <span class="num-title">๕. กำลังศึกษาอยู่ในระดับชั้น</span>
      <span class="dotted-line" style="min-width: 120px; text-align: center;">${a.currentGrade || '................................'}</span>
      โรงเรียน <span class="dotted-line" style="min-width: 275px;">${a.prevSchool || '....................................................................'}</span>
    </div>
    <div class="form-row" style="margin-left: 20px;">
      ตั้งอยู่เลขที่ <span class="dotted-line" style="min-width: 65px; text-align: center;">${toThaiNum(a.prevSchoolHouseNo) || '................'}</span>
      ตำบล <span class="dotted-line" style="min-width: 110px;">${a.prevSchoolSubdistrict || '................................'}</span>
      อำเภอ <span class="dotted-line" style="min-width: 115px;">${a.prevSchoolDistrict || '................................'}</span>
      จังหวัด <span class="dotted-line" style="min-width: 130px;">${a.prevSchoolProvince || '................................'}</span>
    </div>
    <div class="form-row" style="margin-left: 20px;">
      ผลการเรียนเฉลี่ยสะสม <span class="dotted-line" style="min-width: 110px; text-align: center; font-weight: bold;">${toThaiNum(a.gpax) || '................................'}</span>
      สาขาวิชา <span class="dotted-line" style="min-width: 245px;">${a.prevMajor || '................................................................'}</span>
    </div>

    <!-- ข้อ ๖ -->
    <div class="form-row">
      <span class="num-title">๖. ต้องการสมัครเข้าเรียน</span>
      ระดับ ปวส. สาขาวิชา <span class="dotted-line" style="min-width: 380px; font-weight: bold;">${a.majorPws && a.majorPws !== '-' ? a.majorPws : '................................................................................................'}</span>
    </div>
    <div class="form-row" style="margin-left: 175px;">
      ระดับ ปวช. สาขาวิชาลำดับที่ ๑. <span class="dotted-line" style="min-width: 280px; font-weight: bold;">${a.majorPwc1 && a.majorPwc1 !== '-' ? a.majorPwc1 : '................................................................................'}</span>
    </div>
    <div class="form-row" style="margin-left: 245px;">
      สาขาวิชาลำดับที่ ๒. <span class="dotted-line" style="min-width: 280px;">${a.majorPwc2 && a.majorPwc2 !== '-' ? a.majorPwc2 : '................................................................................'}</span>
    </div>

    <!-- หลักฐานประกอบการสมัคร -->
    <div class="evidence-section">
      <div style="float: left; font-weight: normal;">หลักฐานประกอบการสมัคร</div>
      <div class="evidence-item">๑. เอกสารแสดงผลการเรียนเฉลี่ยสะสม ๕ ภาคเรียน จำนวน ๑ ฉบับ</div>
      <div class="evidence-item">๒. สำเนาบัตรประชาชน / สำเนาทะเบียนบ้าน จำนวน ๑ ฉบับ</div>
      <div class="evidence-item">๓. สำเนาเกียรติบัตร/ใบรับรองความสามารถพิเศษ/ใบประกาศ (ถ้ามี)</div>
    </div>

    <!-- ช่องลงชื่อ -->
    <div class="sign-section">
      <div>ลงชื่อ............................................................ผู้สมัคร</div>
      <div style="margin-top: 3px;">( <span style="display: inline-block; min-width: 210px;">${a.fullName || '............................................................'}</span> )</div>
      <div style="margin-top: 3px;">
        วันที่ <span class="dotted-line" style="min-width: 35px; text-align: center;">${toThaiNum(signDay)}</span>
        เดือน <span class="dotted-line" style="min-width: 80px; text-align: center;">${signMonth}</span>
        พ.ศ. <span class="dotted-line" style="min-width: 55px; text-align: center;">${toThaiNum(signYear)}</span>
      </div>
    </div>

    <div style="clear: both;"></div>

    <!-- ข้อมูลเพิ่มเติมสำหรับเจ้าหน้าที่รับสมัคร (ท้ายกระดาษ) -->
    <div style="margin-top: 30px; border-top: 1px dashed #666; padding-top: 10px; font-size: 11pt; color: #444;">
      <table style="width: 100%;">
        <tr>
          <td><strong>รหัสใบสมัคร:</strong> ${a.id} | <strong>เลขบัตรประชาชน:</strong> ${a.idCard} | <strong>เบอร์โทร:</strong> ${a.phone}</td>
          <td style="text-align: right;"><strong>พิมพ์เมื่อ:</strong> ${new Date().toLocaleDateString('th-TH')}</td>
        </tr>
      </table>
    </div>

  </div>

</body>
</html>
  `;

  res.send(html);
});

// ==========================================
// 3. ADMIN ANALYTICS & STATS (แยกตามแผนกวิชา)
// ==========================================

app.get('/api/admin/stats', (req, res) => {
  const db = readDb();
  const apps = db.applications;

  const total = apps.length;
  const pending = apps.filter(a => a.status === 'pending').length;
  const verified = apps.filter(a => a.status === 'verified').length;
  const rejected = apps.filter(a => a.status === 'rejected').length;
  const enrolled = apps.filter(a => a.status === 'enrolled').length;

  // By Quota Type
  const byQuotaType = {};
  // By Department / Major (แยกตามแผนกวิชา)
  const byDepartment = {
    'แผนกวิชาช่างยนต์': 0,
    'แผนกวิชาช่างกลโรงงาน': 0,
    'แผนกวิชาช่างเชื่อมโลหะ': 0,
    'แผนกวิชาช่างไฟฟ้ากำลัง': 0,
    'แผนกวิชาช่างอิเล็กทรอนิกส์': 0,
    'แผนกวิชาการบัญชี': 0,
    'แผนกวิชาเทคโนโลยีธุรกิจดิจิทัล': 0,
    'แผนกวิชาการตลาด': 0,
    'แผนกวิชาการโรงแรม': 0
  };

  // Raw Major Counts
  const byMajor = {};
  const bySchool = {};

  apps.forEach(a => {
    // Quota Type
    const qType = a.quotaType || 'ไม่ระบุ';
    byQuotaType[qType] = (byQuotaType[qType] || 0) + 1;

    // Target Major
    const m = a.targetMajor || (a.quotaType === 'ปวช.' ? a.majorPwc1 : a.majorPws) || 'ไม่ระบุ';
    byMajor[m] = (byMajor[m] || 0) + 1;

    // Map to Department
    if (m.includes('ช่างยนต์') || m.includes('ยานยนต์')) byDepartment['แผนกวิชาช่างยนต์']++;
    else if (m.includes('กลโรงงาน') || m.includes('การผลิต')) byDepartment['แผนกวิชาช่างกลโรงงาน']++;
    else if (m.includes('เชื่อม')) byDepartment['แผนกวิชาช่างเชื่อมโลหะ']++;
    else if (m.includes('ไฟฟ้า')) byDepartment['แผนกวิชาช่างไฟฟ้ากำลัง']++;
    else if (m.includes('อิเล็กทรอนิกส์')) byDepartment['แผนกวิชาช่างอิเล็กทรอนิกส์']++;
    else if (m.includes('บัญชี')) byDepartment['แผนกวิชาการบัญชี']++;
    else if (m.includes('ดิจิทัล') || m.includes('คอมพิวเตอร์')) byDepartment['แผนกวิชาเทคโนโลยีธุรกิจดิจิทัล']++;
    else if (m.includes('ตลาด')) byDepartment['แผนกวิชาการตลาด']++;
    else if (m.includes('โรงแรม') || m.includes('ท่องเที่ยว')) byDepartment['แผนกวิชาการโรงแรม']++;
    else {
      byDepartment[m] = (byDepartment[m] || 0) + 1;
    }

    if (a.prevSchool && a.prevSchool !== '-') {
      bySchool[a.prevSchool] = (bySchool[a.prevSchool] || 0) + 1;
    }
  });

  res.json({
    success: true,
    data: {
      summary: { total, pending, verified, rejected, enrolled },
      byQuotaType,
      byDepartment,
      byMajor,
      bySchool
    }
  });
});

// Admin Applications List
app.get('/api/admin/applications', (req, res) => {
  const { status, quotaType, department, search } = req.query;
  const db = readDb();
  let list = db.applications;

  if (status && status !== 'all') {
    list = list.filter(a => a.status === status);
  }
  if (quotaType && quotaType !== 'all') {
    list = list.filter(a => a.quotaType === quotaType);
  }
  if (department && department !== 'all') {
    list = list.filter(a => (a.targetMajor || '').includes(department));
  }
  if (search && search.trim() !== '') {
    const q = search.trim().toLowerCase();
    list = list.filter(a =>
      a.id.toLowerCase().includes(q) ||
      (a.fullName && a.fullName.toLowerCase().includes(q)) ||
      (a.idCard && a.idCard.includes(q)) ||
      (a.phone && a.phone.includes(q)) ||
      (a.targetMajor && a.targetMajor.toLowerCase().includes(q)) ||
      (a.prevSchool && a.prevSchool.toLowerCase().includes(q))
    );
  }

  list.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  res.json({ success: true, count: list.length, data: list });
});

// Update Status
app.put('/api/admin/application/:id/status', (req, res) => {
  const appId = req.params.id;
  const { status, statusNote } = req.body;
  const db = readDb();
  const idx = db.applications.findIndex(a => a.id === appId);

  if (idx === -1) return res.status(404).json({ success: false, message: 'ไม่พบใบสมัคร' });

  db.applications[idx].status = status;
  if (statusNote !== undefined) db.applications[idx].statusNote = statusNote;
  db.applications[idx].updatedAt = new Date().toISOString();
  writeDb(db);

  res.json({ success: true, message: 'อัปเดตสถานะสำเร็จ', data: db.applications[idx] });
});

// Admin Login
app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body;
  if (username === 'admin' && password === 'admin1234') {
    res.json({ success: true, token: 'admin-token-' + Date.now() });
  } else {
    res.status(401).json({ success: false, message: 'รหัสผ่านไม่ถูกต้อง (ทดสอบ: admin / admin1234)' });
  }
});

// Export Excel
app.get('/api/admin/export', (req, res) => {
  const db = readDb();
  const exportRows = db.applications.map((a, i) => ({
    'ลำดับ': i + 1,
    'รหัสใบสมัคร': a.id,
    'ประเภทโควตา': a.quotaType,
    'ชื่อ-นามสกุล': a.fullName,
    'เลขบัตรประชาชน': a.idCard,
    'วันเกิด': `${a.birthDay} ${a.birthMonth} ${a.birthYear}`,
    'เบอร์โทรศัพท์': a.phone,
    'โรงเรียนเดิม': a.prevSchool,
    'เกรดเฉลี่ย GPAX 5 เทอม': a.gpax,
    'สาขาวิชาที่สมัคร': a.targetMajor,
    'อันดับ 2 (ปวช.)': a.majorPwc2 || '-',
    'ชื่อบิดา': a.fatherName,
    'ชื่อมารดา': a.motherName,
    'สถานะ': a.status === 'enrolled' ? 'มอบตัวแล้ว' : (a.status === 'verified' ? 'ผ่านคัดเลือก' : 'รอตรวจสอบ'),
    'วันที่ยื่นสมัคร': new Date(a.createdAt).toLocaleString('th-TH')
  }));

  const ws = xlsx.utils.json_to_sheet(exportRows);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'รายชื่อผู้สมัครโควตา');

  const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=quota-report-${Date.now()}.xlsx`);
  res.send(buffer);
});

// Settings
app.get('/api/settings', (req, res) => {
  const db = readDb();
  res.json({ success: true, data: db.settings });
});

app.post('/api/settings', (req, res) => {
  const { lineNotifyToken, academicYear } = req.body;
  const db = readDb();
  if (lineNotifyToken !== undefined) db.settings.lineNotifyToken = lineNotifyToken;
  if (academicYear !== undefined) db.settings.academicYear = academicYear;
  writeDb(db);
  res.json({ success: true, message: 'บันทึกเรียบร้อย' });
});

// Fallback SPA
app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`===================================================`);
  console.log(`🎓 ระบบรับสมัครเรียนโควตาออนไลน์เปิดทำงานแล้ว!`);
  console.log(`🏫 วิทยาลัยเทคนิคท่าหลวงซิเมนต์ไทยอนุสรณ์`);
  console.log(`🔗 เข้าใช้งานได้ที่: http://localhost:${PORT}`);
  console.log(`===================================================`);
});
