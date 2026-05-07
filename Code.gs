/**
 * GOOGLE APPS SCRIPT BACKEND V10.1
 * Fix: missingStaff logic, cross-branch, early-out, office dashboard accuracy
 */

const PASSWORDS = { office: "office123", manager: "manager123" };

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate().setTitle('Attendance System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return res({ success: false, message: 'Invalid JSON' }); }
  ensureSheetsExist();
  try {
    switch(data.action) {
      case 'getInitialData':    return res(getInitialData());
      case 'login':             return res(handleLogin(data.password));
      case 'registerStaff':     return res(registerStaff(data));
      case 'updateStaff':       return res(updateStaff(data));
      case 'getStaffList':      return res(getStaffList(data.role));
      case 'getAdminDashboard': return res(getAdminDashboard());
      case 'logAttendance':     return res(logAttendance(data));
      case 'saveBranch':        return res(saveBranch(data.branch));
      case 'addFaceDescriptors': return res(addFaceDescriptors(data));
      case 'checkTodayStatus':  return res(checkTodayStatus(data.name));
      case 'getAttendanceLogs': return res(getAttendanceLogs(data));
      case 'getAbsentByDate':   return res(getAbsentByDate(data));
      default: return res({ success: false, message: 'Action Not Found' });
    }
  } catch (err) { return res({ success: false, message: err.toString() }); }
}

function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─── Sheet Setup ──────────────────────────────────────────────────────────────
function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = {
    'Staff':      ['Name','Nickname','Descriptors','MainBranchID','Status','CreatedAt'],
    'Branches':   ['ID','MallName','Province','TotalStaff','Lat','Lng','Radius','OpenTime','CloseTime','MinStaff'],
    'Attendance': ['Name','Time','Date','Type','BranchID','HomeBranchID','IsCrossBranch','Lat','Lng','Status','LateMinutes']
  };
  for (const name in schemas) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(schemas[name]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, schemas[name].length)
        .setBackground('#cc0000').setFontColor('#fff').setFontWeight('bold');
    }
  }
}

// ─── Auth ─────────────────────────────────────────────────────────────────────
function handleLogin(pwd) {
  if (pwd === PASSWORDS.office)  return { success: true, role: 'office' };
  if (pwd === PASSWORDS.manager) return { success: true, role: 'manager' };
  return { success: false };
}

// ─── Initial Data ─────────────────────────────────────────────────────────────
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const branches = ss.getSheetByName('Branches').getDataRange().getValues().slice(1)
    .filter(r => r[0]).map(r => ({
      id: String(r[0]), name: r[1], province: r[2],
      totalStaff: Number(r[3]) || 0,
      lat: r[4], lng: r[5], radius: r[6],
      openTime:  fmtTime(r[7]),
      closeTime: fmtTime(r[8]),
      minStaff:  Number(r[9]) || 0
    }));

  const staffRows = ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[4] !== 'Inactive');

  const staffDescriptors = [];
  const staffBranches    = {};
  staffRows.forEach(r => {
    staffBranches[r[0]] = String(r[3]);
    try { JSON.parse(r[2]).forEach(d => staffDescriptors.push({ name: r[0], descriptor: d })); }
    catch(e) {}
  });

  return { success: true, branches, staffDescriptors, staffBranches };
}

// ─── Check Today Status ───────────────────────────────────────────────────────
function checkTodayStatus(name) {
  const today = fmtDate(new Date()); // yyyy-MM-dd GMT+7
  const rows  = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Attendance').getDataRange().getValues().slice(1)
    .filter(r => {
      let d;
      if (r[2] instanceof Date) {
        d = fmtDate(r[2]); // แปลงเป็น GMT+7 เหมือนกัน
      } else {
        // string อาจเป็น "2026-05-07" หรือ "5/7/2026" ขึ้นอยู่กับ Sheets format
        const s = String(r[2]).trim();
        // ลอง parse แล้วแปลงเป็น GMT+7
        const parsed = new Date(s);
        d = isNaN(parsed) ? s : fmtDate(parsed);
      }
      return r[0] === name && d === today;
    });

  const hasIn  = rows.some(r => r[3] === 'IN');
  const hasOut = rows.some(r => r[3] === 'OUT');
  const lastIn  = [...rows].filter(r => r[3] === 'IN').pop();
  const lastOut = [...rows].filter(r => r[3] === 'OUT').pop();

  return {
    success: true, hasIn, hasOut,
    canIn:  !hasIn,
    canOut: hasIn && !hasOut, // เลิกงานได้เฉพาะเมื่อเข้าแล้วและยังไม่ได้เลิก
    lastInTime:   lastIn  ? fmtTime(lastIn[1])  : null,
    lastInBranch: lastIn  ? String(lastIn[4])   : null,
    lastOutTime:  lastOut ? fmtTime(lastOut[1]) : null,
  };
}

// ─── Log Attendance ───────────────────────────────────────────────────────────
function logAttendance(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bInfo = ss.getSheetByName('Branches').getDataRange().getValues().slice(1)
    .find(r => String(r[0]) === String(data.branchId));

  const staffRow = ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .find(r => r[0] === data.name);
  const homeBranchId = staffRow ? String(staffRow[3]) : '';
  const isCross      = homeBranchId && String(data.branchId) !== homeBranchId;

  const now     = new Date();
  const timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');
  const dateStr = fmtDate(now);

  let status = 'ทันเวลา', lateMins = 0;
  let isEarlyOut = false, earlyMins = 0;

  if (data.type === 'IN' && bInfo && bInfo[7]) {
    const lr = calcLate(timeStr, bInfo[7]);
    if (lr.isLate) { status = 'สาย'; lateMins = lr.mins; }
  }

  // Early out detection
  if (data.type === 'OUT' && bInfo && bInfo[8]) {
    const closeTime = fmtTime(bInfo[8]);
    if (closeTime && closeTime.includes(':')) {
      const [cH, cM] = timeStr.split(':').map(Number);
      const [eH, eM] = closeTime.split(':').map(Number);
      const diff = (eH * 60 + eM) - (cH * 60 + cM);
      if (diff > 0) { isEarlyOut = true; earlyMins = diff; status = 'ออกก่อนเวลา'; }
    }
  }

  ss.getSheetByName('Attendance').appendRow([
    data.name, timeStr, dateStr, data.type,
    String(data.branchId), homeBranchId, isCross ? 'ใช่' : 'ไม่',
    data.lat || '', data.lng || '', status, isEarlyOut ? earlyMins : lateMins
  ]);

  return {
    success: true,
    time: timeStr.substring(0,5), status, lateMins,
    isCross, isEarlyOut, earlyMins,
    branchName:  bInfo ? bInfo[1] : String(data.branchId),
    closeTime:   bInfo ? fmtTime(bInfo[8]) : ''
  };
}

// ─── Admin Dashboard (FIXED) ──────────────────────────────────────────────────
function getAdminDashboard() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const branchRows = ss.getSheetByName('Branches').getDataRange().getValues().slice(1).filter(r => r[0]);
  const allStaff   = ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[4] !== 'Inactive');
  const allAtt     = ss.getSheetByName('Attendance').getDataRange().getValues().slice(1);
  const today      = fmtDate(new Date());

  // Build global today's IN/OUT records
  const todayRows = allAtt.filter(r => {
    const d = r[2] instanceof Date ? fmtDate(r[2]) : String(r[2]);
    return d === today;
  });

  // Latest IN per person (across all branches)
  const latestInByName = {};
  todayRows.filter(r => r[3] === 'IN').forEach(r => {
    if (!latestInByName[r[0]] || r[1] > latestInByName[r[0]][1]) latestInByName[r[0]] = r;
  });

  // Latest OUT per person
  const latestOutByName = {};
  todayRows.filter(r => r[3] === 'OUT').forEach(r => {
    if (!latestOutByName[r[0]] || r[1] > latestOutByName[r[0]][1]) latestOutByName[r[0]] = r;
  });

  // Staff who checked in today (anywhere)
  const checkedInAnywhere = new Set(Object.keys(latestInByName));

  // Nickname map
  const nickMap = {};
  allStaff.forEach(s => { nickMap[s[0]] = s[1] || ''; });

  const dashboard = branchRows.map(b => {
    const bId      = String(b[0]);
    const total    = Number(b[3]) || 0;
    const minStaff = Number(b[9]) || total;
    const openTime = fmtTime(b[7]);
    const closeTime= fmtTime(b[8]);

    const homeStaff    = allStaff.filter(s => String(s[3]) === bId).map(s => s[0]);
    const checkedInHere= Object.values(latestInByName).filter(r => String(r[4]) === bId);

    const presentDetails = checkedInHere.map(r => {
      const t  = fmtTime(r[1]);  // ← แก้ Sat D
      const lr = openTime ? calcLate(t, openTime) : { isLate: false, mins: 0 };

      const outRec = latestOutByName[r[0]];
      let isEarlyOut = false, earlyMins = 0;
      if (outRec && closeTime && closeTime.includes(':')) {
        const outT = fmtTime(outRec[1]);  // ← แก้ Sat D
        if (outT && outT.includes(':')) {
          const [oH, oM] = outT.split(':').map(Number);
          const [eH, eM] = closeTime.split(':').map(Number);
          const diff = (eH * 60 + eM) - (oH * 60 + oM);
          if (diff > 0) { isEarlyOut = true; earlyMins = diff; }
        }
      }

      return {
        name:      r[0],
        nickname:  nickMap[r[0]] || '',
        time:      t,
        isLate:    lr.isLate,
        mins:      lr.mins,
        isCross:   String(r[6]) === 'ใช่',
        homeBranch:String(r[5] || ''),
        isEarlyOut, earlyMins
      };
    });

    const presentNames = presentDetails.map(d => d.name);
    const actual       = presentNames.length;

    // ไปช่วยสาขาอื่น → object พร้อม nickname + ชื่อสาขา
    const crossBranchOut = homeStaff
      .filter(n => checkedInAnywhere.has(n) && !presentNames.includes(n))
      .map(n => {
        const rec  = latestInByName[n];
        const dest = rec ? String(rec[4]) : '';
        const dRow = branchRows.find(br => String(br[0]) === dest);
        return { name: n, nickname: nickMap[n] || '', branchId: dest, branchName: dRow ? dRow[1] : dest };
      });

    // ยังไม่มา → object พร้อม nickname
    const missingStaff = homeStaff
      .filter(n => !checkedInAnywhere.has(n))
      .map(n => ({ name: n, nickname: nickMap[n] || '' }));

    // colorStatus: green=ครบทุกคน, yellow=ถึงขั้นต่ำ, red=ต่ำกว่าขั้นต่ำ
    let colorStatus = 'red';
    if      (actual >= total)    colorStatus = 'green';
    else if (actual >= minStaff && minStaff > 0) colorStatus = 'yellow';
    else if (actual >= Math.ceil(minStaff / 2))  colorStatus = 'yellow';

    return {
      id: bId, name: b[1], province: b[2],
      total, minStaff, actual, colorStatus,
      openTime, closeTime,
      presentDetails, missingStaff, crossBranchOut,
      crossBranchersIn: presentDetails.filter(d => d.isCross)
    };
  });

  return { success: true, dashboard };
}

// ─── Staff ────────────────────────────────────────────────────────────────────
function registerStaff(data) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff')
    .appendRow([data.name, data.nickname, JSON.stringify(data.descriptors), data.mainBranchId, 'Active', new Date()]);
  return { success: true };
}

function updateStaff(data) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.name) {
      if (data.newBranchId !== undefined) sheet.getRange(i+1, 4).setValue(data.newBranchId);
      if (data.status      !== undefined) sheet.getRange(i+1, 5).setValue(data.status);
      return { success: true };
    }
  }
  return { success: false };
}

function getStaffList(role) {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff')
    .getDataRange().getValues().slice(1).filter(r => r[0]);
  const mapped = rows.map(r => ({
    name: r[0], nickname: r[1], branchId: String(r[3]),
    status: r[4] || 'Active', createdAt: r[5] ? new Date(r[5]).getTime() : 0
  }));
  // ทุก role เห็นทั้งหมด เรียงล่าสุดบนสุด
  return mapped.sort((a,b) => b.createdAt - a.createdAt);
}

// ─── Add Face Descriptors ─────────────────────────────────────────────────────
function addFaceDescriptors(data) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.name) {
      let existing = [];
      try { existing = JSON.parse(values[i][2]) || []; } catch(e) {}
      const combined = existing.concat(data.descriptors);
      sheet.getRange(i+1, 3).setValue(JSON.stringify(combined));
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบพนักงาน' };
}

// ─── Branch ───────────────────────────────────────────────────────────────────
function saveBranch(b) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Branches');
  const data  = sheet.getDataRange().getValues();
  let idx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(b.id)) { idx = i + 1; break; }
  }
  const row = [b.id, b.name, b.province, b.totalStaff, b.lat, b.lng, b.radius, b.openTime, b.closeTime, b.minStaff];
  if (idx > 0) sheet.getRange(idx, 1, 1, 10).setValues([row]);
  else sheet.appendRow(row);
  return { success: true };
}

// ─── Attendance Logs ──────────────────────────────────────────────────────────
function getAttendanceLogs(data) {
  const today = data.date || fmtDate(new Date());
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const rows  = ss.getSheetByName('Attendance').getDataRange().getValues().slice(1);
  const nickMap = {};
  ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0]) nickMap[r[0]] = r[1] || ''; });

  const logs = rows
    .filter(r => {
      let d;
      if (r[2] instanceof Date) {
        d = fmtDate(r[2]);
      } else {
        const parsed = new Date(String(r[2]).trim());
        d = isNaN(parsed) ? String(r[2]).trim() : fmtDate(parsed);
      }
      return d === today;
    })
    .map(r => ({
      name:      r[0],
      nickname:  nickMap[r[0]] || '',
      time:      fmtTime(r[1]),
      type:      r[3],
      branchId:  String(r[4]),
      homeBranch:String(r[5] || ''),
      isCross:   r[6] === 'ใช่',
      status:    r[9] || 'ทันเวลา',
      lateMins:  Number(r[10]) || 0
    }))
    .reverse();
  return { success: true, logs };
}

// ─── Absent By Date (ย้อนหลัง) ────────────────────────────────────────────────
function getAbsentByDate(data) {
  const date = data.date || fmtDate(new Date());
  const ss   = SpreadsheetApp.getActiveSpreadsheet();

  // พนักงานทุกคนที่ active
  const allStaff = ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[4] !== 'Inactive' && r[4] !== 'Transferred')
    .map(r => ({ name: r[0], nickname: r[1]||'', branchId: String(r[3]) }));

  // attendance วันนั้น
  const attRows = ss.getSheetByName('Attendance').getDataRange().getValues().slice(1)
    .filter(r => {
      let d;
      if (r[2] instanceof Date) { d = fmtDate(r[2]); }
      else { const p = new Date(String(r[2]).trim()); d = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
      return d === date;
    });
  const checkedIn = new Set(attRows.filter(r => r[3] === 'IN').map(r => r[0]));

  // สาขา map
  const branchRows = ss.getSheetByName('Branches').getDataRange().getValues().slice(1).filter(r => r[0]);
  const branchMap  = {};
  branchRows.forEach(b => { branchMap[String(b[0])] = { id:String(b[0]), name:b[1], province:b[2], absent:[] }; });

  // แบ่งตามสาขา
  allStaff.forEach(s => {
    if (!checkedIn.has(s.name)) {
      if (!branchMap[s.branchId]) branchMap[s.branchId] = { id:s.branchId, name:s.branchId, province:'', absent:[] };
      branchMap[s.branchId].absent.push({ name:s.name, nickname:s.nickname });
    }
  });

  const branches    = Object.values(branchMap).filter(b => b.absent.length > 0);
  const totalAbsent = branches.reduce((sum,b) => sum + b.absent.length, 0);
  return { success:true, date, totalAbsent, branches };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function fmtDate(d)   { return Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd'); }
function fmtTime(v) {
  if (!v && v !== 0) return '';

  // Date object → format ด้วย GMT+7 ไม่พึ่ง timezone ของ script
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'GMT+7', 'HH:mm');
  }

  const num = Number(v);

  // Fraction 0-1 → เวลาในวัน (Sheets เก็บ time-only แบบนี้บางครั้ง)
  if (!isNaN(num) && num >= 0 && num < 1) {
    const totalMins = Math.round(num * 24 * 60);
    const h = Math.floor(totalMins / 60);
    const m = totalMins % 60;
    return h.toString().padStart(2, '0') + ':' + m.toString().padStart(2, '0');
  }

  const s = v.toString().trim();
  if (s.includes(':')) return s.substring(0, 5);

  // Thai decimal เช่น "9.30" → "09:30"
  if (!isNaN(s) && s.includes('.')) {
    const parts = s.split('.');
    const h = parts[0].padStart(2, '0');
    const m = (parts[1] || '0').padEnd(2, '0').substring(0, 2);
    return h + ':' + m;
  }

  return s;
}
function calcLate(timeStr, startRaw) {
  const start = fmtTime(startRaw);
  if (!start || !start.includes(':')) return { isLate: false, mins: 0 };
  const [cH, cM] = timeStr.split(':').map(Number);
  const [sH, sM] = start.split(':').map(Number);
  const diff = (cH * 60 + cM) - (sH * 60 + sM);
  return diff > 0 ? { isLate: true, mins: diff } : { isLate: false, mins: 0 };
}
