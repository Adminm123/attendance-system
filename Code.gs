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
      case 'getLateByDate':       return res(getLateByDate(data));
      case 'getSuspiciousStaff':  return res(getSuspiciousStaff());
      case 'getTodayAttendance':  return res(getTodayAttendance());
      case 'getEmployeeHistory':  return res(getEmployeeHistory(data));
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
    'Attendance': ['Name','Time','Date','Type','BranchID','HomeBranchID','IsCrossBranch','Lat','Lng','Status','LateMinutes','Accuracy']
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
  const today     = fmtDate(new Date());
  const attSheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
  const lastRow   = attSheet.getLastRow();
  if (lastRow <= 1) return { success:true, hasIn:false, hasOut:false, canIn:true, canOut:false, lastInTime:null, lastInBranch:null, lastOutTime:null };

  const dataRange = attSheet.getDataRange();
  const values    = dataRange.getValues().slice(1);
  const displays  = dataRange.getDisplayValues().slice(1);

  const todayRows = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    let d;
    if (r[2] instanceof Date) { d = fmtDate(r[2]); }
    else { const p = new Date(String(r[2]).trim()); d = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
    if (r[0] === name && d === today) todayRows.push({ v: r, d: displays[i] });
  }

  const hasIn  = todayRows.some(r => r.v[3] === 'IN');
  const hasOut = todayRows.some(r => r.v[3] === 'OUT');
  const lastIn  = [...todayRows].filter(r => r.v[3] === 'IN').pop();
  const lastOut = [...todayRows].filter(r => r.v[3] === 'OUT').pop();

  return {
    success: true, hasIn, hasOut,
    canIn:  !hasIn,
    canOut: hasIn && !hasOut,
    lastInTime:   lastIn  ? normalizeTimeDisplay(lastIn.d[1],  lastIn.v[1])  : null,
    lastInBranch: lastIn  ? String(lastIn.v[4]) : null,
    lastOutTime:  lastOut ? normalizeTimeDisplay(lastOut.d[1], lastOut.v[1]) : null,
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

  // ── ป้องกัน OUT โดยไม่มี IN วันนี้ ──────────────────────────────
  if (data.type === 'OUT') {
    const todayRows = ss.getSheetByName('Attendance').getDataRange().getValues().slice(1)
      .filter(r => {
        let d;
        if (r[2] instanceof Date) { d = fmtDate(r[2]); }
        else { const p = new Date(String(r[2]).trim()); d = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
        return r[0] === data.name && d === dateStr;
      });
    const hasIn  = todayRows.some(r => r[3] === 'IN');
    const hasOut = todayRows.some(r => r[3] === 'OUT');
    if (!hasIn)  return { success: false, message: 'ยังไม่ได้เช็คอินวันนี้' };
    if (hasOut)  return { success: false, message: 'เลิกงานไปแล้ววันนี้' };
  }

  ss.getSheetByName('Attendance').appendRow([
    data.name, timeStr, dateStr, data.type,
    String(data.branchId), homeBranchId, isCross ? 'ใช่' : 'ไม่',
    data.lat || '', data.lng || '', status, isEarlyOut ? earlyMins : lateMins,
    data.accuracy !== undefined ? Math.round(data.accuracy) : -1  // Accuracy (m)
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
  const attRange   = ss.getSheetByName('Attendance').getDataRange();
  const allAtt     = attRange.getValues().slice(1);
  const allAttDisp = attRange.getDisplayValues().slice(1);
  allAtt.forEach((r, i) => { r._dt = allAttDisp[i][1]; });
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
      const t  = normalizeTimeDisplay(r._dt, r[1]);
      const lr = openTime ? calcLate(t, openTime) : { isLate: false, mins: 0 };

      const outRec = latestOutByName[r[0]];
      let isEarlyOut = false, earlyMins = 0;
      if (outRec && closeTime && closeTime.includes(':')) {
        const outT = normalizeTimeDisplay(outRec._dt, outRec[1]);
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
  const nickMap = {};
  ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0]) nickMap[r[0]] = r[1] || ''; });

  const attRange   = ss.getSheetByName('Attendance').getDataRange();
  const rows       = attRange.getValues().slice(1);
  const rowsDisp   = attRange.getDisplayValues().slice(1);
  rows.forEach((r, i) => { r._dt = rowsDisp[i][1]; });

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
      time:      normalizeTimeDisplay(r._dt, r[1]),
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

// ─── Late By Date (มาสายย้อนหลัง) ────────────────────────────────────────────
function getLateByDate(data) {
  const date = data.date || fmtDate(new Date());
  const ss   = SpreadsheetApp.getActiveSpreadsheet();

  // nickname map
  const nickMap = {};
  ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0]) nickMap[r[0]] = r[1] || ''; });

  // branch map
  const branchMap = {};
  ss.getSheetByName('Branches').getDataRange().getValues().slice(1)
    .filter(r => r[0]).forEach(b => {
      branchMap[String(b[0])] = { name: b[1], openTime: fmtTime(b[7]) };
    });

  // attendance วันนั้น เฉพาะ IN
  const lateAttRange = ss.getSheetByName('Attendance').getDataRange();
  const allAttRows   = lateAttRange.getValues().slice(1);
  const allAttDisp   = lateAttRange.getDisplayValues().slice(1);
  allAttRows.forEach((r, i) => { r._dt = allAttDisp[i][1]; });

  const attRows = allAttRows.filter(r => {
    let d;
    if (r[2] instanceof Date) { d = fmtDate(r[2]); }
    else { const p = new Date(String(r[2]).trim()); d = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
    return d === date && r[3] === 'IN';
  });

  const lateStaff = [];
  attRows.forEach(r => {
    const t = normalizeTimeDisplay(r._dt, r[1]);
    const branchId = String(r[4]);
    const branch = branchMap[branchId] || { name: branchId, openTime: '' };
    if (!branch.openTime || !branch.openTime.includes(':')) return;
    const lr = calcLate(t, branch.openTime);
    if (!lr.isLate) return;
    lateStaff.push({
      name:       r[0],
      nickname:   nickMap[r[0]] || '',
      branchId,
      branchName: branch.name,
      time:       t,
      lateMins:   lr.mins
    });
  });

  // เรียงจากสายมากไปน้อย
  lateStaff.sort((a,b) => b.lateMins - a.lateMins);
  return { success: true, date, total: lateStaff.length, staff: lateStaff };
}
// ─── Today Attendance (Public — for home page history) ──────────────────────
function getTodayAttendance() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const today     = fmtDate(new Date());
  const branchMap = {};
  ss.getSheetByName('Branches').getDataRange().getValues().slice(1)
    .filter(r => r[0]).forEach(b => { branchMap[String(b[0])] = b[1]; });
  const nickMap = {};
  ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .filter(r => r[0]).forEach(s => { nickMap[s[0]] = s[1] || ''; });

  const attSheet  = ss.getSheetByName('Attendance');
  const dataRange = attSheet.getDataRange();
  const values    = dataRange.getValues().slice(1);
  const displays  = dataRange.getDisplayValues().slice(1);

  const records = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    let d;
    if (r[2] instanceof Date) { d = fmtDate(r[2]); }
    else { const p = new Date(String(r[2]).trim()); d = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
    if (d !== today) continue;
    records.push({
      name:       r[0],
      nickname:   nickMap[r[0]] || '',
      type:       r[3],
      time:       normalizeTimeDisplay(displays[i][1], r[1]),
      branchName: branchMap[String(r[4])] || String(r[4])
    });
  }

  records.reverse();
  return { success: true, records };
}

// ─── Employee Full History ────────────────────────────────────────────────────
function getEmployeeHistory(data) {
  const name = data.name;
  if (!name) return { success: false, message: 'No name' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffRow = ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .find(r => r[0] === name);
  const nickname = staffRow ? (staffRow[1] || '') : '';

  const branchMap = {};
  ss.getSheetByName('Branches').getDataRange().getValues().slice(1)
    .filter(r => r[0]).forEach(b => { branchMap[String(b[0])] = b[1]; });

  const attRange = ss.getSheetByName('Attendance').getDataRange();
  const values   = attRange.getValues().slice(1);
  const displays = attRange.getDisplayValues().slice(1);

  const records = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (r[0] !== name) continue;
    let dateStr;
    if (r[2] instanceof Date) { dateStr = fmtDate(r[2]); }
    else { const p = new Date(String(r[2]).trim()); dateStr = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
    records.push({
      date:       dateStr,
      time:       normalizeTimeDisplay(displays[i][1], r[1]),
      type:       r[3],
      branchName: branchMap[String(r[4])] || String(r[4]),
      status:     r[9] || '',
      lateMins:   Number(r[10]) || 0
    });
  }

  records.reverse();
  return { success: true, name, nickname, records: records.slice(0, 90) };
}

// อ่านเวลาจาก display value ของ Sheets (ตรงกับที่เห็นใน sheet เสมอ)
function normalizeTimeDisplay(displayVal, rawVal) {
  const s = String(displayVal || '').trim();
  if (!s || !s.includes(':')) return fmtTime(rawVal);
  // รูปแบบ 12h: "2:30 PM" หรือ "2:30:25 PM"
  const ampm = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)/i);
  if (ampm) {
    let h = parseInt(ampm[1]);
    const m = ampm[2];
    if (ampm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
    if (ampm[3].toUpperCase() === 'AM' && h === 12) h = 0;
    return h.toString().padStart(2,'0') + ':' + m;
  }
  // รูปแบบ 24h: "14:30:25" หรือ "14:30"
  return s.length > 5 ? s.substring(0, 5) : s;
}

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
    const totalMins = Math.floor(num * 24 * 60); // floor ป้องกัน rounding error
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

// ─── Suspicious GPS Detection ─────────────────────────────────────────────────
function getSuspiciousStaff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nickMap = {};
  ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0]) nickMap[r[0]] = r[1] || ''; });

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  const cutoffStr = fmtDate(cutoff);

  const rows = ss.getSheetByName('Attendance').getDataRange().getValues().slice(1)
    .filter(r => {
      if (r[3] !== 'IN') return false;
      let d;
      if (r[2] instanceof Date) { d = fmtDate(r[2]); }
      else { const p = new Date(String(r[2]).trim()); d = isNaN(p) ? String(r[2]).trim() : fmtDate(p); }
      return d >= cutoffStr;
    });

  const byName = {};
  rows.forEach(r => {
    if (!byName[r[0]]) byName[r[0]] = [];
    // ใช้ตรวจ undefined/empty แทน || -1 เพราะ 0 เป็น falsy
    const acc = (r[11] !== undefined && r[11] !== '' && r[11] !== null)
      ? Number(r[11]) : -1;
    byName[r[0]].push(acc);
  });

  const suspicious = [];
  Object.entries(byName).forEach(([name, accList]) => {
    const issues = [];
    const valid  = accList.filter(a => a >= 0);
    if (!valid.length) return;

    // ตรวจ 1: Accuracy = 0
    const zeroCount = valid.filter(a => a === 0).length;
    if (zeroCount >= 3) issues.push({ level:'red', text:`Accuracy = 0 จำนวน ${zeroCount} ครั้ง (อาจใช้ Fake GPS)` });

    // ตรวจ 2: Accuracy > 200m เกิน 5 ครั้ง
    const highCount = valid.filter(a => a > 200).length;
    if (highCount >= 5) issues.push({ level:'yellow', text:`GPS อ่อนผิดปกติ ${highCount} ครั้ง (>200m)` });

    // ตรวจ 3: Accuracy ซ้ำกัน ±1m
    const nonZero = valid.filter(a => a > 0);
    let maxRepeat = 0, repeatVal = 0;
    nonZero.forEach(a => {
      const cnt = nonZero.filter(b => Math.abs(b-a) <= 1).length;
      if (cnt > maxRepeat) { maxRepeat = cnt; repeatVal = a; }
    });
    if (maxRepeat >= 3) issues.push({ level:'red', text:`Accuracy ซ้ำกัน ~${repeatVal}m จำนวน ${maxRepeat} ครั้ง (สัญญาณปลอม)` });

    if (issues.length > 0) {
      suspicious.push({ name, nickname: nickMap[name]||'', issues, checkCount: valid.length });
    }
  });

  suspicious.sort((a,b) => b.issues.filter(i=>i.level==='red').length - a.issues.filter(i=>i.level==='red').length);
  return { success:true, suspicious };
}
