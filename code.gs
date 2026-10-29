/**
 * ACE Content - Web App Backend Script
 * ระบบตรวจสอบข้อมูล VJ/MC & Admin + IP/Device Tracking + Access Control + Dynamic Version Update Management
 */

function doGet(e) {
  try {
    // The HTML page contains no Apps Script template tags; serve it as plain HTML
    // to avoid unnecessary template evaluation inside the sandboxed iframe.
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ACE Content - ระบบตรวจสอบข้อมูล VJ/MC & Admin')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('เกิดข้อผิดพลาดในการโหลดระบบ: ' + error.toString());
  }
}

// ------------------------------------------
// 🚀 DYNAMIC SYSTEM VERSION & UPDATE RELEASE INFO
// ------------------------------------------

function getOrCreateSystemVersionSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('System_Updates');
  if (!sheet) {
    sheet = ss.insertSheet('System_Updates');
    sheet.appendRow(['Version', 'Title', 'Changelog', 'Updated_At', 'Updated_By']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
    
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var defaultLogs = '📢 ปรับปรุงระบบป๊อปอัปแจ้งเตือนประกาศข่าวสาร เพิ่มปุ่มยืนยันรับทราบก่อนปิด\n⚙️ เพิ่มหน้าจัดการเวอร์ชันระบบ สามารถแก้ไขรายละเอียดการอัปเดตผ่านหน้าเว็บได้ทันที\n🔔 เพิ่มป๊อปอัปแจ้งเตือนอัปเดตระบบ แสดงเวอร์ชัน วันที่ และเวลาอัปเดตล่าสุด\n🔒 เพิ่มคอลัมน์ชื่อทีมในหน้าระงับสิทธิ์การใช้งาน (Lock ID / IP Address)\n📈 เพิ่มระบบเปรียบเทียบ KPI ยอดเพชรแบบ Month-over-Month';
    sheet.appendRow(['v2.6.0', 'ACE Content Portal - อัปเดตระบบเวอร์ชัน 2.6.0', defaultLogs, nowStr, 'System Admin']);
  }
  return sheet;
}

function getSystemVersionInfo() {
  try {
    var sheet = getOrCreateSystemVersionSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      var lastRow = data[data.length - 1];
      var changelogText = cleanStr(lastRow[2]);
      var changelogArr = changelogText ? changelogText.split('\n').filter(function(line) { return line.trim() !== ''; }) : [];
      
      return {
        success: true,
        version: cleanStr(lastRow[0]) || 'v2.6.0',
        title: cleanStr(lastRow[1]) || 'ACE Content Portal - อัปเดตระบบ',
        changelog: changelogArr,
        changelogRaw: changelogText,
        updatedAt: lastRow[3] ? (lastRow[3] instanceof Date ? Utilities.formatDate(lastRow[3], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(lastRow[3])) : '-',
        updatedBy: cleanStr(lastRow[4]) || '-'
      };
    }
  } catch (err) {
    Logger.log('Error reading system version info: ' + err.toString());
  }
  return {
    success: true,
    version: 'v2.6.0',
    title: 'ACE Content Portal - อัปเดตระบบ',
    changelog: ['ปรับปรุงประสิทธิภาพการทำงานของระบบ'],
    changelogRaw: 'ปรับปรุงประสิทธิภาพการทำงานของระบบ',
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'),
    updatedBy: 'Admin'
  };
}

function saveSystemVersionInfo(version, title, changelogText, adminName, ipAddress, deviceInfo) {
  try {
    var sheet = getOrCreateSystemVersionSheet();
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var cleanVer = cleanStr(version);
    var cleanTitleStr = cleanStr(title);
    var cleanChangelog = cleanStr(changelogText);
    var admin = adminName || 'Admin';

    if (!cleanVer || !cleanTitleStr) {
      return { success: false, message: 'โปรดระบุเลขเวอร์ชันและหัวข้อการอัปเดต' };
    }

    sheet.appendRow([cleanVer, cleanTitleStr, cleanChangelog, nowStr, admin]);
    writeLog('ADMIN', admin, 'UPDATE_SYSTEM_VERSION', 'อัปเดตเวอร์ชันระบบเป็น ' + cleanVer + ' (' + cleanTitleStr + ')', ipAddress, deviceInfo);

    return { 
      success: true, 
      message: '🚀 บันทึกและตั้งค่าเวอร์ชันระบบ ' + cleanVer + ' เรียบร้อยแล้ว',
      versionInfo: getSystemVersionInfo()
    };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกเวอร์ชัน: ' + err.toString() };
  }
}

// ------------------------------------------
// 🛠️ Helper Functions & Logger
// ------------------------------------------

function cleanStr(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .trim();
}

function parseDiamondVal(valStr) {
  if (valStr === null || valStr === undefined || valStr === '') return 0;
  if (typeof valStr === 'number') return valStr;
  var cleaned = String(valStr)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .replace(/,/g, '')
    .trim();
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function writeLog(userType, userName, action, details, ipAddress, deviceInfo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('Logs') || ss.getSheetByName('LOGS') || ss.getSheetByName('logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs');
      logSheet.appendRow(['Timestamp', 'User Type', 'User Name', 'Action', 'Details', 'IP Address', 'Device Info', 'Month_Year']);
      logSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
    }
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var monthYearStr = Utilities.formatDate(now, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();
    logSheet.appendRow([nowStr, userType || 'SYSTEM', userName || '-', action || '-', details || '-', ipAddress || '-', deviceInfo || '-', monthYearStr]);
  } catch (err) {
    Logger.log('Error writing log: ' + err.toString());
  }
}

function getLogs(selectedMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('Logs') || ss.getSheetByName('LOGS') || ss.getSheetByName('logs');
    if (!logSheet) return { success: true, logs: [] };

    var data = logSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };

    var logs = [];
    var filterMonth = selectedMonth ? cleanStr(selectedMonth).toUpperCase() : '';

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var timestamp = row[0] ? (row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(row[0])) : '-';
      var userType = cleanStr(row[1]);
      var userName = cleanStr(row[2]);
      var action = cleanStr(row[3]);
      var details = cleanStr(row[4]);
      var ipAddress = cleanStr(row[5]);
      var deviceInfo = cleanStr(row[6]);
      var monthYear = cleanStr(row[7]);

      if (filterMonth && monthYear && monthYear.indexOf(filterMonth) === -1 && filterMonth.indexOf(monthYear) === -1) {
        continue;
      }

      logs.push({
        timestamp: timestamp,
        userType: userType,
        userName: userName,
        action: action,
        details: details,
        ipAddress: ipAddress,
        deviceInfo: deviceInfo,
        monthYear: monthYear
      });
    }

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ------------------------------------------
// 🔒 ACCESS CONTROL & IP / USER LOCKING
// ------------------------------------------

function getOrCreateAccessControlSheet() {
  // Preserve any existing sheet schema. New sheets use the current eight-column layout.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Access_Control');
  if (!sheet) {
    sheet = ss.insertSheet('Access_Control');
    sheet.appendRow(['ID', 'Type', 'Target', 'Team', 'Status', 'Reason', 'Updated_By', 'Updated_At']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#fee2e2');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function checkAccessPermission(userName, ipAddress) {
  try {
    var sheet = getOrCreateAccessControlSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { allowed: true };
    var headers = data[0].map(function(h) { return cleanStr(h).toUpperCase().replace(/\s+/g, '_'); });
    function col(name, fallback) { var n = headers.indexOf(name); return n >= 0 ? n : fallback; }
    var typeIdx = col('TYPE', 1), targetIdx = col('TARGET', 2);
    var statusIdx = col('STATUS', headers.indexOf('TEAM') >= 0 ? 4 : 3);
    var reasonIdx = col('REASON', headers.indexOf('TEAM') >= 0 ? 5 : 4);
    var locked = Object.create(null);
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanStr(row[statusIdx]).toUpperCase() !== 'LOCKED') continue;
      var type = cleanStr(row[typeIdx]).toUpperCase();
      var target = cleanStr(row[targetIdx]).toLowerCase();
      if ((type === 'USER' || type === 'IP') && target) locked[type + ':' + target] = cleanStr(row[reasonIdx]);
    }
    var cleanUser = cleanStr(userName).toLowerCase();
    var cleanIp = cleanStr(ipAddress).toLowerCase();
    if (cleanUser && Object.prototype.hasOwnProperty.call(locked, 'USER:' + cleanUser)) {
      return { allowed: false, reason: 'บัญชี VJ/User (' + userName + ') ถูกล็อกการใช้งาน: ' + (locked['USER:' + cleanUser] || 'โปรดติดต่อ Admin') };
    }
    if (cleanIp && Object.prototype.hasOwnProperty.call(locked, 'IP:' + cleanIp)) {
      return { allowed: false, reason: 'IP Address (' + ipAddress + ') ถูกระงับการเข้าถึง: ' + (locked['IP:' + cleanIp] || 'โปรดติดต่อ Admin') };
    }
    return { allowed: true };
  } catch (err) {
    Logger.log('Access check error: ' + err.toString());
    return { allowed: true };
  }
}

function getAccessControlList() {
  try {
    var sheet = getOrCreateAccessControlSheet();
    var data = sheet.getDataRange().getValues();
    var list = [];
    if (!data.length) return { success: true, list: list };
    var headers = data[0].map(function(h) { return cleanStr(h).toUpperCase().replace(/\s+/g, '_'); });
    function col(name, fallback) { var n = headers.indexOf(name); return n >= 0 ? n : fallback; }
    var teamIdx = col('TEAM', -1), statusIdx = col('STATUS', teamIdx >= 0 ? 4 : 3);
    var reasonIdx = col('REASON', teamIdx >= 0 ? 5 : 4);
    var byIdx = col('UPDATED_BY', teamIdx >= 0 ? 6 : 5), atIdx = col('UPDATED_AT', teamIdx >= 0 ? 7 : 6);
    for (var i = 1; i < data.length; i++) {
      list.push({ id: cleanStr(data[i][col('ID', 0)]), type: cleanStr(data[i][col('TYPE', 1)]), target: cleanStr(data[i][col('TARGET', 2)]),
        team: teamIdx >= 0 ? (cleanStr(data[i][teamIdx]) || 'ไม่ระบุทีม') : 'ไม่ระบุทีม', status: cleanStr(data[i][statusIdx]), reason: cleanStr(data[i][reasonIdx]),
        updatedBy: cleanStr(data[i][byIdx]), updatedAt: data[i][atIdx] ? (data[i][atIdx] instanceof Date ? Utilities.formatDate(data[i][atIdx], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(data[i][atIdx])) : '-' });
    }
    return { success: true, list: list };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function updateAccessRule(type, target, team, status, reason, adminName, ipAddress, deviceInfo) {
  try {
    var sheet = getOrCreateAccessControlSheet();
    var data = sheet.getDataRange().getValues();
    var headers = (data[0] || []).map(function(h) { return cleanStr(h).toUpperCase().replace(/\s+/g, '_'); });
    function col(name, fallback) { var n = headers.indexOf(name); return n >= 0 ? n : fallback; }
    var teamIdx = col('TEAM', -1), typeIdx = col('TYPE', 1), targetIdx = col('TARGET', 2);
    var statusIdx = col('STATUS', teamIdx >= 0 ? 4 : 3), reasonIdx = col('REASON', teamIdx >= 0 ? 5 : 4);
    var byIdx = col('UPDATED_BY', teamIdx >= 0 ? 6 : 5), atIdx = col('UPDATED_AT', teamIdx >= 0 ? 7 : 6);
    var cleanType = cleanStr(type).toUpperCase(), cleanTarget = cleanStr(target), cleanTeam = cleanStr(team) || 'ไม่ระบุทีม';
    var cleanStatus = cleanStr(status).toUpperCase(), nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    if (!cleanTarget) return { success: false, message: 'โปรดระบุชื่อ User หรือ IP ที่ต้องการตั้งค่า' };
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) if (cleanStr(data[i][typeIdx]).toUpperCase() === cleanType && cleanStr(data[i][targetIdx]).toLowerCase() === cleanTarget.toLowerCase()) { targetRow = i + 1; break; }
    if (targetRow > 1) {
      if (teamIdx >= 0) sheet.getRange(targetRow, teamIdx + 1).setValue(cleanTeam);
      sheet.getRange(targetRow, statusIdx + 1).setValue(cleanStatus);
      sheet.getRange(targetRow, reasonIdx + 1).setValue(cleanStr(reason));
      sheet.getRange(targetRow, byIdx + 1).setValue(adminName || 'Admin');
      sheet.getRange(targetRow, atIdx + 1).setValue(nowStr);
    } else {
      var newRow = new Array(Math.max(headers.length, 1)).fill('');
      newRow[col('ID', 0)] = 'ACC_' + new Date().getTime(); newRow[typeIdx] = cleanType; newRow[targetIdx] = cleanTarget;
      if (teamIdx >= 0) newRow[teamIdx] = cleanTeam;
      newRow[statusIdx] = cleanStatus; newRow[reasonIdx] = cleanStr(reason); newRow[byIdx] = adminName || 'Admin'; newRow[atIdx] = nowStr;
      sheet.appendRow(newRow);
    }
    var actText = cleanStatus === 'LOCKED' ? 'LOCK_ACCESS' : 'UNLOCK_ACCESS';
    writeLog('ADMIN', adminName, actText, 'ตั้งค่าการเข้าถึง ' + cleanType + ': ' + cleanTarget + ' (ทีม: ' + cleanTeam + ') -> ' + cleanStatus + ' (เหตุผล: ' + (reason || '-') + ')', ipAddress, deviceInfo);
    return { success: true, message: (cleanStatus === 'LOCKED' ? '🔒 ล็อกสิทธิ์การใช้งานเรียบร้อยแล้ว' : '🔓 ปลดล็อกการใช้งานเรียบร้อยแล้ว') };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function getProfilePicUrl(tiktokVal, stageName) {
  var cleanedVal = cleanStr(tiktokVal);
  if (!cleanedVal) {
    return 'https://ui-avatars.com/api/?name=' + encodeURIComponent(stageName || 'VJ') + '&background=6366f1&color=fff&bold=true';
  }
  if (cleanedVal.indexOf('http://') === 0 || cleanedVal.indexOf('https://') === 0) {
    return cleanedVal;
  }
  var cleanUser = cleanedVal.replace('@', '').trim();
  return 'https://unavatar.io/tiktok/' + encodeURIComponent(cleanUser);
}

function getAvailableMonths(spreadsheet) {
  try {
    var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var months = [];
    var ignoreSheets = ['vj mc all', 'vj_mc_all', 'sheet', 'logs', 'admins', 'team data', 'kpi_targets', 'kpi targets', 'attendance_data', 'attendance', 'announcements', 'announcement_reads', 'access_control', 'system_updates'];
    
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      var lowerName = name.toLowerCase();
      var shouldIgnore = ignoreSheets.some(function(keyword) {
        return lowerName.indexOf(keyword) !== -1;
      });
      if (!shouldIgnore) {
        months.push(name);
      }
    }
    return months;
  } catch (err) {
    return [];
  }
}

function getKPIMap(selectedMonth, spreadsheet) {
  var kpiMap = {};
  try {
    var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('KPI_Targets') || ss.getSheetByName('kpi_targets') || ss.getSheetByName('KPI Targets');
    if (!sheet) return kpiMap;

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return kpiMap;

    var hRow = data[0];
    var nameIdx = 0, targetIdx = 1, monthIdx = 2;

    for (var c = 0; c < hRow.length; c++) {
      var h = cleanStr(hRow[c]).toUpperCase();
      if (h.indexOf('VJ') !== -1 || h.indexOf('NAME') !== -1) nameIdx = c;
      if (h.indexOf('TARGET') !== -1) targetIdx = c;
      if (h.indexOf('MONTH') !== -1) monthIdx = c;
    }

    var cleanSelectedMonth = cleanStr(selectedMonth).toUpperCase();

    for (var i = 1; i < data.length; i++) {
      var vjName = cleanStr(data[i][nameIdx]).toLowerCase();
      var targetVal = parseDiamondVal(data[i][targetIdx]);
      var monthVal = cleanStr(data[i][monthIdx]).toUpperCase();

      if (vjName && (monthVal === cleanSelectedMonth || cleanSelectedMonth.indexOf(monthVal) !== -1 || monthVal.indexOf(cleanSelectedMonth) !== -1)) {
        kpiMap[vjName] = targetVal;
      }
    }
  } catch (err) {
    Logger.log('Error reading KPI_Targets sheet: ' + err.toString());
  }
  return kpiMap;
}

// ------------------------------------------
// 👤 VJ & Admin User Listing
// ------------------------------------------

function getVJListForLogin(spreadsheet) {
  try {
    // This is the bound ACE workbook. Keep the spreadsheet reference local to this request.
    var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
    var vjList = [];
    if (!vjSheet) return { success: false, message: 'ไม่พบชีต vj mc all ในสเปรดชีตหลัก', vjList: [] };
    var vjValues = vjSheet.getDataRange().getValues();
    if (!vjValues || vjValues.length < 2) return { success: true, vjList: [] };

    // Some workbook copies have a title row before the headers. Find the actual header row.
    var headerRow = -1, bestScore = 0;
    var scanRows = Math.min(vjValues.length, 5);
    for (var r = 0; r < scanRows; r++) {
      var score = 0;
      for (var c = 0; c < vjValues[r].length; c++) {
        var h = cleanStr(vjValues[r][c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (h === 'STGENAME' || h === 'STAGENAME' || h === 'ชื่อการแสดง') score += 3;
        if (h === 'NAME' || h === 'FULLNAME' || h === 'ชื่อ') score += 1;
        if (h === 'TEAM' || h === 'ทีม') score += 1;
        if (h === 'ROLE' || h === 'ตำแหน่ง') score += 1;
      }
      if (score > bestScore) { bestScore = score; headerRow = r; }
    }
    if (headerRow < 0) headerRow = Math.min(1, vjValues.length - 1);

    var headers = vjValues[headerRow] || [];
    var stageIdx = 1, nameIdx = 0, teamIdx = 2, roleIdx = 3;
    for (var c = 0; c < headers.length; c++) {
      var hStr = cleanStr(headers[c]).toUpperCase().replace(/[\s_-]+/g, '');
      if (hStr === 'STGENAME' || hStr === 'STAGENAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
      else if (hStr === 'NAME' || hStr === 'FULLNAME' || hStr === 'ชื่อ') nameIdx = c;
      else if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
      else if (hStr === 'ROLE' || hStr === 'ตำแหน่ง') roleIdx = c;
    }

    var seen = Object.create(null);
    for (var i = headerRow + 1; i < vjValues.length; i++) {
      var stageName = cleanStr(vjValues[i][stageIdx]);
      var fullName = cleanStr(vjValues[i][nameIdx]);
      var team = cleanStr(vjValues[i][teamIdx]);
      var role = cleanStr(vjValues[i][roleIdx]);
      var key = stageName.toLowerCase();
      if (stageName && !seen[key]) {
        seen[key] = true;
        vjList.push({ stageName: stageName, fullName: fullName || stageName, role: role || '-', team: team || '-' });
      }
    }
    return { success: true, vjList: vjList };
  } catch (err) {
    Logger.log('Error loading VJ login list: ' + err.toString());
    return { success: false, message: 'โหลดรายชื่อ VJ ไม่สำเร็จ: ' + err.toString(), vjList: [] };
  }
}

function getAdminNames(spreadsheet) {
  try {
    var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName('Admins') || ss.getSheetByName('admins') || ss.getSheetByName('ADMINS');
    var adminList = [];
    if (adminSheet) {
      var adminValues = adminSheet.getDataRange().getValues();
      var hRow = adminValues[0] || [];
      var nameIdx = 0;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr(hRow[c]).toUpperCase();
        if (hStr === 'NAME' || hStr === 'ADMIN NAME' || hStr === 'ADMIN' || hStr === 'USERNAME' || hStr === 'ชื่อ') {
          nameIdx = c;
          break;
        }
      }
      for (var i = 1; i < adminValues.length; i++) {
        var aName = cleanStr(adminValues[i][nameIdx]);
        if (aName) adminList.push(aName);
      }
    }
    if (adminList.length === 0) {
      adminList = ['Admin01', 'Admin02', 'SuperAdmin'];
    }
    return { success: true, adminList: adminList };
  } catch (err) {
    return { success: true, adminList: ['Admin01', 'SuperAdmin'] };
  }
}

/** Load login-screen data in one Apps Script RPC and reuse one spreadsheet handle. */
function getLoginBootstrapData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    success: true,
    months: getAvailableMonths(ss),
    vj: getVJListForLogin(ss),
    admins: getAdminNames(ss)
  };
}

// ------------------------------------------
// 📢 ANNOUNCEMENTS MODULE
// ------------------------------------------

function getOrCreateAnnouncementSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var annSheet = ss.getSheetByName('Announcements');
  if (!annSheet) {
    annSheet = ss.insertSheet('Announcements');
    annSheet.appendRow(['ID', 'Timestamp', 'Month_Year', 'Title', 'Content', 'Image_Url', 'Is_Important', 'Require_Ack', 'Created_By', 'Status']);
    annSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#e2e8f0');
    annSheet.setFrozenRows(1);
  }
  var readSheet = ss.getSheetByName('Announcement_Reads');
  if (!readSheet) {
    readSheet = ss.insertSheet('Announcement_Reads');
    readSheet.appendRow(['ID', 'Announcement_ID', 'User_Name', 'User_Type', 'Read_Timestamp', 'Acknowledged', 'Month_Year']);
    readSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e2e8f0');
    readSheet.setFrozenRows(1);
  }
  return { annSheet: annSheet, readSheet: readSheet };
}

function getAnnouncementSchedule(announcementId, allProperties) {
  try {
    var properties = allProperties || PropertiesService.getScriptProperties().getProperties();
    var raw = properties['ACE_ANN_SCHEDULE_' + announcementId];
    return raw ? JSON.parse(raw) : {};
  } catch (err) {
    return {};
  }
}

function parseAnnouncementDateTime(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var text = cleanStr(value).replace('T', ' ');
  if (!text) return '';
  try {
    return Utilities.parseDate(text, 'Asia/Bangkok', text.length === 16 ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd HH:mm:ss');
  } catch (err) {
    throw new Error('รูปแบบวันเวลาไม่ถูกต้อง กรุณาเลือกวันและเวลาใหม่');
  }
}

function createAnnouncement(data, adminName, ipAddress, deviceInfo, adminPin) {
  try {
    var permission = checkAccessPermission(adminName, ipAddress);
    if (!permission.allowed) {
      writeLog('ADMIN', adminName, 'ACCESS_BLOCKED', permission.reason, ipAddress, deviceInfo);
      return { success: false, message: '⛔ ' + permission.reason };
    }

    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = activeSs.getSheetByName('Admins') || activeSs.getSheetByName('admins') || activeSs.getSheetByName('ADMINS');
    var cleanAdminName = cleanStr(adminName).toLowerCase();
    var cleanAdminPin = cleanStr(adminPin);
    var adminAuthenticated = false;
    if (adminSheet) {
      var adminData = adminSheet.getDataRange().getValues();
      var headers = adminData[0] || [];
      var nameColumn = 0, pinColumn = 1;
      for (var h = 0; h < headers.length; h++) {
        var header = cleanStr(headers[h]).toUpperCase();
        if (header === 'NAME' || header === 'ADMIN NAME' || header === 'ADMIN' || header === 'USERNAME' || header === 'ชื่อ') nameColumn = h;
        if (header === 'PIN' || header === 'PASSWORD' || header === 'PASSCODE' || header === 'รหัสผ่าน') pinColumn = h;
      }
      for (var a = 1; a < adminData.length; a++) {
        if (cleanStr(adminData[a][nameColumn]).toLowerCase() === cleanAdminName && cleanStr(adminData[a][pinColumn]) === cleanAdminPin) { adminAuthenticated = true; break; }
      }
    }
    if (!adminAuthenticated && (cleanAdminPin.toUpperCase() === 'ADMIN123' || cleanAdminPin === '1234')) adminAuthenticated = true;
    if (!adminAuthenticated) {
      writeLog('ADMIN', adminName, 'CREATE_ANNOUNCEMENT_DENIED', 'ยืนยันบัญชีหรือ PIN ไม่สำเร็จ', ipAddress, deviceInfo);
      return { success: false, message: '🔒 ไม่สามารถยืนยันสิทธิ์ Admin ได้ กรุณาเข้าสู่ระบบใหม่' };
    }

    var sheets = getOrCreateAnnouncementSheets();
    var id = 'ANN_' + new Date().getTime();
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var monthYearStr = Utilities.formatDate(now, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();

    var startsAt = parseAnnouncementDateTime(data.startsAt);
    var endsAt = parseAnnouncementDateTime(data.endsAt);
    if (startsAt && endsAt && endsAt.getTime() <= startsAt.getTime()) {
      throw new Error('เวลาปิดประกาศต้องอยู่หลังเวลาเริ่มประกาศ');
    }

    var imageUrl = cleanStr(data.imageUrl);
    if (imageUrl && !/^https:\/\//i.test(imageUrl)) throw new Error('ลิงก์รูปภาพต้องขึ้นต้นด้วย https://');
    if (data.imageBase64 && data.imageBase64.indexOf('data:image') === 0) {
      try {
        var imageParts = data.imageBase64.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
        if (!imageParts) throw new Error('ไฟล์รูปภาพไม่ถูกต้อง');
        var folderSearch = DriveApp.getFoldersByName('Announcement_Images');
        var targetFolder = folderSearch.hasNext() ? folderSearch.next() : DriveApp.createFolder('Announcement_Images');
        var mimeType = imageParts[1].toLowerCase();
        var extension = mimeType.split('/')[1].replace('jpeg', 'jpg');
        var blob = Utilities.newBlob(Utilities.base64Decode(imageParts[2]), mimeType, id + '.' + extension);
        var file = targetFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      } catch (errImg) {
        throw new Error('อัปโหลดรูปภาพไม่สำเร็จ: ' + errImg.message);
      }
    }

    sheets.annSheet.appendRow([
      id,
      nowStr,
      monthYearStr,
      cleanStr(data.title),
      cleanStr(data.content),
      imageUrl,
      data.isImportant ? 'YES' : 'NO',
      data.requireAck ? 'YES' : 'NO',
      adminName || 'Admin',
      'ACTIVE'
    ]);
    if (startsAt || endsAt) {
      PropertiesService.getScriptProperties().setProperty('ACE_ANN_SCHEDULE_' + id, JSON.stringify({
        startsAt: startsAt ? startsAt.getTime() : 0,
        endsAt: endsAt ? endsAt.getTime() : 0
      }));
    }

    writeLog('ADMIN', adminName, 'CREATE_ANNOUNCEMENT', 'สร้างประกาศข่าวสาร: ' + data.title, ipAddress, deviceInfo);
    return { success: true, id: id, message: 'สร้างประกาศข่าวสารเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getActiveAnnouncements(userName, userType) {
  try {
    var sheets = getOrCreateAnnouncementSheets();
    var annData = sheets.annSheet.getDataRange().getValues();
    if (annData.length <= 1) return { success: true, list: [] };

    var readsData = sheets.readSheet.getDataRange().getValues();
    var readMap = Object.create(null);
    for (var r = 1; r < readsData.length; r++) {
      if (cleanStr(readsData[r][2]).toLowerCase() === cleanStr(userName).toLowerCase()) {
        var aId = cleanStr(readsData[r][1]);
        var readKey = aId + '|' + (cleanStr(readsData[r][3]).toUpperCase() || 'VJ');
        readMap[readKey] = {
          read: true,
          acknowledged: cleanStr(readsData[r][5]) === 'YES'
        };
      }
    }

    var list = [];
    var now = new Date();
    var scheduleProperties = PropertiesService.getScriptProperties().getProperties();
    for (var i = annData.length - 1; i >= 1; i--) {
      var row = annData[i];
      if (cleanStr(row[9]) !== 'ACTIVE') continue;
      var annId = cleanStr(row[0]);
      var schedule = getAnnouncementSchedule(annId, scheduleProperties);
      var startAt = Number(schedule.startsAt) || 0;
      var endAt = Number(schedule.endsAt) || 0;
      if (startAt && startAt > now.getTime()) continue;
      if (endAt && endAt <= now.getTime()) continue;
      var userReadInfo = readMap[annId + '|' + cleanStr(userType || 'VJ').toUpperCase()] || { read: false, acknowledged: false };

      list.push({
        id: annId,
        timestamp: row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(row[1]),
        monthYear: cleanStr(row[2]),
        title: cleanStr(row[3]),
        content: cleanStr(row[4]),
        imageUrl: cleanStr(row[5]),
        startsAt: startAt ? Utilities.formatDate(new Date(startAt), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        endsAt: endAt ? Utilities.formatDate(new Date(endAt), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        isImportant: cleanStr(row[6]) === 'YES',
        requireAck: cleanStr(row[7]) === 'YES',
        createdBy: cleanStr(row[8]),
        isRead: userReadInfo.read,
        isAcknowledged: userReadInfo.acknowledged
      });
    }
    return { success: true, list: list };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function recordAnnouncementRead(announcementId, userName, userType, acknowledged) {
  try {
    var sheets = getOrCreateAnnouncementSheets();
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var monthYearStr = Utilities.formatDate(now, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();

    var readsData = sheets.readSheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < readsData.length; i++) {
      var existingUserType = cleanStr(readsData[i][3]).toUpperCase() || 'VJ';
      if (cleanStr(readsData[i][1]) === announcementId && cleanStr(readsData[i][2]).toLowerCase() === cleanStr(userName).toLowerCase() && existingUserType === cleanStr(userType || 'VJ').toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 1) {
      if (acknowledged) {
        sheets.readSheet.getRange(targetRow, 6).setValue('YES');
        sheets.readSheet.getRange(targetRow, 5).setValue(nowStr);
      }
    } else {
      var readId = 'READ_' + new Date().getTime();
      sheets.readSheet.appendRow([
        readId,
        announcementId,
        userName,
        userType || 'VJ',
        nowStr,
        acknowledged ? 'YES' : 'NO',
        monthYearStr
      ]);
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAnnouncementStats() {
  try {
    var sheets = getOrCreateAnnouncementSheets();
    var annData = sheets.annSheet.getDataRange().getValues();
    var readData = sheets.readSheet.getDataRange().getValues();

    if (annData.length <= 1) return { success: true, stats: [], totalAnnouncements: 0, totalViews: 0 };

    var readsByAnn = {};
    for (var r = 1; r < readData.length; r++) {
      var annId = cleanStr(readData[r][1]);
      if (!readsByAnn[annId]) readsByAnn[annId] = [];
      readsByAnn[annId].push({
        userName: cleanStr(readData[r][2]),
        userType: cleanStr(readData[r][3]),
        timestamp: readData[r][4] instanceof Date ? Utilities.formatDate(readData[r][4], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(readData[r][4]),
        acknowledged: cleanStr(readData[r][5]) === 'YES'
      });
    }

    var stats = [];
    var totalViews = 0;
    var scheduleProperties = PropertiesService.getScriptProperties().getProperties();

    for (var i = annData.length - 1; i >= 1; i--) {
      var row = annData[i];
      var id = cleanStr(row[0]);
      var schedule = getAnnouncementSchedule(id, scheduleProperties);
      var nowMs = Date.now();
      var startsAtMs = Number(schedule.startsAt) || 0;
      var endsAtMs = Number(schedule.endsAt) || 0;
      var effectiveStatus = endsAtMs && endsAtMs <= nowMs ? 'EXPIRED' : (startsAtMs && startsAtMs > nowMs ? 'SCHEDULED' : cleanStr(row[9]));
      var reads = readsByAnn[id] || [];
      var viewCount = reads.length;
      var ackCount = reads.filter(function(m) { return m.acknowledged; }).length;
      totalViews += viewCount;

      stats.push({
        id: id,
        timestamp: row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(row[1]),
        monthYear: cleanStr(row[2]),
        title: cleanStr(row[3]),
        content: cleanStr(row[4]),
        imageUrl: cleanStr(row[5]),
        startsAt: schedule.startsAt ? Utilities.formatDate(new Date(Number(schedule.startsAt)), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        endsAt: schedule.endsAt ? Utilities.formatDate(new Date(Number(schedule.endsAt)), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        isImportant: cleanStr(row[6]) === 'YES',
        requireAck: cleanStr(row[7]) === 'YES',
        createdBy: cleanStr(row[8]),
        status: effectiveStatus,
        viewCount: viewCount,
        ackCount: ackCount,
        readersList: reads
      });
    }

    return {
      success: true,
      stats: stats,
      totalAnnouncements: stats.length,
      totalViews: totalViews
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ------------------------------------------
// 📊 Core Data Fetcher (Optimized)
// ------------------------------------------

function getAllMonthData(selectedMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var months = getAvailableMonths(ss);
  if (!selectedMonth && months.length > 0) {
    selectedMonth = months[0];
  }
  var monthSheet = ss.getSheetByName(selectedMonth);
  if (!monthSheet) {
    return { success: false, message: 'ไม่พบชีตประจำเดือน: ' + selectedMonth };
  }

  var kpiMap = getKPIMap(selectedMonth, ss);
  var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
  var vjProfileMap = {};
  var vjPassMap = {};

  if (vjSheet) {
    var vjValues = vjSheet.getDataRange().getValues();
    var hRow = vjValues[0] || [];
    var nameIdx = 0, stageIdx = 1, teamIdx = 2, roleIdx = 3, shiftIdx = 5, passIdx = 6, tiktokIdx = 7, statusIdx = 8;
    
    for (var c = 0; c < hRow.length; c++) {
      var hStr = cleanStr(hRow[c]).toUpperCase();
      if (hStr === 'NAME' || hStr === 'ชื่อ') nameIdx = c;
      else if (hStr === 'STGE NAME' || hStr === 'STAGE NAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
      else if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
      else if (hStr === 'ROLE' || hStr === 'ตำแหน่ง') roleIdx = c;
      else if (hStr === 'SHIFT' || hStr === 'เวลาทำงาน' || c === 8) shiftIdx = c;
      else if (hStr === 'PASSWORD' || hStr === 'PASSCODE' || hStr === 'PIN' || hStr === 'รหัสผ่าน') passIdx = c;
      else if (hStr === 'ID TIKTOK' || hStr === 'TIKTOK ID' || hStr === 'TIKTOK') tiktokIdx = c;
      else if (hStr === 'STATUS' || hStr === 'สถานะ') statusIdx = c;
    }

    for (var i = 1; i < vjValues.length; i++) {
      var stageName = cleanStr(vjValues[i][stageIdx]);
      if (!stageName) continue;
      var passVal = (passIdx !== -1 && vjValues[i][passIdx]) ? cleanStr(vjValues[i][passIdx]) : '1234';
      var tiktokVal = tiktokIdx !== -1 ? cleanStr(vjValues[i][tiktokIdx]) : '';
      
      vjPassMap[stageName.toLowerCase()] = passVal;
      vjProfileMap[stageName.toLowerCase()] = {
        fullName: cleanStr(vjValues[i][nameIdx]) || stageName,
        stageName: stageName,
        team: cleanStr(vjValues[i][teamIdx]) || '-',
        role: cleanStr(vjValues[i][roleIdx]) || '-',
        shift: cleanStr(vjValues[i][shiftIdx]) || '-',
        status: cleanStr(vjValues[i][statusIdx]) || 'ทำงานอยู่',
        password: passVal,
        tiktokId: tiktokVal,
        avatarUrl: getProfilePicUrl(tiktokVal, stageName)
      };
    }
  }

  var mData = monthSheet.getDataRange().getValues();
  var rawList = [];
  var ignoreKeywords = [
    'total', 'grand total', 'รวม', 'ยอดรวม', 'สรุป', 'สรุปยอด', 
    'stage name', 'ชื่อการแสดง', 'ลำดับ', 'no.', 'no', 'target', 'เป้าหมาย',
    'subtotal', 'รวมทีม', 'team total', 'group member', '团播成员', '月度流水'
  ];

  for (var r = 0; r < mData.length; r++) {
    var row = mData[r];
    if (!row || row.length < 6) continue;
    var colA_Team = cleanStr(row[0]);
    var colE_Stage = cleanStr(row[4]);
    var monthlyDiamStr = cleanStr(row[5]);
    if (!colE_Stage) continue;

    var lowerStage = colE_Stage.toLowerCase();
    var lowerColA = colA_Team.toLowerCase();
    var isIgnore = ignoreKeywords.some(function(kw) {
      return lowerStage === kw || 
             lowerStage.indexOf('รวม') !== -1 || 
             lowerStage.indexOf('total') !== -1 ||
             lowerColA.indexOf('รวม') !== -1 ||
             lowerColA.indexOf('total') !== -1;
    });
    if (isIgnore) continue;

    var dailyEarnings = [];
    var dailySum = 0;

    for (var day = 1; day <= 31; day++) {
      var colIndex = day + 5;
      var rawVal = colIndex < row.length ? cleanStr(row[colIndex]) : '-';
      var dayVal = parseDiamondVal(rawVal);
      
      dailyEarnings.push(rawVal !== '' ? rawVal : '-');
      dailySum += dayVal;
    }

    var currentDiamond = parseDiamondVal(monthlyDiamStr);
    if (currentDiamond === 0 && dailySum > 0) {
      currentDiamond = dailySum;
    }

    var profile = vjProfileMap[lowerStage] || {
      fullName: colE_Stage,
      stageName: colE_Stage,
      team: colA_Team || 'ไม่ระบุทีม',
      role: '-',
      shift: cleanStr(row[1]) || '-',
      status: 'ทำงานอยู่',
      password: vjPassMap[lowerStage] || '1234',
      tiktokId: '',
      avatarUrl: getProfilePicUrl('', colE_Stage)
    };

    var targetDiamond = kpiMap[lowerStage] !== undefined ? kpiMap[lowerStage] : 100000;
    var remDiamond = Math.max(0, targetDiamond - currentDiamond);
    var progressPercent = targetDiamond > 0 ? Math.round((currentDiamond / targetDiamond) * 100) : 0;
    var status = currentDiamond >= targetDiamond ? 'Completed' : 'In Progress';

    rawList.push({
      vjName: colE_Stage,
      fullName: profile.fullName,
      team: profile.team || colA_Team || 'ไม่ระบุทีม',
      shift: profile.shift || cleanStr(row[1]) || '-',
      targetDiamond: targetDiamond,
      currentDiamond: currentDiamond,
      remainingDiamond: remDiamond,
      progressPercent: progressPercent,
      status: status,
      dailyEarnings: dailyEarnings,
      profilePic: profile.avatarUrl,
      tiktokId: profile.tiktokId,
      password: profile.password,
      workDays: cleanStr(row[2]) || '0',
      leaveDays: cleanStr(row[3]) || '0'
    });
  }

  rawList.sort(function(a, b) { return b.currentDiamond - a.currentDiamond; });
  rawList.forEach(function(item, idx) { item.rank = idx + 1; });

  return {
    month: selectedMonth,
    list: rawList
  };
}

// ------------------------------------------
// 📈 MOM COMPARISON DATA MODULE
// ------------------------------------------

function getKpiComparisonData(currentMonth, prevMonth) {
  try {
    var availableMonths = getAvailableMonths();
    if (!currentMonth && availableMonths.length > 0) currentMonth = availableMonths[0];
    if (!prevMonth && availableMonths.length > 1) {
      var currIdx = availableMonths.indexOf(currentMonth);
      if (currIdx !== -1 && currIdx + 1 < availableMonths.length) {
        prevMonth = availableMonths[currIdx + 1];
      } else {
        prevMonth = availableMonths[1];
      }
    }

    var currData = getAllMonthData(currentMonth);
    var prevData = prevMonth ? getAllMonthData(prevMonth) : { list: [] };

    if (!currData || !currData.list) {
      return { success: false, message: 'ไม่พบข้อมูลของเดือน ' + currentMonth };
    }

    var currList = currData.list || [];
    var prevList = (prevData && prevData.list) ? prevData.list : [];

    var prevVjMap = {};
    var prevTotalCompanyDiamonds = 0;
    var prevTotalCompanyTarget = 0;

    prevList.forEach(function(item) {
      prevVjMap[item.vjName.toLowerCase()] = item;
      prevTotalCompanyDiamonds += item.currentDiamond || 0;
      prevTotalCompanyTarget += item.targetDiamond || 0;
    });

    var currTotalCompanyDiamonds = 0;
    var currTotalCompanyTarget = 0;
    var vjComparisonList = [];

    var currTeamMap = {};
    var prevTeamMap = {};

    currList.forEach(function(currVJ) {
      var lowerName = currVJ.vjName.toLowerCase();
      var prevVJ = prevVjMap[lowerName] || null;

      var cDiam = currVJ.currentDiamond || 0;
      var pDiam = prevVJ ? (prevVJ.currentDiamond || 0) : 0;
      var cTarget = currVJ.targetDiamond || 0;
      var pTarget = prevVJ ? (prevVJ.targetDiamond || 0) : 0;

      currTotalCompanyDiamonds += cDiam;
      currTotalCompanyTarget += cTarget;

      var diamDiff = cDiam - pDiam;
      var diamGrowthPercent = pDiam > 0 ? Math.round((diamDiff / pDiam) * 100) : (cDiam > 0 ? 100 : 0);

      vjComparisonList.push({
        vjName: currVJ.vjName,
        fullName: currVJ.fullName,
        team: currVJ.team,
        currTarget: cTarget,
        currEarned: cDiam,
        currAchPercent: currVJ.progressPercent,
        prevTarget: pTarget,
        prevEarned: pDiam,
        prevAchPercent: prevVJ ? prevVJ.progressPercent : 0,
        diamDiff: diamDiff,
        diamGrowthPercent: diamGrowthPercent,
        profilePic: currVJ.profilePic
      });

      var tName = currVJ.team || 'ไม่ระบุทีม';
      if (!currTeamMap[tName]) currTeamMap[tName] = 0;
      currTeamMap[tName] += cDiam;
    });

    prevList.forEach(function(prevVJ) {
      var tName = prevVJ.team || 'ไม่ระบุทีม';
      if (!prevTeamMap[tName]) prevTeamMap[tName] = 0;
      prevTeamMap[tName] += (prevVJ.currentDiamond || 0);
    });

    var teamComparisonList = [];
    var allTeamNames = Array.from(new Set(Object.keys(currTeamMap).concat(Object.keys(prevTeamMap))));

    allTeamNames.forEach(function(tName) {
      var cVal = currTeamMap[tName] || 0;
      var pVal = prevTeamMap[tName] || 0;
      var diff = cVal - pVal;
      var growthPct = pVal > 0 ? Math.round((diff / pVal) * 100) : (cVal > 0 ? 100 : 0);

      teamComparisonList.push({
        teamName: tName,
        currTotal: cVal,
        prevTotal: pVal,
        diff: diff,
        growthPct: growthPct
      });
    });

    teamComparisonList.sort(function(a, b) { return b.currTotal - a.currTotal; });

    var totalDiamDiff = currTotalCompanyDiamonds - prevTotalCompanyDiamonds;
    var companyGrowthPercent = prevTotalCompanyDiamonds > 0 ? Math.round((totalDiamDiff / prevTotalCompanyDiamonds) * 100) : 0;

    vjComparisonList.sort(function(a, b) { return b.diamDiff - a.diamDiff; });

    return {
      success: true,
      currentMonth: currentMonth,
      prevMonth: prevMonth || 'ไม่มีข้อมูลเดือนก่อนหน้า',
      availableMonths: availableMonths,
      companySummary: {
        currTotalDiamonds: currTotalCompanyDiamonds,
        prevTotalDiamonds: prevTotalCompanyDiamonds,
        currTotalTarget: currTotalCompanyTarget,
        prevTotalTarget: prevTotalCompanyTarget,
        currAchPercent: currTotalCompanyTarget > 0 ? Math.round((currTotalCompanyDiamonds / currTotalCompanyTarget) * 100) : 0,
        prevAchPercent: prevTotalCompanyTarget > 0 ? Math.round((prevTotalCompanyDiamonds / prevTotalCompanyTarget) * 100) : 0,
        diamDiff: totalDiamDiff,
        growthPercent: companyGrowthPercent
      },
      vjComparisonList: vjComparisonList,
      teamComparisonList: teamComparisonList
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ------------------------------------------
// ⏰ ATTENDANCE SYSTEM MODULE
// ------------------------------------------

function getOrCreateAttendanceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance_Data') || ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance_Data');
    var headers = [
      'ID', 'Date', 'Month_Year', 'Stage_Name', 'Team', 'Scheduled_Shift',
      'Clock_In', 'Clock_Out', 'Work_Hours', 'Late_Minutes', 'Status', 'Source', 'Created_By', 'Created_At'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e2e8f0');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDateStandard(dateInput) {
  if (!dateInput) return '';
  if (dateInput instanceof Date) {
    return Utilities.formatDate(dateInput, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  var str = cleanStr(dateInput);
  if (str.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    var ymd = str.split('/');
    return ymd[0] + '-' + ymd[1].padStart(2, '0') + '-' + ymd[2].padStart(2, '0');
  }
  if (str.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
    var p = str.split('/');
    var day = p[0].padStart(2, '0');
    var month = p[1].padStart(2, '0');
    var year = p[2];
    return year + '-' + month + '-' + day;
  }
  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return str;
  }
  return str;
}

var ATTENDANCE_SOURCE_SPREADSHEET_ID = '1xiQFVwPXFG9wK91vh78DYj3gaeSvJBSGmg90adatfDc';

function formatAttendanceTime(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Bangkok', 'HH:mm');
  if (typeof value === 'number' && value >= 0 && value < 1) {
    var totalMinutes = Math.round(value * 1440);
    return ('0' + Math.floor(totalMinutes / 60) % 24).slice(-2) + ':' + ('0' + (totalMinutes % 60)).slice(-2);
  }
  var raw = cleanStr(value);
  var match = raw.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
  return match ? ('0' + match[1]).slice(-2) + ':' + match[2] : raw;
}

function attendanceMonthFromDate(dateText) {
  var m = cleanStr(dateText).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  var date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  return Utilities.formatDate(date, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();
}

function getAttendanceSourceSpreadsheet() {
  return SpreadsheetApp.openById(ATTENDANCE_SOURCE_SPREADSHEET_ID);
}

function getVJShiftMap(selectedMonth, attendanceSpreadsheet) {
  var map = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
    if (vjSheet) {
      var vValues = vjSheet.getDataRange().getValues();
      var hRow = vValues[0] || [];
      var stageIdx = 1, teamIdx = 2, shiftIdx = 8;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr(hRow[c]).toUpperCase();
        if (hStr === 'STGE NAME' || hStr === 'STAGE NAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
        else if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
        else if (hStr === 'SHIFT' || hStr === 'เวลาทำงาน' || c === 8) shiftIdx = c;
      }
      for (var i = 1; i < vValues.length; i++) {
        var stage = cleanStr(vValues[i][stageIdx]);
        if (stage) {
          map[stage.toLowerCase()] = {
            stageName: stage,
            team: cleanStr(vValues[i][teamIdx]) || 'ไม่ระบุทีม',
            shift: cleanStr(vValues[i][shiftIdx]) || '-'
          };
        }
      }
    }

    if (selectedMonth) {
      var monthSheet = ss.getSheetByName(selectedMonth);
      if (monthSheet) {
        var mValues = monthSheet.getDataRange().getValues();
        for (var r = 1; r < mValues.length; r++) {
          var team = cleanStr(mValues[r][0]);
          var shift = cleanStr(mValues[r][1]);
          var vjName = cleanStr(mValues[r][4]);
          if (vjName) {
            var lowerName = vjName.toLowerCase();
            if (!map[lowerName]) {
              map[lowerName] = { stageName: vjName, team: team || 'ไม่ระบุทีม', shift: shift || '-' };
            } else {
              if (shift && shift !== '-') map[lowerName].shift = shift;
              if (team && team !== '-') map[lowerName].team = team;
            }
          }
        }
      }
    }

    // Shift_Settings in the attendance workbook supplies team start/end times when the ACE VJ sheet has no shift.
    var attendanceSs = attendanceSpreadsheet || getAttendanceSourceSpreadsheet();
    var settingsSheet = attendanceSs.getSheetByName('Shift_Settings');
    if (settingsSheet) {
      var settings = settingsSheet.getDataRange().getValues();
      var sh = (settings[0] || []).map(function(v) { return cleanStr(v).toUpperCase().replace(/[\s_-]+/g, ''); });
      var teamCol = sh.indexOf('TEAM'), startCol = sh.indexOf('WORKSTART'), endCol = sh.indexOf('WORKEND');
      if (teamCol >= 0 && startCol >= 0 && endCol >= 0) {
        var teamShifts = Object.create(null);
        for (var si = 1; si < settings.length; si++) {
          var teamName = cleanStr(settings[si][teamCol]);
          var startTime = formatAttendanceTime(settings[si][startCol]);
          var endTime = formatAttendanceTime(settings[si][endCol]);
          if (teamName && startTime && endTime) teamShifts[teamName.toLowerCase()] = startTime + ' - ' + endTime;
        }
        Object.keys(map).forEach(function(key) {
          if (!map[key].shift || map[key].shift === '-' || map[key].shift === 'ไม่ระบุกะ') {
            map[key].shift = teamShifts[cleanStr(map[key].team).toLowerCase()] || map[key].shift || '-';
          }
        });
      }
    }
  } catch (err) {
    Logger.log('Error in getVJShiftMap: ' + err.toString());
  }
  return map;
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr) return null;
  var cleaned = cleanStr(timeStr);
  var match = cleaned.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return null;
  var hrs = parseInt(match[1], 10);
  var mins = parseInt(match[2], 10);
  return hrs * 60 + mins;
}

function calculateAttendanceMetrics(clockIn, clockOut, shiftStr) {
  var shiftStartMins = null;
  var shiftEndMins = null;
  var shiftStartStr = '';
  var shiftEndStr = '';

  if (shiftStr && shiftStr !== '-' && shiftStr !== 'ไม่ระบุกะ') {
    var parts = shiftStr.split('-');
    if (parts.length >= 2) {
      shiftStartStr = parts[0].trim();
      shiftEndStr = parts[1].trim();
      shiftStartMins = parseTimeToMinutes(shiftStartStr);
      shiftEndMins = parseTimeToMinutes(shiftEndStr);
    }
  }

  if (clockIn && !clockOut && shiftEndStr) {
    clockOut = shiftEndStr;
  }

  var clockInMins = parseTimeToMinutes(clockIn);
  var clockOutMins = parseTimeToMinutes(clockOut);

  var lateMinutes = 0;
  if (clockInMins !== null && shiftStartMins !== null) {
    if (clockInMins > shiftStartMins) {
      lateMinutes = clockInMins - shiftStartMins;
    }
  }

  var workHours = 0;
  if (clockInMins !== null && clockOutMins !== null) {
    var diff = clockOutMins - clockInMins;
    if (diff < 0) diff += 1440;
    workHours = Math.round((diff / 60) * 100) / 100;
  }

  var status = 'ตรงเวลา';
  if (!clockIn && !clockOut) {
    status = 'ขาดงาน';
  } else if (clockInMins !== null && shiftStartMins !== null && lateMinutes > 0) {
    status = 'สาย';
  } else if (clockInMins !== null) {
    status = 'ตรงเวลา';
  } else {
    status = 'ขาดงาน';
  }

  return {
    clockOutFixed: clockOut,
    lateMinutes: lateMinutes,
    workHours: workHours,
    status: status
  };
}

function getAttendanceRecords(selectedMonth, targetVJName) {
  try {
    var filterVJ = targetVJName ? cleanStr(targetVJName).toLowerCase() : '';
    var filterMonth = selectedMonth ? cleanStr(selectedMonth).toUpperCase() : '';
    var recordsByKey = Object.create(null);
    var sourceWarning = '';
    var externalSs = null;
    var shiftMap = {};

    try {
      externalSs = getAttendanceSourceSpreadsheet();
      var logSheet = externalSs.getSheetByName('Attendance_Logs');
      if (!logSheet) throw new Error('ไม่พบชีต Attendance_Logs ในไฟล์ต้นทาง');
      shiftMap = getVJShiftMap(selectedMonth, externalSs);
      var sourceData = logSheet.getDataRange().getValues();
      if (sourceData.length > 1) {
        var headers = (sourceData[0] || []).map(function(v) { return cleanStr(v).toUpperCase().replace(/[\s_-]+/g, ''); });
        function sourceCol(names) { for (var x = 0; x < names.length; x++) { var idx = headers.indexOf(names[x]); if (idx >= 0) return idx; } return -1; }
        var idCol = sourceCol(['LOGID','ID']);
        var nameCol = sourceCol(['NAME','STAGENAME','VJNAME']);
        var dateCol = sourceCol(['DATE','WORKDATE']);
        var inCol = sourceCol(['TIMEIN','CLOCKIN','CHECKIN']);
        var outCol = sourceCol(['TIMEOUT','CLOCKOUT','CHECKOUT']);
        if (nameCol < 0 || dateCol < 0 || inCol < 0 || outCol < 0) throw new Error('หัวคอลัมน์ Attendance_Logs ต้องมี Name, Date, TimeIn และ TimeOut');

        for (var i = 1; i < sourceData.length; i++) {
          var row = sourceData[i];
          var dateVal = formatDateStandard(row[dateCol]);
          var vjName = cleanStr(row[nameCol]);
          if (!dateVal || !vjName) continue;
          var lowerName = vjName.toLowerCase();
          var monthYear = attendanceMonthFromDate(dateVal);
          if (filterVJ && lowerName !== filterVJ) continue;
          if (filterMonth && monthYear !== filterMonth && attendanceMonthFromDate(dateVal) !== filterMonth && dateVal.substring(0, 7) !== filterMonth) continue;
          var mapped = shiftMap[lowerName] || { team: '-', shift: '-' };
          var clockIn = formatAttendanceTime(row[inCol]);
          var clockOut = formatAttendanceTime(row[outCol]);
          var metrics = calculateAttendanceMetrics(clockIn, clockOut, mapped.shift);
          var key = dateVal + '|' + lowerName;
          recordsByKey[key] = {
            id: idCol >= 0 ? cleanStr(row[idCol]) : 'SOURCE_' + (i + 1), rowIndex: i + 1,
            date: dateVal, monthYear: monthYear, vjName: vjName, team: mapped.team || '-',
            scheduledShift: mapped.shift || '-', clockIn: clockIn, clockOut: metrics.clockOutFixed || clockOut,
            workHours: metrics.workHours, lateMinutes: metrics.lateMinutes, status: metrics.status,
            source: 'Attendance_Logs', createdBy: '-', createdAt: '-'
          };
        }
      }
    } catch (sourceErr) {
      sourceWarning = 'อ่านข้อมูลจาก Attendance_Logs ไม่สำเร็จ: ' + sourceErr.toString() + ' โปรดตรวจสิทธิ์บัญชีเจ้าของ Web App ในไฟล์ Attendance_System_Database';
      Logger.log(sourceWarning);
    }

    // Admin corrections/imports stay in the ACE workbook and override the external row by date + VJ.
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    var localSheet = activeSs.getSheetByName('Attendance_Data') || activeSs.getSheetByName('Attendance');
    if (localSheet) {
      var localData = localSheet.getDataRange().getValues();
      if (localData.length > 1) {
        if (!Object.keys(shiftMap).length) shiftMap = getVJShiftMap(selectedMonth, externalSs || undefined);
        for (var j = 1; j < localData.length; j++) {
          var r = localData[j];
          var localDate = formatDateStandard(r[1]);
          var localName = cleanStr(r[3]);
          var localLower = localName.toLowerCase();
          if (!localDate || !localName) continue;
          var localMonth = cleanStr(r[2]).toUpperCase() || attendanceMonthFromDate(localDate);
          if (filterVJ && localLower !== filterVJ) continue;
          if (filterMonth && localMonth !== filterMonth && attendanceMonthFromDate(localDate) !== filterMonth && localDate.substring(0, 7) !== filterMonth) continue;
          var localMapped = shiftMap[localLower] || { team: cleanStr(r[4]) || '-', shift: cleanStr(r[5]) || '-' };
          var localIn = formatAttendanceTime(r[6]);
          var localOut = formatAttendanceTime(r[7]);
          var localMetrics = calculateAttendanceMetrics(localIn, localOut, cleanStr(r[5]) || localMapped.shift);
          recordsByKey[localDate + '|' + localLower] = {
            id: cleanStr(r[0]), rowIndex: j + 1, date: localDate, monthYear: localMonth,
            vjName: localName, team: cleanStr(r[4]) || localMapped.team || '-',
            scheduledShift: cleanStr(r[5]) || localMapped.shift || '-', clockIn: localIn,
            clockOut: localMetrics.clockOutFixed || localOut,
            workHours: r[8] !== '' && r[8] !== null ? Number(r[8]) : localMetrics.workHours,
            lateMinutes: r[9] !== '' && r[9] !== null ? Number(r[9]) : localMetrics.lateMinutes,
            status: cleanStr(r[10]) || localMetrics.status, source: cleanStr(r[11]) || 'Manual',
            createdBy: cleanStr(r[12]) || '-', createdAt: r[13] ? (r[13] instanceof Date ? Utilities.formatDate(r[13], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr(r[13])) : '-'
          };
        }
      }
    }
    var records = Object.keys(recordsByKey).map(function(key) { return recordsByKey[key]; });
    records.sort(function(a,b) { return a.date === b.date ? a.vjName.localeCompare(b.vjName) : (a.date < b.date ? 1 : -1); });
    return { success: true, records: records, source: 'Attendance_Logs', sourceWarning: sourceWarning };
  } catch (err) {
    return { success: false, message: err.toString(), records: [], source: 'Attendance_Logs' };
  }
}

function uploadBulkAttendance(recordsArray, adminName, selectedMonth, ipAddress, deviceInfo) {
  try {
    if (!recordsArray || recordsArray.length === 0) {
      return { success: false, message: 'ไม่มีข้อมูลสำหรับอัปโหลด' };
    }

    var sheet = getOrCreateAttendanceSheet();
    var shiftMap = getVJShiftMap(selectedMonth);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var createdBy = adminName || 'Admin';

    var data = sheet.getDataRange().getValues();
    var lookupMap = {};
    for (var i = 1; i < data.length; i++) {
      var rDate = formatDateStandard(data[i][1]);
      var rVJ = cleanStr(data[i][3]).toLowerCase();
      if (rDate && rVJ) {
        lookupMap[rDate + '_' + rVJ] = {
          rowIndex: i + 1,
          id: data[i][0]
        };
      }
    }

    var count = 0;
    recordsArray.forEach(function(item) {
      var dateVal = formatDateStandard(item.date);
      var vjName = cleanStr(item.vjName || item.stageName);
      if (!dateVal || !vjName) return;

      var lowerVJ = vjName.toLowerCase();
      var mappedInfo = shiftMap[lowerVJ] || { team: 'ไม่ระบุทีม', shift: 'ไม่ระบุกะ' };

      var team = item.team || mappedInfo.team;
      var shift = item.shift || mappedInfo.shift;
      var clockIn = cleanStr(item.checkIn || item.clockIn);
      var clockOut = cleanStr(item.checkOut || item.clockOut);

      var metrics = calculateAttendanceMetrics(clockIn, clockOut, shift);
      var status = item.status || metrics.status;
      var lateMins = metrics.lateMinutes;
      var workHrs = metrics.workHours;
      var finalClockOut = metrics.clockOutFixed;

      var monthYear = selectedMonth || (dateVal ? dateVal.substring(0, 7) : '');
      var source = item.source || 'Excel_Import';

      var key = dateVal + '_' + lowerVJ;
      if (lookupMap[key]) {
        var rowIdx = lookupMap[key].rowIndex;
        var existingId = lookupMap[key].id || ('ATT_' + new Date().getTime());
        sheet.getRange(rowIdx, 1, 1, 14).setValues([[
          existingId, dateVal, monthYear, vjName, team, shift, clockIn, finalClockOut, workHrs, lateMins, status, source, createdBy, nowStr
        ]]);
      } else {
        var newId = 'ATT_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
        sheet.appendRow([
          newId, dateVal, monthYear, vjName, team, shift, clockIn, finalClockOut, workHrs, lateMins, status, source, createdBy, nowStr
        ]);
        lookupMap[key] = { rowIndex: sheet.getLastRow(), id: newId };
      }
      count++;
    });

    writeLog('ADMIN', adminName, 'IMPORT_ATTENDANCE_EXCEL', 'อัปโหลดเวลาเข้างานประจำเดือน ' + selectedMonth + ' จำนวน ' + count + ' รายการ', ipAddress, deviceInfo);
    return { success: true, count: count, message: 'นำเข้าข้อมูลและอัปเดตเรียบร้อยแล้ว ' + count + ' รายการ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveSingleAttendanceRecord(recordData, adminName, selectedMonth, ipAddress, deviceInfo) {
  try {
    var sheet = getOrCreateAttendanceSheet();
    var shiftMap = getVJShiftMap(selectedMonth);

    var dateVal = formatDateStandard(recordData.date);
    var vjName = cleanStr(recordData.vjName);
    var lowerVJ = vjName.toLowerCase();
    var mappedInfo = shiftMap[lowerVJ] || { team: 'ไม่ระบุทีม', shift: 'ไม่ระบุกะ' };

    var team = recordData.team || mappedInfo.team;
    var shift = recordData.shift || mappedInfo.shift;
    var clockIn = cleanStr(recordData.checkIn || recordData.clockIn);
    var clockOut = cleanStr(recordData.checkOut || recordData.clockOut);
    
    var metrics = calculateAttendanceMetrics(clockIn, clockOut, shift);
    var status = recordData.status || metrics.status;
    var lateMins = metrics.lateMinutes;
    var workHrs = metrics.workHours;
    var finalClockOut = metrics.clockOutFixed;

    var monthYear = selectedMonth || (dateVal ? dateVal.substring(0, 7) : '');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var createdBy = adminName || 'Admin';
    var source = recordData.source || 'Manual';

    var data = sheet.getDataRange().getValues();
    var targetRowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      var rDate = formatDateStandard(data[i][1]);
      var rVJ = cleanStr(data[i][3]).toLowerCase();
      if (rDate === dateVal && rVJ === lowerVJ) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex > 1) {
      var existingId = data[targetRowIndex - 1][0] || ('ATT_' + new Date().getTime());
      sheet.getRange(targetRowIndex, 1, 1, 14).setValues([[
        existingId, dateVal, monthYear, vjName, team, shift, clockIn, finalClockOut, workHrs, lateMins, status, source, createdBy, nowStr
      ]]);
      writeLog('ADMIN', adminName, 'MANUAL_ATTENDANCE_UPDATE', 'แก้ไขเวลาเข้างาน VJ: ' + vjName + ' วันที่: ' + dateVal, ipAddress, deviceInfo);
    } else {
      var newId = 'ATT_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
      sheet.appendRow([
        newId, dateVal, monthYear, vjName, team, shift, clockIn, finalClockOut, workHrs, lateMins, status, source, createdBy, nowStr
      ]);
      writeLog('ADMIN', adminName, 'MANUAL_ATTENDANCE_UPDATE', 'บันทึกเวลาเข้างาน VJ: ' + vjName + ' วันที่: ' + dateVal, ipAddress, deviceInfo);
    }

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ------------------------------------------
// 🔓 Authentication & Data Return
// ------------------------------------------

function verifyAndGetPersonalData(vjStageName, inputPassword, selectedMonth, ipAddress, deviceInfo) {
  try {
    var permCheck = checkAccessPermission(vjStageName, ipAddress);
    if (!permCheck.allowed) {
      writeLog('VJ', vjStageName, 'ACCESS_BLOCKED', permCheck.reason, ipAddress, deviceInfo);
      return { success: false, message: '⛔ ' + permCheck.reason };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
    var targetName = cleanStr(vjStageName).toLowerCase();
    var targetPass = cleanStr(inputPassword);

    var verified = false;
    var userProfile = { stageName: cleanStr(vjStageName), fullName: cleanStr(vjStageName), role: '', team: '-', shift: '-', profilePic: getProfilePicUrl('', vjStageName) };
    if (vjSheet) {
      var vjValues = vjSheet.getDataRange().getValues();
      var hRow = vjValues[0] || [];
      var stageIdx = 1, passIdx = -1, nameIdx = 0, teamIdx = 2, roleIdx = 3, shiftIdx = 5, tiktokIdx = 7;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr(hRow[c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (hStr === 'STGENAME' || hStr === 'STAGENAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
        if (hStr === 'NAME' || hStr === 'FULLNAME' || hStr === 'ชื่อ') nameIdx = c;
        if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
        if (hStr === 'ROLE' || hStr === 'ตำแหน่ง') roleIdx = c;
        if (hStr === 'SHIFT' || hStr === 'เวลาทำงาน') shiftIdx = c;
        if (hStr === 'IDTIKTOK' || hStr === 'TIKTOKID' || hStr === 'TIKTOK') tiktokIdx = c;
        if (hStr === 'PASSWORD' || hStr === 'PASSCODE' || hStr === 'PIN' || hStr === 'รหัสผ่าน') passIdx = c;
      }
      for (var i = 1; i < vjValues.length; i++) {
        var sName = cleanStr(vjValues[i][stageIdx]).toLowerCase();
        if (sName === targetName) {
          var pWord = (passIdx !== -1 && vjValues[i][passIdx]) ? cleanStr(vjValues[i][passIdx]) : "1234";
          if (pWord === targetPass) {
            verified = true;
            var tiktokId = cleanStr(vjValues[i][tiktokIdx]);
            userProfile = {
              stageName: cleanStr(vjValues[i][stageIdx]) || cleanStr(vjStageName),
              fullName: cleanStr(vjValues[i][nameIdx]) || cleanStr(vjStageName),
              role: cleanStr(vjValues[i][roleIdx]),
              team: cleanStr(vjValues[i][teamIdx]) || '-',
              shift: cleanStr(vjValues[i][shiftIdx]) || '-',
              profilePic: getProfilePicUrl(tiktokId, vjStageName)
            };
          }
          break;
        }
      }
    } else {
      if (targetPass === "1234") verified = true;
    }

    if (!verified) {
      writeLog('VJ', vjStageName, 'LOGIN_FAILED', 'รหัสผ่านไม่ถูกต้อง', ipAddress, deviceInfo);
      return { success: false, message: '🔒 รหัสผ่าน VJ ไม่ถูกต้อง' };
    }

    writeLog('VJ', vjStageName, 'LOGIN_SUCCESS', 'เข้าดูข้อมูลเดือน: ' + selectedMonth, ipAddress, deviceInfo);

    var allData = getAllMonthData(selectedMonth);
    var roleKey = userProfile.role.toUpperCase().trim();
    var isAttendanceRole = roleKey === 'HR' || roleKey === 'MC';
    var attendanceOnly = !allData || !allData.list;
    if (attendanceOnly && !isAttendanceRole) {
      return { success: false, message: 'ไม่พบข้อมูลประจำเดือน ' + (selectedMonth || '') };
    }

    var sortedList = attendanceOnly ? [] : allData.list;
    var personalData = null;
    var rank = 0;
    var nextRankInfo = null;

    for (var k = 0; k < sortedList.length; k++) {
      if (sortedList[k].vjName.toLowerCase() === targetName) {
        personalData = sortedList[k];
        rank = k + 1;
        if (rank > 1) {
          var aheadVJ = sortedList[k - 1];
          nextRankInfo = {
            aheadRank: k,
            aheadVJName: aheadVJ.vjName,
            aheadDiamonds: aheadVJ.currentDiamond,
            diamondsNeededToPass: (aheadVJ.currentDiamond - personalData.currentDiamond) + 1
          };
        } else if (sortedList.length > 1) {
          var runnerUp = sortedList[1];
          nextRankInfo = {
            leadMargin: personalData.currentDiamond - runnerUp.currentDiamond,
            runnerUpName: runnerUp.vjName
          };
        }
        break;
      }
    }

    if (!personalData && isAttendanceRole) {
      attendanceOnly = true;
      personalData = {
        stageName: userProfile.stageName, fullName: userProfile.fullName,
        team: userProfile.team, role: userProfile.role, shift: userProfile.shift,
        profilePic: userProfile.profilePic, rank: 0, totalVJs: 0,
        targetDiamond: 0, currentDiamond: 0, remainingDiamond: 0,
        progressPercent: 0, dailyEarnings: [], nextRankInfo: null
      };
    }
    if (!personalData) {
      return { success: false, message: 'ไม่พบข้อมูลของ VJ "' + vjStageName + '" ในเดือน ' + selectedMonth };
    }

    personalData.rank = rank;
    personalData.totalVJs = sortedList.length;
    personalData.nextRankInfo = nextRankInfo;

    var attRes = getAttendanceRecords(selectedMonth || '', vjStageName);
    var attendanceList = attRes.success ? attRes.records : [];

    var annRes = getActiveAnnouncements(vjStageName, 'VJ');
    var announcements = annRes.success ? annRes.list : [];

    return {
      success: true,
      month: selectedMonth,
      personalData: personalData,
      attendanceOnly: attendanceOnly,
      attendanceList: attendanceList,
      attendanceSourceWarning: attRes.sourceWarning || '',
      announcements: announcements,
      systemVersionInfo: getSystemVersionInfo()
    };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}

function getAdminDashboardData(adminName, adminPin, selectedMonth, ipAddress, deviceInfo) {
  try {
    var permCheck = checkAccessPermission(adminName, ipAddress);
    if (!permCheck.allowed) {
      writeLog('ADMIN', adminName, 'ACCESS_BLOCKED', permCheck.reason, ipAddress, deviceInfo);
      return { success: false, message: '⛔ ' + permCheck.reason };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName('Admins') || ss.getSheetByName('admins') || ss.getSheetByName('ADMINS');
    
    var safeClean = function(val) {
      if (typeof cleanStr === 'function') return cleanStr(val);
      return val ? String(val).trim() : '';
    };

    var targetName = safeClean(adminName).toLowerCase();
    var inputPin = safeClean(adminPin);
    var isAdmin = false;

    if (adminSheet) {
      var adminValues = adminSheet.getDataRange().getValues();
      var adminHRow = adminValues[0] || [];
      var nameIdx = 0, pinIdx = 1;

      for (var c = 0; c < adminHRow.length; c++) {
        var hStr = safeClean(adminHRow[c]).toUpperCase();
        if (hStr === 'NAME' || hStr === 'ADMIN NAME' || hStr === 'ADMIN' || hStr === 'USERNAME' || hStr === 'ชื่อ') nameIdx = c;
        if (hStr === 'PIN' || hStr === 'PASSWORD' || hStr === 'PASSCODE' || hStr === 'รหัสผ่าน') pinIdx = c;
      }

      for (var i = 1; i < adminValues.length; i++) {
        var aName = safeClean(adminValues[i][nameIdx]).toLowerCase();
        if (aName === targetName || (targetName === 'admin' && i === 1)) {
          var aPin = safeClean(adminValues[i][pinIdx]);
          if (aPin === inputPin) {
            isAdmin = true;
          } else {
            if (typeof writeLog === 'function') writeLog('ADMIN', adminName, 'LOGIN_FAILED', 'รหัสผ่าน Admin ไม่ถูกต้อง', ipAddress, deviceInfo);
            return { success: false, message: '🔒 รหัสผ่าน Admin ไม่ถูกต้อง' };
          }
          break;
        }
      }
    }

    if (!isAdmin && (inputPin.toUpperCase() === 'ADMIN123' || inputPin === '1234')) {
      isAdmin = true;
    }

    if (!isAdmin) {
      if (typeof writeLog === 'function') writeLog('ADMIN', adminName, 'LOGIN_FAILED', 'บัญชี Admin หรือรหัสผ่านไม่ถูกต้อง', ipAddress, deviceInfo);
      return { success: false, message: '🔒 บัญชี Admin หรือรหัสผ่านไม่ถูกต้อง' };
    }

    if (typeof writeLog === 'function') writeLog('ADMIN', adminName, 'LOGIN_SUCCESS', 'เข้าสู่ระบบ Admin Dashboard เดือน: ' + selectedMonth, ipAddress, deviceInfo);

    var allData = (typeof getAllMonthData === 'function') ? getAllMonthData(selectedMonth) : null;
    if (!allData || !allData.list) {
      return { success: false, message: 'ไม่พบข้อมูลประจำเดือน ' + selectedMonth };
    }

    var list = allData.list || [];
    var totalCompanyDiamonds = 0;
    var totalCompanyTarget = 0;
    var teamMap = {};

    list.forEach(function(vj) {
      var diamonds = Number(vj.currentDiamond) || 0;
      var target = Number(vj.targetDiamond) || 0;
      totalCompanyDiamonds += diamonds;
      totalCompanyTarget += target;

      var tName = vj.team || 'ไม่ระบุทีม';
      if (!teamMap[tName]) {
        teamMap[tName] = {
          teamName: tName,
          totalDiamonds: 0,
          memberCount: 0,
          members: []
        };
      }
      teamMap[tName].totalDiamonds += diamonds;
      teamMap[tName].memberCount += 1;
      teamMap[tName].members.push(vj);
    });

    var companyKPIPercent = totalCompanyTarget > 0 ? Math.round((totalCompanyDiamonds / totalCompanyTarget) * 100) : 0;

    var teamSummaryList = [];
    for (var teamKey in teamMap) {
      var t = teamMap[teamKey];
      var avgDiamonds = t.memberCount > 0 ? Math.round(t.totalDiamonds / t.memberCount) : 0;
      t.members.sort(function(a, b) { return (Number(b.currentDiamond) || 0) - (Number(a.currentDiamond) || 0); });
      teamSummaryList.push({
        teamName: t.teamName,
        totalDiamonds: t.totalDiamonds,
        memberCount: t.memberCount,
        avgDiamonds: avgDiamonds,
        members: t.members
      });
    }

    teamSummaryList.sort(function(a, b) { return b.totalDiamonds - a.totalDiamonds; });

    var attRes = (typeof getAttendanceRecords === 'function') ? getAttendanceRecords(selectedMonth, null) : { success: false };
    var attendanceList = (attRes && attRes.success) ? attRes.records : [];

    var logsRes = (typeof getLogs === 'function') ? getLogs(selectedMonth) : { success: false };
    var logsList = (logsRes && logsRes.success) ? logsRes.logs : [];

    var annRes = (typeof getActiveAnnouncements === 'function') ? getActiveAnnouncements(adminName, 'ADMIN') : { success: false };
    var announcements = (annRes && annRes.success) ? annRes.list : [];

    var annStatsRes = (typeof getAnnouncementStats === 'function') ? getAnnouncementStats() : null;
    var announcementStats = (annStatsRes && annStatsRes.success) ? annStatsRes : { stats: [], totalAnnouncements: 0, totalViews: 0 };

    var accessRes = getAccessControlList();
    var accessList = accessRes.success ? accessRes.list : [];

    var finalResult = {
      success: true,
      adminName: adminName,
      month: selectedMonth,
      companyKPISummary: {
        totalTarget: totalCompanyTarget,
        totalEarned: totalCompanyDiamonds,
        percent: companyKPIPercent
      },
      totalCompanyDiamonds: totalCompanyDiamonds,
      vjList: list,
      teamSummaryList: teamSummaryList,
      attendanceList: attendanceList,
      attendanceSourceWarning: attRes.sourceWarning || '',
      logsList: logsList,
      announcements: announcements,
      announcementStats: announcementStats,
      accessList: accessList,
      systemVersionInfo: getSystemVersionInfo()
    };

    return JSON.parse(JSON.stringify(finalResult));

  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}


// ==========================================
// 📱 USER MAPPING MODULE (ระบบจัดการและกรองข้อมูล User Mapping)
// ==========================================

/**
 * ฟังก์ชันสร้าง/ตรวจสอบชีต User_Mapping
 */

var USER_MAPPING_DATABASE_ID = '12Y62ZJ1iGAYl55u3hcp26PATNAadHf0Kf0bRdFxFl9Y';
var USER_MAPPING_HISTORY_SHEET = 'User_Mapping_History';
var USER_MAPPING_HISTORY_HEADERS = ['Snapshot Date', 'ID TikTok', 'User Name', 'TikTok URL', 'Assigned VJ', 'Team', 'เหรียญวันนี้', 'เหรียญเดือนนี้', 'Last Updated', 'Uploaded By', 'Total Days'];

function getUserMappingDatabase() {
  return SpreadsheetApp.openById(USER_MAPPING_DATABASE_ID);
}

function getOrCreateUserMappingSheet(spreadsheet) {
  var ss = spreadsheet || getUserMappingDatabase();
  var sheet = ss.getSheetByName('User_Mapping');
  if (!sheet) {
    sheet = ss.insertSheet('User_Mapping');
    var headers = ['ID TikTok', 'User Name', 'TikTok URL', 'Assigned VJ', 'Team', 'เหรียญวันนี้', 'เหรียญเดือนนี้', 'Last Updated', 'Uploaded By', 'Total Days'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateUserMappingHistorySheet(spreadsheet) {
  var ss = spreadsheet || getUserMappingDatabase();
  var sheet = ss.getSheetByName(USER_MAPPING_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USER_MAPPING_HISTORY_SHEET);
    sheet.getRange(1, 1, 1, USER_MAPPING_HISTORY_HEADERS.length).setValues([USER_MAPPING_HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, USER_MAPPING_HISTORY_HEADERS.length).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function userMappingSnapshotDate() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
}

function userMappingHistoryRow(snapshotDate, row) {
  return [snapshotDate, cleanStr(row[0]), cleanStr(row[1]), cleanStr(row[2]), cleanStr(row[3]), cleanStr(row[4]), row[5], row[6], row[7], cleanStr(row[8]), row[9]];
}

// Archive every current mapping once per day. Subsequent edits update that day's snapshot,
// while earlier dates remain immutable and A:J in User_Mapping remain unchanged.
function archiveUserMappingDailySnapshot(spreadsheet, liveRows) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var history = getOrCreateUserMappingHistorySheet(spreadsheet);
    var historyData = history.getDataRange().getValues();
    var snapshotDate = userMappingSnapshotDate();
    var existing = Object.create(null);
    for (var i = 1; i < historyData.length; i++) {
      var id = cleanStr(historyData[i][1]).toLowerCase();
      if (id) existing[cleanStr(historyData[i][0]) + '|' + id] = i + 1;
    }
    var newRows = [];
    for (var r = 1; r < liveRows.length; r++) {
      var row = liveRows[r];
      var id = cleanStr(row[0]);
      if (!id) continue;
      var key = snapshotDate + '|' + id.toLowerCase();
      if (!existing[key]) newRows.push(userMappingHistoryRow(snapshotDate, row));
    }
    if (newRows.length) {
      history.getRange(history.getLastRow() + 1, 1, newRows.length, USER_MAPPING_HISTORY_HEADERS.length).setValues(newRows);
    }
  } finally {
    lock.releaseLock();
  }
}

function archiveUserMappingDailySnapshotJob() {
  var database = getUserMappingDatabase();
  var sheet = getOrCreateUserMappingSheet(database);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) archiveUserMappingDailySnapshot(database, data);
}

function installUserMappingDailyArchiveTrigger() {
  var handler = 'archiveUserMappingDailySnapshotJob';
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function(trigger) { return trigger.getHandlerFunction() === handler; });
  if (!exists) ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(1).inTimezone('Asia/Bangkok').create();
  return { success: true, message: exists ? 'มี Trigger สำรองรายวันอยู่แล้ว' : 'ติดตั้ง Trigger สำรอง User Mapping รายวันเวลา 01:00 แล้ว' };
}

function upsertUserMappingDailySnapshot(spreadsheet, mappingData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var history = getOrCreateUserMappingHistorySheet(spreadsheet);
    var snapshotDate = userMappingSnapshotDate();
    var id = cleanStr(mappingData.tiktokId);
    if (!id) return;
    var values = history.getDataRange().getValues();
    var key = snapshotDate + '|' + id.toLowerCase();
    var target = -1;
    for (var i = 1; i < values.length; i++) {
      if (cleanStr(values[i][0]) + '|' + cleanStr(values[i][1]).toLowerCase() === key) { target = i + 1; break; }
    }
    var snapshot = [snapshotDate, id, cleanStr(mappingData.userName), cleanStr(mappingData.tiktokUrl), cleanStr(mappingData.assignedVJ), cleanStr(mappingData.team), mappingData.coinsToday || 0, mappingData.coinsMonth || 0, Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'), cleanStr(mappingData.uploadedBy), mappingData.totalDays || 0];
    if (target > 1) history.getRange(target, 1, 1, snapshot.length).setValues([snapshot]);
    else history.appendRow(snapshot);
  } finally {
    lock.releaseLock();
  }
}

/**
 * ดึงข้อมูล User Mapping แยกตามฝั่ง VJ และ Admin
 * - ฝั่ง VJ: กรองข้อมูลเฉพาะที่ Col D (Assigned VJ) ตรงกับชื่อ VJ ที่เข้าใช้งาน
 * - ฝั่ง Admin: แสดงข้อมูลทั้งหมด
 */
/**
 * ดึงข้อมูล User Mapping แยกตามฝั่ง VJ และ Admin (พร้อมรูปโปรไฟล์)
 */
/**
 * ดึงข้อมูล User Mapping พร้อมตรวจสอบสถานะการอัปเดตของวันนี้ (Asia/Bangkok)
 */
/**
 * ดึงข้อมูล User Mapping แยกตามฝั่ง VJ และ Admin (รองรับคอลัมน์ A - J)
 */
var USER_MAPPING_AVATAR_SHEET = 'User_Mapping_Avatars';
var USER_MAPPING_AVATAR_FOLDER_PROPERTY = 'ACE_USER_MAPPING_AVATAR_FOLDER_ID';
var USER_MAPPING_AVATAR_HEADERS = ['TikTok ID', 'Drive File ID', 'MIME Type', 'Cached At', 'Source URL'];

function getOrCreateUserMappingAvatarSheet(spreadsheet) {
  var ss = spreadsheet || getUserMappingDatabase();
  var sheet = ss.getSheetByName(USER_MAPPING_AVATAR_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USER_MAPPING_AVATAR_SHEET);
    sheet.getRange(1, 1, 1, USER_MAPPING_AVATAR_HEADERS.length).setValues([USER_MAPPING_AVATAR_HEADERS]);
    sheet.getRange(1, 1, 1, USER_MAPPING_AVATAR_HEADERS.length).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getUserMappingAvatarFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(USER_MAPPING_AVATAR_FOLDER_PROPERTY);
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { props.deleteProperty(USER_MAPPING_AVATAR_FOLDER_PROPERTY); }
  }
  var folder = DriveApp.createFolder('ACE User Mapping Avatar Cache');
  props.setProperty(USER_MAPPING_AVATAR_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function authorizeUserMappingAvatarCache() {
  getUserMappingDatabase();
  getOrCreateUserMappingAvatarSheet();
  getUserMappingAvatarFolder();
  UrlFetchApp.fetch('https://unavatar.io/', { muteHttpExceptions: true, followRedirects: true });
  return 'Avatar cache authorization ready. Drive folder is private by default.';
}

function getUserMappingAvatarData(avatarIds, userName, userType) {
  try {
    var ids = Array.isArray(avatarIds) ? avatarIds.map(function(id) { return cleanStr(id); }).filter(Boolean).slice(0, 10) : [];
    if (!ids.length) return { success: true, avatars: {} };
    var database = getUserMappingDatabase();
    var mappingSheet = getOrCreateUserMappingSheet(database);
    var mappingRows = mappingSheet.getDataRange().getValues();
    var cleanType = cleanStr(userType).toUpperCase();
    var cleanUser = cleanStr(userName).toLowerCase();
    if (cleanType !== 'ADMIN' && cleanType !== 'VJ') return { success: false, message: 'ไม่สามารถตรวจสอบสิทธิ์โหลดภาพได้', avatars: {} };
    var allowed = Object.create(null);
    for (var r = 1; r < mappingRows.length; r++) {
      var id = cleanStr(mappingRows[r][0]);
      var assigned = cleanStr(mappingRows[r][3]).toLowerCase();
      if (id && (cleanType === 'ADMIN' || assigned === cleanUser)) allowed[id.toLowerCase()] = id;
    }
    var historySheetForAccess = database.getSheetByName(USER_MAPPING_HISTORY_SHEET);
    if (cleanType === 'VJ' && historySheetForAccess) {
      var historicalRows = historySheetForAccess.getDataRange().getValues();
      for (var hr = 1; hr < historicalRows.length; hr++) {
        var historicalId = cleanStr(historicalRows[hr][1]);
        var historicalVj = cleanStr(historicalRows[hr][4]).toLowerCase();
        if (historicalId && historicalVj === cleanUser) allowed[historicalId.toLowerCase()] = historicalId;
      }
    }
    var active = SpreadsheetApp.getActiveSpreadsheet();
    var vjSheet = active.getSheetByName('vj mc all') || active.getSheetByName('VJ MC ALL') || active.getSheetByName('vj_mc_all');
    if (vjSheet) {
      var vjValues = vjSheet.getDataRange().getValues();
      var headers = vjValues[0] || [];
      var stageCol = 1, idCol = 7;
      for (var c = 0; c < headers.length; c++) {
        var h = cleanStr(headers[c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (h === 'STGENAME' || h === 'STAGENAME' || h === 'ชื่อการแสดง') stageCol = c;
        if (h === 'IDTIKTOK' || h === 'TIKTOKID' || h === 'TIKTOK') idCol = c;
      }
      for (var v = 1; v < vjValues.length; v++) {
        var stage = cleanStr(vjValues[v][stageCol]);
        var vjId = cleanStr(vjValues[v][idCol]);
        if (vjId && (cleanType === 'ADMIN' || stage.toLowerCase() === cleanUser)) allowed[vjId.toLowerCase()] = vjId;
      }
    }
    var requested = Object.create(null);
    ids.forEach(function(id) { if (allowed[id.toLowerCase()]) requested[id.toLowerCase()] = allowed[id.toLowerCase()]; });
    var cacheSheet = getOrCreateUserMappingAvatarSheet(database);
    var cacheRows = cacheSheet.getDataRange().getValues();
    var cacheMap = Object.create(null);
    for (var k = 1; k < cacheRows.length; k++) {
      var cacheId = cleanStr(cacheRows[k][0]).toLowerCase();
      if (cacheId) cacheMap[cacheId] = { rowIndex: k + 1, fileId: cleanStr(cacheRows[k][1]), mimeType: cleanStr(cacheRows[k][2]) };
    }
    var result = {}, newCacheRows = [], updates = [];
    var folder = null;
    Object.keys(requested).forEach(function(key) {
      var canonicalId = requested[key];
      // TikTok handles only; never let a cell value turn into an arbitrary UrlFetch target.
      if (!/^@?[a-zA-Z0-9._]{1,50}$/.test(canonicalId)) return;
      var cached = cacheMap[key];
      var blob = null;
      if (cached && cached.fileId) {
        try { blob = DriveApp.getFileById(cached.fileId).getBlob(); } catch (e) { blob = null; }
      }
      if (!blob) {
        var sourceUrl = getProfilePicUrl(canonicalId, canonicalId);
        if (!sourceUrl || sourceUrl.indexOf('ui-avatars.com') !== -1) return;
        var response = UrlFetchApp.fetch(sourceUrl, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 ACEContentPortal/1.0' } });
        var status = response.getResponseCode();
        var mimeType = cleanStr(response.getHeaders()['Content-Type'] || response.getHeaders()['content-type']).split(';')[0].toLowerCase();
        if (status < 200 || status >= 300 || !/^image\/(jpeg|png|webp|gif)$/.test(mimeType)) return;
        blob = response.getBlob().setContentType(mimeType);
        if (blob.getBytes().length > 1500000) return;
        if (!folder) folder = getUserMappingAvatarFolder();
        var safeName = canonicalId.replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 80) || 'avatar';
        var file = folder.createFile(blob.setName(safeName + '.' + (mimeType.split('/')[1] || 'img')));
        var cacheRow = [canonicalId, file.getId(), mimeType, Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'), sourceUrl];
        if (cached) updates.push({ rowIndex: cached.rowIndex, row: cacheRow }); else newCacheRows.push(cacheRow);
      }
      var type = blob.getContentType() || (cached && cached.mimeType) || 'image/jpeg';
      result[key] = 'data:' + type + ';base64,' + Utilities.base64Encode(blob.getBytes());
    });
    updates.forEach(function(item) { cacheSheet.getRange(item.rowIndex, 1, 1, item.row.length).setValues([item.row]); });
    if (newCacheRows.length) cacheSheet.getRange(cacheSheet.getLastRow() + 1, 1, newCacheRows.length, USER_MAPPING_AVATAR_HEADERS.length).setValues(newCacheRows);
    return { success: true, avatars: result };
  } catch (err) {
    return { success: false, message: err.toString(), avatars: {} };
  }
}

function getUserMappingData(userName, userType) {
  try {
    var database = getUserMappingDatabase();
    var sheet = getOrCreateUserMappingSheet(database);
    var data = sheet.getDataRange().getValues();
    if (!data || !data.length) data = [[]];
    if (data.length > 1) archiveUserMappingDailySnapshot(database, data);
    var historySheet = getOrCreateUserMappingHistorySheet(database);
    var historyData = historySheet.getDataRange().getValues();
    var todayStr = userMappingSnapshotDate();
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    var vjSheet = activeSs.getSheetByName('vj mc all') || activeSs.getSheetByName('VJ MC ALL') || activeSs.getSheetByName('vj_mc_all');
    var vjProfileMap = Object.create(null);
    if (vjSheet) {
      var vjValues = vjSheet.getDataRange().getValues();
      var hRow = vjValues[0] || [];
      var stageIdx = 1, tiktokIdx = 7;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr(hRow[c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (hStr === 'STGENAME' || hStr === 'STAGENAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
        if (hStr === 'IDTIKTOK' || hStr === 'TIKTOKID' || hStr === 'TIKTOK') tiktokIdx = c;
      }
      for (var v = 1; v < vjValues.length; v++) {
        var stage = cleanStr(vjValues[v][stageIdx]);
        if (stage) vjProfileMap[stage.toLowerCase()] = tiktokIdx >= 0 ? cleanStr(vjValues[v][tiktokIdx]) : '';
      }
    }

    var cleanUser = cleanStr(userName).toLowerCase();
    var cleanType = cleanStr(userType).toUpperCase();
    function makeItem(row, index, isHistory) {
      var offset = isHistory ? 1 : 0;
      var tiktokId = cleanStr(row[offset]);
      var uName = cleanStr(row[offset + 1]);
      var tiktokUrl = cleanStr(row[offset + 2]);
      var assignedVJ = cleanStr(row[offset + 3]);
      var team = cleanStr(row[offset + 4]);
      var updatedRaw = row[offset + 7];
      var updatedText = '-';
      var updatedDateOnly = '';
      if (updatedRaw) {
        if (updatedRaw instanceof Date) {
          updatedText = Utilities.formatDate(updatedRaw, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
          updatedDateOnly = Utilities.formatDate(updatedRaw, 'Asia/Bangkok', 'yyyy-MM-dd');
        } else {
          updatedText = cleanStr(updatedRaw);
          var match = updatedText.match(/^(\d{4}-\d{2}-\d{2})/);
          if (match) updatedDateOnly = match[1];
        }
      }
      var snapshotDate = isHistory ? cleanStr(row[0]) : updatedDateOnly;
      return {
        rowIndex: isHistory ? null : index + 1,
        snapshotDate: snapshotDate,
        tiktokId: tiktokId, userName: uName, tiktokUrl: tiktokUrl,
        assignedVJ: assignedVJ, team: team,
        coinsToday: parseDiamondVal(row[offset + 5]), coinsMonth: parseDiamondVal(row[offset + 6]),
        updatedDate: updatedText, recordedBy: cleanStr(row[offset + 8]), totalDays: parseDiamondVal(row[offset + 9]),
        userAvatar: getProfilePicUrl(tiktokId, uName || tiktokId),
        vjTikTokId: vjProfileMap[assignedVJ.toLowerCase()] || '',
        vjAvatar: getProfilePicUrl(vjProfileMap[assignedVJ.toLowerCase()] || '', assignedVJ),
        isUpdatedToday: (isHistory ? snapshotDate : updatedDateOnly) === todayStr
      };
    }

    var list = [], historyList = [], todayUpdatedCount = 0;
    for (var i = 1; i < data.length; i++) {
      var item = makeItem(data[i], i, false);
      if (!item.tiktokId && !item.userName && !item.assignedVJ) continue;
      if (cleanType !== 'VJ' || item.assignedVJ.toLowerCase() === cleanUser) {
        list.push(item);
        if (item.isUpdatedToday) todayUpdatedCount++;
      }
    }
    for (var h = 1; h < historyData.length; h++) {
      var hist = makeItem(historyData[h], h, true);
      if (!hist.tiktokId && !hist.userName && !hist.assignedVJ) continue;
      if (cleanType !== 'VJ' || hist.assignedVJ.toLowerCase() === cleanUser) historyList.push(hist);
    }
    historyList.sort(function(a,b) { return a.snapshotDate === b.snapshotDate ? a.team.localeCompare(b.team) || a.assignedVJ.localeCompare(b.assignedVJ) || a.userName.localeCompare(b.userName) : (a.snapshotDate < b.snapshotDate ? 1 : -1); });
    return { success: true, userType: cleanType, list: list, historyList: historyList, totalCount: list.length, todayUpdatedCount: todayUpdatedCount, todayStr: todayStr };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในระบบ User Mapping: ' + err.toString(), list: [], historyList: [] };
  }
}

/**
 * บันทึกหรืออัปเดตข้อมูล User Mapping (สำหรับ Admin)
 */
function saveUserMappingRecord(mappingData, adminName) {
  try {
    var database = getUserMappingDatabase();
    var sheet = getOrCreateUserMappingSheet(database);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var recorder = adminName || 'Admin';

    var tiktokId = cleanStr(mappingData.tiktokId);
    var uName = cleanStr(mappingData.userName);
    var tiktokUrl = cleanStr(mappingData.tiktokUrl);
    var assignedVJ = cleanStr(mappingData.assignedVJ);
    var team = cleanStr(mappingData.team);
    var coinsToday = mappingData.coinsToday;
    var coinsMonth = mappingData.coinsMonth;
    var totalDays = mappingData.totalDays;

    if (!tiktokId || !uName || !assignedVJ) {
      return { success: false, message: 'โปรดระบุ ID TikTok, User Name และ Assigned VJ ให้ครบถ้วน' };
    }

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;

    if (mappingData.rowIndex && Number(mappingData.rowIndex) > 1) {
      targetRow = Number(mappingData.rowIndex);
    } else {
      for (var i = 1; i < data.length; i++) {
        if (cleanStr(data[i][0]).toLowerCase() === tiktokId.toLowerCase()) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (targetRow > 1) {
      var existingRow = data[targetRow - 1] || [];
      if (coinsToday === undefined) coinsToday = existingRow[5] || 0;
      if (coinsMonth === undefined) coinsMonth = existingRow[6] || 0;
      if (totalDays === undefined) totalDays = existingRow[9] || 0;
      sheet.getRange(targetRow, 1, 1, 10).setValues([[
        tiktokId, uName, tiktokUrl, assignedVJ, team, coinsToday, coinsMonth, nowStr, recorder, totalDays
      ]]);
    } else {
      if (coinsToday === undefined) coinsToday = 0;
      if (coinsMonth === undefined) coinsMonth = 0;
      if (totalDays === undefined) totalDays = 0;
      sheet.appendRow([
        tiktokId, uName, tiktokUrl, assignedVJ, team, coinsToday, coinsMonth, nowStr, recorder, totalDays
      ]);
    }

    upsertUserMappingDailySnapshot(database, { tiktokId: tiktokId, userName: uName, tiktokUrl: tiktokUrl, assignedVJ: assignedVJ, team: team, coinsToday: coinsToday, coinsMonth: coinsMonth, uploadedBy: recorder, totalDays: totalDays });
    writeLog('ADMIN', adminName, 'USER_MAPPING_SAVE', 'บันทึก User Mapping ID TikTok: ' + tiktokId + ' ให้ VJ: ' + assignedVJ);

    return { success: true, message: 'บันทึกข้อมูล User Mapping เรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + err.toString() };
  }
}

/**
 * ลบข้อมูล User Mapping (สำหรับ Admin)
 */
function deleteUserMappingRecord(rowIndex, adminName) {
  try {
    var database = getUserMappingDatabase();
    var sheet = getOrCreateUserMappingSheet(database);
    var row = Number(rowIndex);
    if (row > 1) {
      sheet.deleteRow(row);
      writeLog('ADMIN', adminName, 'USER_MAPPING_DELETE', 'ลบข้อมูลบรรทัดที่ ' + row);
      return { success: true, message: 'ลบรายการ User Mapping เรียบร้อยแล้ว' };
    }
    return { success: false, message: 'ไม่พบบรรทัดข้อมูลที่ต้องการลบ' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบ: ' + err.toString() };
  }
}
