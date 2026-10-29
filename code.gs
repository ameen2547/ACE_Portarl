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
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    Logger.log('Web app rendering failed: ' + error.toString());
    return HtmlService.createHtmlOutput('ไม่สามารถโหลดระบบได้ในขณะนี้ กรุณาติดต่อผู้ดูแลระบบ');
  }
}

// ------------------------------------------
// 🚀 DYNAMIC SYSTEM VERSION & UPDATE RELEASE INFO
// ------------------------------------------

function getOrCreateSystemVersionSheet_() {
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
    // Public read endpoint: do not create or modify sheets as a side effect.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('System_Updates');
    if (!sheet) throw new Error('System_Updates sheet does not exist');
    var data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      var lastRow = data[data.length - 1];
      var changelogText = cleanStr_(lastRow[2]);
      var changelogArr = changelogText ? changelogText.split('\n').filter(function(line) { return line.trim() !== ''; }) : [];
      
      return {
        success: true,
        version: cleanStr_(lastRow[0]) || 'v2.6.0',
        title: cleanStr_(lastRow[1]) || 'ACE Content Portal - อัปเดตระบบ',
        changelog: changelogArr,
        changelogRaw: changelogText,
        updatedAt: lastRow[3] ? (lastRow[3] instanceof Date ? Utilities.formatDate(lastRow[3], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(lastRow[3])) : '-',
        updatedBy: cleanStr_(lastRow[4]) || '-'
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

function saveSystemVersionInfo(version, title, changelogText, adminName, ipAddress, deviceInfo, sessionToken) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    var sheet = getOrCreateSystemVersionSheet_();
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var cleanVer = cleanStr_(version);
    var cleanTitleStr = cleanStr_(title);
    var cleanChangelog = cleanStr_(changelogText);
    var admin = adminName || 'Admin';

    if (!cleanVer || !cleanTitleStr) {
      return { success: false, message: 'โปรดระบุเลขเวอร์ชันและหัวข้อการอัปเดต' };
    }

    sheet.appendRow([cleanVer, cleanTitleStr, cleanChangelog, nowStr, admin]);
    writeLog_('ADMIN', admin, 'UPDATE_SYSTEM_VERSION', 'อัปเดตเวอร์ชันระบบเป็น ' + cleanVer + ' (' + cleanTitleStr + ')', ipAddress, deviceInfo);

    return { 
      success: true, 
      message: '🚀 บันทึกและตั้งค่าเวอร์ชันระบบ ' + cleanVer + ' เรียบร้อยแล้ว',
      versionInfo: getSystemVersionInfo()
    };
  } catch (err) {
    Logger.log('System version update failed: ' + err.toString());
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกเวอร์ชัน' };
  }
}

// ------------------------------------------
// 🛠️ Helper Functions & Logger
// ------------------------------------------

function cleanStr_(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .trim();
}

function safeSheetText_(value) {
  var text = cleanStr_(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

// Server-side bearer sessions keep PINs out of browser storage and out of
// subsequent API calls. Cache entries expire automatically after six hours.
var ACE_SESSION_TTL_SECONDS = 21600;
var ACE_LOGIN_FAILURE_TTL_SECONDS = 900;
var ACE_LOGIN_MAX_FAILURES = 8;

function issueSession_(userType, userName) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var session = {
    userType: cleanStr_(userType).toUpperCase(),
    userName: cleanStr_(userName),
    issuedAt: Date.now()
  };
  CacheService.getScriptCache().put('ACE_SESSION_' + token, JSON.stringify(session), ACE_SESSION_TTL_SECONDS);
  return token;
}

function readSession_(token, expectedType) {
  var cleanToken = cleanStr_(token);
  if (!/^[a-f0-9]{64}$/i.test(cleanToken)) return null;
  var raw = CacheService.getScriptCache().get('ACE_SESSION_' + cleanToken);
  if (!raw) return null;
  try {
    var session = JSON.parse(raw);
    if (!session.userName || !session.userType) return null;
    if (expectedType && session.userType !== cleanStr_(expectedType).toUpperCase()) return null;
    return session;
  } catch (err) {
    return null;
  }
}

function requireSession_(token, expectedType, ipAddress) {
  var session = readSession_(token, expectedType);
  if (!session) return { success: false, message: 'เซสชันหมดอายุหรือไม่มีสิทธิ์ กรุณาเข้าสู่ระบบใหม่' };
  var permission = checkAccessPermission_(session.userName, ipAddress);
  if (!permission.allowed) return { success: false, message: '⛔ ' + permission.reason };
  return { success: true, session: session };
}

function loginFailureCacheKey_(userType, userName) {
  var source = cleanStr_(userType).toUpperCase() + '|' + cleanStr_(userName).toLowerCase();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source);
  return 'ACE_LOGIN_FAIL_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '').substring(0, 32);
}

function isLoginRateLimited_(userType, userName) {
  var count = Number(CacheService.getScriptCache().get(loginFailureCacheKey_(userType, userName))) || 0;
  return count >= ACE_LOGIN_MAX_FAILURES;
}

function recordLoginFailure_(userType, userName) {
  var cache = CacheService.getScriptCache();
  var key = loginFailureCacheKey_(userType, userName);
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(1000);
    var count = Number(cache.get(key)) || 0;
    cache.put(key, String(Math.min(count + 1, ACE_LOGIN_MAX_FAILURES)), ACE_LOGIN_FAILURE_TTL_SECONDS);
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function clearLoginFailures_(userType, userName) {
  CacheService.getScriptCache().remove(loginFailureCacheKey_(userType, userName));
}

function endSession(sessionToken) {
  var token = cleanStr_(sessionToken);
  if (/^[a-f0-9]{64}$/i.test(token)) CacheService.getScriptCache().remove('ACE_SESSION_' + token);
  return { success: true };
}

function parseDiamondVal_(valStr) {
  if (valStr === null || valStr === undefined || valStr === '') return 0;
  if (typeof valStr === 'number') return valStr;
  var cleaned = String(valStr)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .replace(/,/g, '')
    .trim();
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function writeLog_(userType, userName, action, details, ipAddress, deviceInfo) {
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

function getLogs_(selectedMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('Logs') || ss.getSheetByName('LOGS') || ss.getSheetByName('logs');
    if (!logSheet) return { success: true, logs: [] };

    var data = logSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };

    var logs = [];
    var filterMonth = selectedMonth ? cleanStr_(selectedMonth).toUpperCase() : '';

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var timestamp = row[0] ? (row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(row[0])) : '-';
      var userType = cleanStr_(row[1]);
      var userName = cleanStr_(row[2]);
      var action = cleanStr_(row[3]);
      var details = cleanStr_(row[4]);
      var ipAddress = cleanStr_(row[5]);
      var deviceInfo = cleanStr_(row[6]);
      var monthYear = cleanStr_(row[7]);

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

function getOrCreateAccessControlSheet_() {
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

function checkAccessPermission_(userName, ipAddress) {
  try {
    var sheet = getOrCreateAccessControlSheet_();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { allowed: true };
    var headers = data[0].map(function(h) { return cleanStr_(h).toUpperCase().replace(/\s+/g, '_'); });
    function col(name, fallback) { var n = headers.indexOf(name); return n >= 0 ? n : fallback; }
    var typeIdx = col('TYPE', 1), targetIdx = col('TARGET', 2);
    var statusIdx = col('STATUS', headers.indexOf('TEAM') >= 0 ? 4 : 3);
    var reasonIdx = col('REASON', headers.indexOf('TEAM') >= 0 ? 5 : 4);
    var locked = Object.create(null);
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (cleanStr_(row[statusIdx]).toUpperCase() !== 'LOCKED') continue;
      var type = cleanStr_(row[typeIdx]).toUpperCase();
      var target = cleanStr_(row[targetIdx]).toLowerCase();
      if ((type === 'USER' || type === 'IP') && target) locked[type + ':' + target] = cleanStr_(row[reasonIdx]);
    }
    var cleanUser = cleanStr_(userName).toLowerCase();
    var cleanIp = cleanStr_(ipAddress).toLowerCase();
    if (cleanUser && Object.prototype.hasOwnProperty.call(locked, 'USER:' + cleanUser)) {
      return { allowed: false, reason: 'บัญชี VJ/User (' + userName + ') ถูกล็อกการใช้งาน: ' + (locked['USER:' + cleanUser] || 'โปรดติดต่อ Admin') };
    }
    if (cleanIp && Object.prototype.hasOwnProperty.call(locked, 'IP:' + cleanIp)) {
      return { allowed: false, reason: 'IP Address (' + ipAddress + ') ถูกระงับการเข้าถึง: ' + (locked['IP:' + cleanIp] || 'โปรดติดต่อ Admin') };
    }
    return { allowed: true };
  } catch (err) {
    Logger.log('Access check error: ' + err.toString());
    return { allowed: false, reason: 'ไม่สามารถตรวจสอบสิทธิ์ได้ในขณะนี้ กรุณาลองใหม่ภายหลัง' };
  }
}

function getAccessControlList_() {
  try {
    var sheet = getOrCreateAccessControlSheet_();
    var data = sheet.getDataRange().getValues();
    var list = [];
    if (!data.length) return { success: true, list: list };
    var headers = data[0].map(function(h) { return cleanStr_(h).toUpperCase().replace(/\s+/g, '_'); });
    function col(name, fallback) { var n = headers.indexOf(name); return n >= 0 ? n : fallback; }
    var teamIdx = col('TEAM', -1), statusIdx = col('STATUS', teamIdx >= 0 ? 4 : 3);
    var reasonIdx = col('REASON', teamIdx >= 0 ? 5 : 4);
    var byIdx = col('UPDATED_BY', teamIdx >= 0 ? 6 : 5), atIdx = col('UPDATED_AT', teamIdx >= 0 ? 7 : 6);
    for (var i = 1; i < data.length; i++) {
      list.push({ id: cleanStr_(data[i][col('ID', 0)]), type: cleanStr_(data[i][col('TYPE', 1)]), target: cleanStr_(data[i][col('TARGET', 2)]),
        team: teamIdx >= 0 ? (cleanStr_(data[i][teamIdx]) || 'ไม่ระบุทีม') : 'ไม่ระบุทีม', status: cleanStr_(data[i][statusIdx]), reason: cleanStr_(data[i][reasonIdx]),
        updatedBy: cleanStr_(data[i][byIdx]), updatedAt: data[i][atIdx] ? (data[i][atIdx] instanceof Date ? Utilities.formatDate(data[i][atIdx], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(data[i][atIdx])) : '-' });
    }
    return { success: true, list: list };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function updateAccessRule(type, target, team, status, reason, adminName, ipAddress, deviceInfo, sessionToken) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    var sheet = getOrCreateAccessControlSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = (data[0] || []).map(function(h) { return cleanStr_(h).toUpperCase().replace(/\s+/g, '_'); });
    function col(name, fallback) { var n = headers.indexOf(name); return n >= 0 ? n : fallback; }
    var teamIdx = col('TEAM', -1), typeIdx = col('TYPE', 1), targetIdx = col('TARGET', 2);
    var statusIdx = col('STATUS', teamIdx >= 0 ? 4 : 3), reasonIdx = col('REASON', teamIdx >= 0 ? 5 : 4);
    var byIdx = col('UPDATED_BY', teamIdx >= 0 ? 6 : 5), atIdx = col('UPDATED_AT', teamIdx >= 0 ? 7 : 6);
    var cleanType = cleanStr_(type).toUpperCase(), cleanTarget = cleanStr_(target), cleanTeam = cleanStr_(team) || 'ไม่ระบุทีม';
    var cleanStatus = cleanStr_(status).toUpperCase(), nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    if (cleanType !== 'USER' && cleanType !== 'IP') return { success: false, message: 'ประเภทการล็อกไม่ถูกต้อง' };
    if (cleanStatus !== 'LOCKED' && cleanStatus !== 'ACTIVE' && cleanStatus !== 'UNLOCKED') return { success: false, message: 'สถานะการล็อกไม่ถูกต้อง' };
    if (!cleanTarget) return { success: false, message: 'โปรดระบุชื่อ User หรือ IP ที่ต้องการตั้งค่า' };
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) if (cleanStr_(data[i][typeIdx]).toUpperCase() === cleanType && cleanStr_(data[i][targetIdx]).toLowerCase() === cleanTarget.toLowerCase()) { targetRow = i + 1; break; }
    if (targetRow > 1) {
      if (teamIdx >= 0) sheet.getRange(targetRow, teamIdx + 1).setValue(cleanTeam);
      sheet.getRange(targetRow, statusIdx + 1).setValue(cleanStatus);
      sheet.getRange(targetRow, reasonIdx + 1).setValue(cleanStr_(reason));
      sheet.getRange(targetRow, byIdx + 1).setValue(adminName || 'Admin');
      sheet.getRange(targetRow, atIdx + 1).setValue(nowStr);
    } else {
      var newRow = new Array(Math.max(headers.length, 1)).fill('');
      newRow[col('ID', 0)] = 'ACC_' + new Date().getTime(); newRow[typeIdx] = cleanType; newRow[targetIdx] = cleanTarget;
      if (teamIdx >= 0) newRow[teamIdx] = cleanTeam;
      newRow[statusIdx] = cleanStatus; newRow[reasonIdx] = cleanStr_(reason); newRow[byIdx] = adminName || 'Admin'; newRow[atIdx] = nowStr;
      sheet.appendRow(newRow);
    }
    var actText = cleanStatus === 'LOCKED' ? 'LOCK_ACCESS' : 'UNLOCK_ACCESS';
    writeLog_('ADMIN', adminName, actText, 'ตั้งค่าการเข้าถึง ' + cleanType + ': ' + cleanTarget + ' (ทีม: ' + cleanTeam + ') -> ' + cleanStatus + ' (เหตุผล: ' + (reason || '-') + ')', ipAddress, deviceInfo);
    return { success: true, message: (cleanStatus === 'LOCKED' ? '🔒 ล็อกสิทธิ์การใช้งานเรียบร้อยแล้ว' : '🔓 ปลดล็อกการใช้งานเรียบร้อยแล้ว') };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function getProfilePicUrl_(tiktokVal, stageName) {
  var cleanedVal = cleanStr_(tiktokVal);
  if (!cleanedVal) {
    return 'https://ui-avatars.com/api/?name=' + encodeURIComponent(stageName || 'VJ') + '&background=6366f1&color=fff&bold=true';
  }
  if (cleanedVal.indexOf('http://') === 0 || cleanedVal.indexOf('https://') === 0) {
    return cleanedVal;
  }
  var cleanUser = cleanedVal.replace('@', '').trim();
  return 'https://unavatar.io/tiktok/' + encodeURIComponent(cleanUser);
}

function getAvailableMonths_(spreadsheet) {
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

function getKPIMap_(selectedMonth, spreadsheet) {
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
      var h = cleanStr_(hRow[c]).toUpperCase();
      if (h.indexOf('VJ') !== -1 || h.indexOf('NAME') !== -1) nameIdx = c;
      if (h.indexOf('TARGET') !== -1) targetIdx = c;
      if (h.indexOf('MONTH') !== -1) monthIdx = c;
    }

    var cleanSelectedMonth = cleanStr_(selectedMonth).toUpperCase();

    for (var i = 1; i < data.length; i++) {
      var vjName = cleanStr_(data[i][nameIdx]).toLowerCase();
      var targetVal = parseDiamondVal_(data[i][targetIdx]);
      var monthVal = cleanStr_(data[i][monthIdx]).toUpperCase();

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

function getVJListForLogin_(spreadsheet) {
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
        var h = cleanStr_(vjValues[r][c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (h === 'STGENAME' || h === 'STAGENAME' || h === 'ชื่อการแสดง') score += 3;
        if (h === 'TEAM' || h === 'ทีม') score += 1;
        if (h === 'ROLE' || h === 'ตำแหน่ง') score += 1;
      }
      if (score > bestScore) { bestScore = score; headerRow = r; }
    }
    if (headerRow < 0) headerRow = Math.min(1, vjValues.length - 1);

    var headers = vjValues[headerRow] || [];
    var stageIdx = 1, teamIdx = 2, roleIdx = 3;
    for (var c = 0; c < headers.length; c++) {
      var hStr = cleanStr_(headers[c]).toUpperCase().replace(/[\s_-]+/g, '');
      if (hStr === 'STGENAME' || hStr === 'STAGENAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
      else if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
      else if (hStr === 'ROLE' || hStr === 'ตำแหน่ง') roleIdx = c;
    }

    var seen = Object.create(null);
    for (var i = headerRow + 1; i < vjValues.length; i++) {
      var stageName = cleanStr_(vjValues[i][stageIdx]);
      var team = cleanStr_(vjValues[i][teamIdx]);
      var role = cleanStr_(vjValues[i][roleIdx]);
      var key = stageName.toLowerCase();
      if (stageName && !seen[key]) {
        seen[key] = true;
        // Public login bootstrap needs only the fields shown in the picker.
        vjList.push({ stageName: stageName, role: role || '-', team: team || '-' });
      }
    }
    return { success: true, vjList: vjList };
  } catch (err) {
    Logger.log('Error loading VJ login list: ' + err.toString());
    return { success: false, message: 'โหลดรายชื่อ VJ ไม่สำเร็จ: ' + err.toString(), vjList: [] };
  }
}

function getAdminNames_(spreadsheet) {
  try {
    var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName('Admins') || ss.getSheetByName('admins') || ss.getSheetByName('ADMINS');
    var adminList = [];
    if (adminSheet) {
      var adminValues = adminSheet.getDataRange().getValues();
      var hRow = adminValues[0] || [];
      var nameIdx = 0;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr_(hRow[c]).toUpperCase();
        if (hStr === 'NAME' || hStr === 'ADMIN NAME' || hStr === 'ADMIN' || hStr === 'USERNAME' || hStr === 'ชื่อ') {
          nameIdx = c;
          break;
        }
      }
      for (var i = 1; i < adminValues.length; i++) {
        var aName = cleanStr_(adminValues[i][nameIdx]);
        if (aName) adminList.push(aName);
      }
    }
    return { success: true, adminList: adminList };
  } catch (err) {
    Logger.log('Error loading admin names: ' + err.toString());
    return { success: false, adminList: [], message: 'โหลดรายชื่อ Admin ไม่สำเร็จ' };
  }
}

/** Load login-screen data in one Apps Script RPC and reuse one spreadsheet handle. */
function getLoginBootstrapData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    success: true,
    months: getAvailableMonths_(ss),
    vj: getVJListForLogin_(ss),
    admins: getAdminNames_(ss)
  };
}

// ------------------------------------------
// 📢 ANNOUNCEMENTS MODULE
// ------------------------------------------

function getOrCreateAnnouncementSheets_() {
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

function getAnnouncementSchedule_(announcementId, allProperties) {
  try {
    var properties = allProperties || PropertiesService.getScriptProperties().getProperties();
    var raw = properties['ACE_ANN_SCHEDULE_' + announcementId];
    return raw ? JSON.parse(raw) : {};
  } catch (err) {
    return {};
  }
}

function parseAnnouncementDateTime_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var text = cleanStr_(value).replace('T', ' ');
  if (!text) return '';
  try {
    return Utilities.parseDate(text, 'Asia/Bangkok', text.length === 16 ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd HH:mm:ss');
  } catch (err) {
    throw new Error('รูปแบบวันเวลาไม่ถูกต้อง กรุณาเลือกวันและเวลาใหม่');
  }
}

function createAnnouncement(data, adminName, ipAddress, deviceInfo, sessionToken) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    if (!data || typeof data !== 'object') return { success: false, message: 'ข้อมูลประกาศไม่ถูกต้อง' };
    if (cleanStr_(data.title).length > 200 || cleanStr_(data.content).length > 5000) return { success: false, message: 'หัวข้อหรือเนื้อหาประกาศยาวเกินกำหนด' };

    var sheets = getOrCreateAnnouncementSheets_();
    var id = 'ANN_' + new Date().getTime();
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var monthYearStr = Utilities.formatDate(now, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();

    var startsAt = parseAnnouncementDateTime_(data.startsAt);
    var endsAt = parseAnnouncementDateTime_(data.endsAt);
    if (startsAt && endsAt && endsAt.getTime() <= startsAt.getTime()) {
      throw new Error('เวลาปิดประกาศต้องอยู่หลังเวลาเริ่มประกาศ');
    }

    var imageUrl = cleanStr_(data.imageUrl);
    if (imageUrl && !/^https:\/\//i.test(imageUrl)) throw new Error('ลิงก์รูปภาพต้องขึ้นต้นด้วย https://');
    if (data.imageBase64 && data.imageBase64.indexOf('data:image') === 0) {
      try {
        if (data.imageBase64.length > 6000000) throw new Error('ไฟล์ภาพต้องมีขนาดไม่เกิน 4 MB');
        var imageParts = data.imageBase64.match(/^data:(image\/(?:jpeg|png|webp|gif));base64,([A-Za-z0-9+/=]+)$/i);
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
      cleanStr_(data.title),
      cleanStr_(data.content),
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

    writeLog_('ADMIN', adminName, 'CREATE_ANNOUNCEMENT', 'สร้างประกาศข่าวสาร: ' + data.title, ipAddress, deviceInfo);
    return { success: true, id: id, message: 'สร้างประกาศข่าวสารเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getActiveAnnouncements_(userName, userType) {
  try {
    var sheets = getOrCreateAnnouncementSheets_();
    var annData = sheets.annSheet.getDataRange().getValues();
    if (annData.length <= 1) return { success: true, list: [] };

    var readsData = sheets.readSheet.getDataRange().getValues();
    var readMap = Object.create(null);
    for (var r = 1; r < readsData.length; r++) {
      if (cleanStr_(readsData[r][2]).toLowerCase() === cleanStr_(userName).toLowerCase()) {
        var aId = cleanStr_(readsData[r][1]);
        var readKey = aId + '|' + (cleanStr_(readsData[r][3]).toUpperCase() || 'VJ');
        readMap[readKey] = {
          read: true,
          acknowledged: cleanStr_(readsData[r][5]) === 'YES'
        };
      }
    }

    var list = [];
    var now = new Date();
    var scheduleProperties = PropertiesService.getScriptProperties().getProperties();
    for (var i = annData.length - 1; i >= 1; i--) {
      var row = annData[i];
      if (cleanStr_(row[9]) !== 'ACTIVE') continue;
      var annId = cleanStr_(row[0]);
      var schedule = getAnnouncementSchedule_(annId, scheduleProperties);
      var startAt = Number(schedule.startsAt) || 0;
      var endAt = Number(schedule.endsAt) || 0;
      if (startAt && startAt > now.getTime()) continue;
      if (endAt && endAt <= now.getTime()) continue;
      var userReadInfo = readMap[annId + '|' + cleanStr_(userType || 'VJ').toUpperCase()] || { read: false, acknowledged: false };

      list.push({
        id: annId,
        timestamp: row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(row[1]),
        monthYear: cleanStr_(row[2]),
        title: cleanStr_(row[3]),
        content: cleanStr_(row[4]),
        imageUrl: cleanStr_(row[5]),
        startsAt: startAt ? Utilities.formatDate(new Date(startAt), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        endsAt: endAt ? Utilities.formatDate(new Date(endAt), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        isImportant: cleanStr_(row[6]) === 'YES',
        requireAck: cleanStr_(row[7]) === 'YES',
        createdBy: cleanStr_(row[8]),
        isRead: userReadInfo.read,
        isAcknowledged: userReadInfo.acknowledged
      });
    }
    return { success: true, list: list };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function recordAnnouncementRead(announcementId, userName, userType, acknowledged, sessionToken, ipAddress) {
  try {
    var auth = requireSession_(sessionToken, null, ipAddress);
    if (!auth.success) return auth;
    if (auth.session.userType !== 'ADMIN' && auth.session.userType !== 'VJ') return { success: false, message: 'ไม่มีสิทธิ์บันทึกการรับทราบ' };
    userName = auth.session.userName;
    userType = auth.session.userType;
    announcementId = cleanStr_(announcementId);
    if (!announcementId) return { success: false, message: 'ไม่พบรายการประกาศ' };
    var sheets = getOrCreateAnnouncementSheets_();
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var monthYearStr = Utilities.formatDate(now, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();

    var readsData = sheets.readSheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < readsData.length; i++) {
      var existingUserType = cleanStr_(readsData[i][3]).toUpperCase() || 'VJ';
      if (cleanStr_(readsData[i][1]) === announcementId && cleanStr_(readsData[i][2]).toLowerCase() === cleanStr_(userName).toLowerCase() && existingUserType === cleanStr_(userType || 'VJ').toUpperCase()) {
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

function getAnnouncementStats_() {
  try {
    var sheets = getOrCreateAnnouncementSheets_();
    var annData = sheets.annSheet.getDataRange().getValues();
    var readData = sheets.readSheet.getDataRange().getValues();

    if (annData.length <= 1) return { success: true, stats: [], totalAnnouncements: 0, totalViews: 0 };

    var readsByAnn = {};
    for (var r = 1; r < readData.length; r++) {
      var annId = cleanStr_(readData[r][1]);
      if (!readsByAnn[annId]) readsByAnn[annId] = [];
      readsByAnn[annId].push({
        userName: cleanStr_(readData[r][2]),
        userType: cleanStr_(readData[r][3]),
        timestamp: readData[r][4] instanceof Date ? Utilities.formatDate(readData[r][4], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(readData[r][4]),
        acknowledged: cleanStr_(readData[r][5]) === 'YES'
      });
    }

    var stats = [];
    var totalViews = 0;
    var scheduleProperties = PropertiesService.getScriptProperties().getProperties();

    for (var i = annData.length - 1; i >= 1; i--) {
      var row = annData[i];
      var id = cleanStr_(row[0]);
      var schedule = getAnnouncementSchedule_(id, scheduleProperties);
      var nowMs = Date.now();
      var startsAtMs = Number(schedule.startsAt) || 0;
      var endsAtMs = Number(schedule.endsAt) || 0;
      var effectiveStatus = endsAtMs && endsAtMs <= nowMs ? 'EXPIRED' : (startsAtMs && startsAtMs > nowMs ? 'SCHEDULED' : cleanStr_(row[9]));
      var reads = readsByAnn[id] || [];
      var viewCount = reads.length;
      var ackCount = reads.filter(function(m) { return m.acknowledged; }).length;
      totalViews += viewCount;

      stats.push({
        id: id,
        timestamp: row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(row[1]),
        monthYear: cleanStr_(row[2]),
        title: cleanStr_(row[3]),
        content: cleanStr_(row[4]),
        imageUrl: cleanStr_(row[5]),
        startsAt: schedule.startsAt ? Utilities.formatDate(new Date(Number(schedule.startsAt)), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        endsAt: schedule.endsAt ? Utilities.formatDate(new Date(Number(schedule.endsAt)), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm") : '',
        isImportant: cleanStr_(row[6]) === 'YES',
        requireAck: cleanStr_(row[7]) === 'YES',
        createdBy: cleanStr_(row[8]),
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

function getAllMonthData_(selectedMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var months = getAvailableMonths_(ss);
  if (!selectedMonth && months.length > 0) {
    selectedMonth = months[0];
  }
  var monthSheet = ss.getSheetByName(selectedMonth);
  if (!monthSheet) {
    return { success: false, message: 'ไม่พบชีตประจำเดือน: ' + selectedMonth };
  }

  var kpiMap = getKPIMap_(selectedMonth, ss);
  var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
  var vjProfileMap = {};

  if (vjSheet) {
    var vjValues = vjSheet.getDataRange().getValues();
    var hRow = vjValues[0] || [];
    var nameIdx = 0, stageIdx = 1, teamIdx = 2, roleIdx = 3, shiftIdx = 5, tiktokIdx = 7, statusIdx = 8;
    
    for (var c = 0; c < hRow.length; c++) {
      var hStr = cleanStr_(hRow[c]).toUpperCase();
      if (hStr === 'NAME' || hStr === 'ชื่อ') nameIdx = c;
      else if (hStr === 'STGE NAME' || hStr === 'STAGE NAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
      else if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
      else if (hStr === 'ROLE' || hStr === 'ตำแหน่ง') roleIdx = c;
      else if (hStr === 'SHIFT' || hStr === 'เวลาทำงาน' || c === 8) shiftIdx = c;
      else if (hStr === 'ID TIKTOK' || hStr === 'TIKTOK ID' || hStr === 'TIKTOK') tiktokIdx = c;
      else if (hStr === 'STATUS' || hStr === 'สถานะ') statusIdx = c;
    }

    for (var i = 1; i < vjValues.length; i++) {
      var stageName = cleanStr_(vjValues[i][stageIdx]);
      if (!stageName) continue;
      var tiktokVal = tiktokIdx !== -1 ? cleanStr_(vjValues[i][tiktokIdx]) : '';

      vjProfileMap[stageName.toLowerCase()] = {
        fullName: cleanStr_(vjValues[i][nameIdx]) || stageName,
        stageName: stageName,
        team: cleanStr_(vjValues[i][teamIdx]) || '-',
        role: cleanStr_(vjValues[i][roleIdx]) || '-',
        shift: cleanStr_(vjValues[i][shiftIdx]) || '-',
        status: cleanStr_(vjValues[i][statusIdx]) || 'ทำงานอยู่',
        tiktokId: tiktokVal,
        avatarUrl: getProfilePicUrl_(tiktokVal, stageName)
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
    var colA_Team = cleanStr_(row[0]);
    var colE_Stage = cleanStr_(row[4]);
    var monthlyDiamStr = cleanStr_(row[5]);
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
      var rawVal = colIndex < row.length ? cleanStr_(row[colIndex]) : '-';
      var dayVal = parseDiamondVal_(rawVal);
      
      dailyEarnings.push(rawVal !== '' ? rawVal : '-');
      dailySum += dayVal;
    }

    var currentDiamond = parseDiamondVal_(monthlyDiamStr);
    if (currentDiamond === 0 && dailySum > 0) {
      currentDiamond = dailySum;
    }

    var profile = vjProfileMap[lowerStage] || {
      fullName: colE_Stage,
      stageName: colE_Stage,
      team: colA_Team || 'ไม่ระบุทีม',
      role: '-',
      shift: cleanStr_(row[1]) || '-',
      status: 'ทำงานอยู่',
      tiktokId: '',
      avatarUrl: getProfilePicUrl_('', colE_Stage)
    };

    var targetDiamond = kpiMap[lowerStage] !== undefined ? kpiMap[lowerStage] : 100000;
    var remDiamond = Math.max(0, targetDiamond - currentDiamond);
    var progressPercent = targetDiamond > 0 ? Math.round((currentDiamond / targetDiamond) * 100) : 0;
    var status = currentDiamond >= targetDiamond ? 'Completed' : 'In Progress';

    rawList.push({
      vjName: colE_Stage,
      fullName: profile.fullName,
      team: profile.team || colA_Team || 'ไม่ระบุทีม',
      shift: profile.shift || cleanStr_(row[1]) || '-',
      targetDiamond: targetDiamond,
      currentDiamond: currentDiamond,
      remainingDiamond: remDiamond,
      progressPercent: progressPercent,
      status: status,
      dailyEarnings: dailyEarnings,
      profilePic: profile.avatarUrl,
      tiktokId: profile.tiktokId,
      workDays: cleanStr_(row[2]) || '0',
      leaveDays: cleanStr_(row[3]) || '0'
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

function getKpiComparisonData(currentMonth, prevMonth, sessionToken, ipAddress) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    var availableMonths = getAvailableMonths_();
    if (!currentMonth && availableMonths.length > 0) currentMonth = availableMonths[0];
    if (!prevMonth && availableMonths.length > 1) {
      var currIdx = availableMonths.indexOf(currentMonth);
      if (currIdx !== -1 && currIdx + 1 < availableMonths.length) {
        prevMonth = availableMonths[currIdx + 1];
      } else {
        prevMonth = availableMonths[1];
      }
    }

    var currData = getAllMonthData_(currentMonth);
    var prevData = prevMonth ? getAllMonthData_(prevMonth) : { list: [] };

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

function getOrCreateAttendanceSheet_() {
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

function formatDateStandard_(dateInput) {
  if (!dateInput) return '';
  if (dateInput instanceof Date) {
    return Utilities.formatDate(dateInput, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  var str = cleanStr_(dateInput);
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

function formatAttendanceTime_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Bangkok', 'HH:mm');
  if (typeof value === 'number' && value >= 0 && value < 1) {
    var totalMinutes = Math.round(value * 1440);
    return ('0' + Math.floor(totalMinutes / 60) % 24).slice(-2) + ':' + ('0' + (totalMinutes % 60)).slice(-2);
  }
  var raw = cleanStr_(value);
  var match = raw.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
  return match ? ('0' + match[1]).slice(-2) + ':' + match[2] : raw;
}

function attendanceMonthFromDate_(dateText) {
  var m = cleanStr_(dateText).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  var date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  return Utilities.formatDate(date, 'Asia/Bangkok', 'MMMM yyyy').toUpperCase();
}

function getAttendanceSourceSpreadsheet_() {
  return SpreadsheetApp.openById(ATTENDANCE_SOURCE_SPREADSHEET_ID);
}

function getVJShiftMap_(selectedMonth, attendanceSpreadsheet) {
  var map = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
    if (vjSheet) {
      var vValues = vjSheet.getDataRange().getValues();
      var hRow = vValues[0] || [];
      var stageIdx = 1, teamIdx = 2, shiftIdx = 8;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr_(hRow[c]).toUpperCase();
        if (hStr === 'STGE NAME' || hStr === 'STAGE NAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
        else if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
        else if (hStr === 'SHIFT' || hStr === 'เวลาทำงาน' || c === 8) shiftIdx = c;
      }
      for (var i = 1; i < vValues.length; i++) {
        var stage = cleanStr_(vValues[i][stageIdx]);
        if (stage) {
          map[stage.toLowerCase()] = {
            stageName: stage,
            team: cleanStr_(vValues[i][teamIdx]) || 'ไม่ระบุทีม',
            shift: cleanStr_(vValues[i][shiftIdx]) || '-'
          };
        }
      }
    }

    if (selectedMonth) {
      var monthSheet = ss.getSheetByName(selectedMonth);
      if (monthSheet) {
        var mValues = monthSheet.getDataRange().getValues();
        for (var r = 1; r < mValues.length; r++) {
          var team = cleanStr_(mValues[r][0]);
          var shift = cleanStr_(mValues[r][1]);
          var vjName = cleanStr_(mValues[r][4]);
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
    var attendanceSs = attendanceSpreadsheet || getAttendanceSourceSpreadsheet_();
    var settingsSheet = attendanceSs.getSheetByName('Shift_Settings');
    if (settingsSheet) {
      var settings = settingsSheet.getDataRange().getValues();
      var sh = (settings[0] || []).map(function(v) { return cleanStr_(v).toUpperCase().replace(/[\s_-]+/g, ''); });
      var teamCol = sh.indexOf('TEAM'), startCol = sh.indexOf('WORKSTART'), endCol = sh.indexOf('WORKEND');
      if (teamCol >= 0 && startCol >= 0 && endCol >= 0) {
        var teamShifts = Object.create(null);
        for (var si = 1; si < settings.length; si++) {
          var teamName = cleanStr_(settings[si][teamCol]);
          var startTime = formatAttendanceTime_(settings[si][startCol]);
          var endTime = formatAttendanceTime_(settings[si][endCol]);
          if (teamName && startTime && endTime) teamShifts[teamName.toLowerCase()] = startTime + ' - ' + endTime;
        }
        Object.keys(map).forEach(function(key) {
          if (!map[key].shift || map[key].shift === '-' || map[key].shift === 'ไม่ระบุกะ') {
            map[key].shift = teamShifts[cleanStr_(map[key].team).toLowerCase()] || map[key].shift || '-';
          }
        });
      }
    }
  } catch (err) {
    Logger.log('Error in getVJShiftMap_: ' + err.toString());
  }
  return map;
}

function parseTimeToMinutes_(timeStr) {
  if (!timeStr) return null;
  var cleaned = cleanStr_(timeStr);
  var match = cleaned.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return null;
  var hrs = parseInt(match[1], 10);
  var mins = parseInt(match[2], 10);
  return hrs * 60 + mins;
}

function calculateAttendanceMetrics_(clockIn, clockOut, shiftStr) {
  var shiftStartMins = null;
  var shiftEndMins = null;
  var shiftStartStr = '';
  var shiftEndStr = '';

  if (shiftStr && shiftStr !== '-' && shiftStr !== 'ไม่ระบุกะ') {
    var parts = shiftStr.split('-');
    if (parts.length >= 2) {
      shiftStartStr = parts[0].trim();
      shiftEndStr = parts[1].trim();
      shiftStartMins = parseTimeToMinutes_(shiftStartStr);
      shiftEndMins = parseTimeToMinutes_(shiftEndStr);
    }
  }

  if (clockIn && !clockOut && shiftEndStr) {
    clockOut = shiftEndStr;
  }

  var clockInMins = parseTimeToMinutes_(clockIn);
  var clockOutMins = parseTimeToMinutes_(clockOut);

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

function getAttendanceRecords_(selectedMonth, targetVJName) {
  try {
    var filterVJ = targetVJName ? cleanStr_(targetVJName).toLowerCase() : '';
    var filterMonth = selectedMonth ? cleanStr_(selectedMonth).toUpperCase() : '';
    var recordsByKey = Object.create(null);
    var sourceWarning = '';
    var externalSs = null;
    var shiftMap = {};

    try {
      externalSs = getAttendanceSourceSpreadsheet_();
      var logSheet = externalSs.getSheetByName('Attendance_Logs');
      if (!logSheet) throw new Error('ไม่พบชีต Attendance_Logs ในไฟล์ต้นทาง');
      shiftMap = getVJShiftMap_(selectedMonth, externalSs);
      var sourceData = logSheet.getDataRange().getValues();
      if (sourceData.length > 1) {
        var headers = (sourceData[0] || []).map(function(v) { return cleanStr_(v).toUpperCase().replace(/[\s_-]+/g, ''); });
        function sourceCol(names) { for (var x = 0; x < names.length; x++) { var idx = headers.indexOf(names[x]); if (idx >= 0) return idx; } return -1; }
        var idCol = sourceCol(['LOGID','ID']);
        var nameCol = sourceCol(['NAME','STAGENAME','VJNAME']);
        var dateCol = sourceCol(['DATE','WORKDATE']);
        var inCol = sourceCol(['TIMEIN','CLOCKIN','CHECKIN']);
        var outCol = sourceCol(['TIMEOUT','CLOCKOUT','CHECKOUT']);
        if (nameCol < 0 || dateCol < 0 || inCol < 0 || outCol < 0) throw new Error('หัวคอลัมน์ Attendance_Logs ต้องมี Name, Date, TimeIn และ TimeOut');

        for (var i = 1; i < sourceData.length; i++) {
          var row = sourceData[i];
          var dateVal = formatDateStandard_(row[dateCol]);
          var vjName = cleanStr_(row[nameCol]);
          if (!dateVal || !vjName) continue;
          var lowerName = vjName.toLowerCase();
          var monthYear = attendanceMonthFromDate_(dateVal);
          if (filterVJ && lowerName !== filterVJ) continue;
          if (filterMonth && monthYear !== filterMonth && attendanceMonthFromDate_(dateVal) !== filterMonth && dateVal.substring(0, 7) !== filterMonth) continue;
          var mapped = shiftMap[lowerName] || { team: '-', shift: '-' };
          var clockIn = formatAttendanceTime_(row[inCol]);
          var clockOut = formatAttendanceTime_(row[outCol]);
          var metrics = calculateAttendanceMetrics_(clockIn, clockOut, mapped.shift);
          var key = dateVal + '|' + lowerName;
          recordsByKey[key] = {
            id: idCol >= 0 ? cleanStr_(row[idCol]) : 'SOURCE_' + (i + 1), rowIndex: i + 1,
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
        if (!Object.keys(shiftMap).length) shiftMap = getVJShiftMap_(selectedMonth, externalSs || undefined);
        for (var j = 1; j < localData.length; j++) {
          var r = localData[j];
          var localDate = formatDateStandard_(r[1]);
          var localName = cleanStr_(r[3]);
          var localLower = localName.toLowerCase();
          if (!localDate || !localName) continue;
          var localMonth = cleanStr_(r[2]).toUpperCase() || attendanceMonthFromDate_(localDate);
          if (filterVJ && localLower !== filterVJ) continue;
          if (filterMonth && localMonth !== filterMonth && attendanceMonthFromDate_(localDate) !== filterMonth && localDate.substring(0, 7) !== filterMonth) continue;
          var localMapped = shiftMap[localLower] || { team: cleanStr_(r[4]) || '-', shift: cleanStr_(r[5]) || '-' };
          var localIn = formatAttendanceTime_(r[6]);
          var localOut = formatAttendanceTime_(r[7]);
          var localMetrics = calculateAttendanceMetrics_(localIn, localOut, cleanStr_(r[5]) || localMapped.shift);
          recordsByKey[localDate + '|' + localLower] = {
            id: cleanStr_(r[0]), rowIndex: j + 1, date: localDate, monthYear: localMonth,
            vjName: localName, team: cleanStr_(r[4]) || localMapped.team || '-',
            scheduledShift: cleanStr_(r[5]) || localMapped.shift || '-', clockIn: localIn,
            clockOut: localMetrics.clockOutFixed || localOut,
            workHours: r[8] !== '' && r[8] !== null ? Number(r[8]) : localMetrics.workHours,
            lateMinutes: r[9] !== '' && r[9] !== null ? Number(r[9]) : localMetrics.lateMinutes,
            status: cleanStr_(r[10]) || localMetrics.status, source: cleanStr_(r[11]) || 'Manual',
            createdBy: cleanStr_(r[12]) || '-', createdAt: r[13] ? (r[13] instanceof Date ? Utilities.formatDate(r[13], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') : cleanStr_(r[13])) : '-'
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

function uploadBulkAttendance(recordsArray, adminName, selectedMonth, ipAddress, deviceInfo, sessionToken) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    if (!Array.isArray(recordsArray) || recordsArray.length === 0 || recordsArray.length > 1000) {
      return { success: false, message: 'ไม่มีข้อมูลสำหรับอัปโหลด' };
    }

    var sheet = getOrCreateAttendanceSheet_();
    var shiftMap = getVJShiftMap_(selectedMonth);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var createdBy = adminName || 'Admin';

    var data = sheet.getDataRange().getValues();
    var lookupMap = {};
    for (var i = 1; i < data.length; i++) {
      var rDate = formatDateStandard_(data[i][1]);
      var rVJ = cleanStr_(data[i][3]).toLowerCase();
      if (rDate && rVJ) {
        lookupMap[rDate + '_' + rVJ] = {
          rowIndex: i + 1,
          id: data[i][0]
        };
      }
    }

    var count = 0;
    recordsArray.forEach(function(item) {
      var dateVal = formatDateStandard_(item.date);
      var vjName = cleanStr_(item.vjName || item.stageName);
      if (!dateVal || !vjName) return;

      var lowerVJ = vjName.toLowerCase();
      var mappedInfo = shiftMap[lowerVJ] || { team: 'ไม่ระบุทีม', shift: 'ไม่ระบุกะ' };

      var team = item.team || mappedInfo.team;
      var shift = item.shift || mappedInfo.shift;
      var clockIn = cleanStr_(item.checkIn || item.clockIn);
      var clockOut = cleanStr_(item.checkOut || item.clockOut);

      var metrics = calculateAttendanceMetrics_(clockIn, clockOut, shift);
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

    writeLog_('ADMIN', adminName, 'IMPORT_ATTENDANCE_EXCEL', 'อัปโหลดเวลาเข้างานประจำเดือน ' + selectedMonth + ' จำนวน ' + count + ' รายการ', ipAddress, deviceInfo);
    return { success: true, count: count, message: 'นำเข้าข้อมูลและอัปเดตเรียบร้อยแล้ว ' + count + ' รายการ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveSingleAttendanceRecord(recordData, adminName, selectedMonth, ipAddress, deviceInfo, sessionToken) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    if (!recordData || typeof recordData !== 'object') return { success: false, message: 'ข้อมูลเวลาเข้าออกไม่ถูกต้อง' };
    var sheet = getOrCreateAttendanceSheet_();
    var shiftMap = getVJShiftMap_(selectedMonth);

    var dateVal = formatDateStandard_(recordData.date);
    var vjName = cleanStr_(recordData.vjName);
    if (!dateVal || !vjName || vjName.length > 100) return { success: false, message: 'โปรดระบุวันที่และชื่อ VJ ให้ถูกต้อง' };
    var lowerVJ = vjName.toLowerCase();
    var mappedInfo = shiftMap[lowerVJ] || { team: 'ไม่ระบุทีม', shift: 'ไม่ระบุกะ' };

    var team = recordData.team || mappedInfo.team;
    var shift = recordData.shift || mappedInfo.shift;
    var clockIn = cleanStr_(recordData.checkIn || recordData.clockIn);
    var clockOut = cleanStr_(recordData.checkOut || recordData.clockOut);
    
    var metrics = calculateAttendanceMetrics_(clockIn, clockOut, shift);
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
      var rDate = formatDateStandard_(data[i][1]);
      var rVJ = cleanStr_(data[i][3]).toLowerCase();
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
      writeLog_('ADMIN', adminName, 'MANUAL_ATTENDANCE_UPDATE', 'แก้ไขเวลาเข้างาน VJ: ' + vjName + ' วันที่: ' + dateVal, ipAddress, deviceInfo);
    } else {
      var newId = 'ATT_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
      sheet.appendRow([
        newId, dateVal, monthYear, vjName, team, shift, clockIn, finalClockOut, workHrs, lateMins, status, source, createdBy, nowStr
      ]);
      writeLog_('ADMIN', adminName, 'MANUAL_ATTENDANCE_UPDATE', 'บันทึกเวลาเข้างาน VJ: ' + vjName + ' วันที่: ' + dateVal, ipAddress, deviceInfo);
    }

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ------------------------------------------
// 🔓 Authentication & Data Return
// ------------------------------------------

function verifyAndGetPersonalData(vjStageName, inputPassword, selectedMonth, ipAddress, deviceInfo, sessionToken) {
  try {
    var isNewLogin = !sessionToken;
    var session = null;
    var targetName = cleanStr_(vjStageName).toLowerCase();
    if (isNewLogin) {
      if (!targetName || isLoginRateLimited_('VJ', targetName)) return { success: false, message: 'เข้าสู่ระบบไม่สำเร็จ กรุณารอสักครู่แล้วลองใหม่' };
      var permCheck = checkAccessPermission_(vjStageName, ipAddress);
      if (!permCheck.allowed) return { success: false, message: '⛔ ' + permCheck.reason };
    } else {
      var auth = requireSession_(sessionToken, 'VJ', ipAddress);
      if (!auth.success) return auth;
      session = auth.session;
      vjStageName = session.userName;
      targetName = cleanStr_(vjStageName).toLowerCase();
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vjSheet = ss.getSheetByName('vj mc all') || ss.getSheetByName('VJ MC ALL') || ss.getSheetByName('vj_mc_all');
    var targetPass = cleanStr_(inputPassword);
    if (!vjSheet) return { success: false, message: 'ไม่พบข้อมูลบัญชีผู้ใช้' };
    var vjValues = vjSheet.getDataRange().getValues();
    var hRow = vjValues[0] || [];
    var stageIdx = 1, passIdx = -1, nameIdx = 0, teamIdx = 2, roleIdx = 3, shiftIdx = 5, tiktokIdx = 7;
    for (var c = 0; c < hRow.length; c++) {
      var hStr = cleanStr_(hRow[c]).toUpperCase().replace(/[\s_-]+/g, '');
      if (hStr === 'STGENAME' || hStr === 'STAGENAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
      if (hStr === 'NAME' || hStr === 'FULLNAME' || hStr === 'ชื่อ') nameIdx = c;
      if (hStr === 'TEAM' || hStr === 'ทีม') teamIdx = c;
      if (hStr === 'ROLE' || hStr === 'ตำแหน่ง') roleIdx = c;
      if (hStr === 'SHIFT' || hStr === 'เวลาทำงาน') shiftIdx = c;
      if (hStr === 'IDTIKTOK' || hStr === 'TIKTOKID' || hStr === 'TIKTOK') tiktokIdx = c;
      if (hStr === 'PASSWORD' || hStr === 'PASSCODE' || hStr === 'PIN' || hStr === 'รหัสผ่าน') passIdx = c;
    }
    var userProfile = null;
    for (var i = 1; i < vjValues.length; i++) {
      if (cleanStr_(vjValues[i][stageIdx]).toLowerCase() !== targetName) continue;
      var storedPin = passIdx >= 0 ? cleanStr_(vjValues[i][passIdx]) : '';
      if (isNewLogin && (!storedPin || !targetPass || storedPin !== targetPass)) {
        recordLoginFailure_('VJ', targetName);
        writeLog_('VJ', vjStageName, 'LOGIN_FAILED', 'รหัสผ่านไม่ถูกต้อง', ipAddress, deviceInfo);
        return { success: false, message: 'ชื่อบัญชีหรือรหัสผ่านไม่ถูกต้อง' };
      }
      userProfile = {
        stageName: cleanStr_(vjValues[i][stageIdx]) || cleanStr_(vjStageName),
        fullName: cleanStr_(vjValues[i][nameIdx]) || cleanStr_(vjStageName),
        role: cleanStr_(vjValues[i][roleIdx]),
        team: cleanStr_(vjValues[i][teamIdx]) || '-',
        shift: cleanStr_(vjValues[i][shiftIdx]) || '-',
        profilePic: getProfilePicUrl_(cleanStr_(vjValues[i][tiktokIdx]), vjStageName)
      };
      break;
    }
    if (!userProfile) {
      if (isNewLogin) recordLoginFailure_('VJ', targetName);
      writeLog_('VJ', vjStageName, 'LOGIN_FAILED', 'ไม่พบบัญชีผู้ใช้', ipAddress, deviceInfo);
      return { success: false, message: 'ชื่อบัญชีหรือรหัสผ่านไม่ถูกต้อง' };
    }
    if (isNewLogin) {
      clearLoginFailures_('VJ', targetName);
    }

    var allData = getAllMonthData_(selectedMonth);
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

    var attRes = getAttendanceRecords_(selectedMonth || '', vjStageName);
    var attendanceList = attRes.success ? attRes.records : [];

    var annRes = getActiveAnnouncements_(vjStageName, 'VJ');
    var announcements = annRes.success ? annRes.list : [];

    if (isNewLogin) {
      sessionToken = issueSession_('VJ', userProfile.stageName);
      writeLog_('VJ', userProfile.stageName, 'LOGIN_SUCCESS', 'เข้าสู่ระบบเดือน: ' + (selectedMonth || '-'), ipAddress, deviceInfo);
    }

    return {
      success: true,
      month: selectedMonth,
      personalData: personalData,
      attendanceOnly: attendanceOnly,
      attendanceList: attendanceList,
      attendanceSourceWarning: attRes.sourceWarning || '',
      announcements: announcements,
      systemVersionInfo: getSystemVersionInfo(),
      sessionToken: sessionToken
    };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}

function getAdminDashboardData(adminName, adminPin, selectedMonth, ipAddress, deviceInfo, sessionToken) {
  try {
    var isNewLogin = !sessionToken;
    var session = null;
    var targetName = cleanStr_(adminName).toLowerCase();
    if (isNewLogin) {
      if (!targetName || isLoginRateLimited_('ADMIN', targetName)) return { success: false, message: 'เข้าสู่ระบบไม่สำเร็จ กรุณารอสักครู่แล้วลองใหม่' };
      var permission = checkAccessPermission_(adminName, ipAddress);
      if (!permission.allowed) return { success: false, message: '⛔ ' + permission.reason };
    } else {
      var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
      if (!auth.success) return auth;
      session = auth.session;
      adminName = session.userName;
      targetName = cleanStr_(adminName).toLowerCase();
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName('Admins') || ss.getSheetByName('admins') || ss.getSheetByName('ADMINS');
    var actualAdminName = '';
    if (isNewLogin) {
      var inputPin = cleanStr_(adminPin);
      if (adminSheet && inputPin) {
        var adminValues = adminSheet.getDataRange().getValues();
        var adminHRow = adminValues[0] || [];
        var nameIdx = 0, pinIdx = 1;
        for (var c = 0; c < adminHRow.length; c++) {
          var hStr = cleanStr_(adminHRow[c]).toUpperCase();
          if (hStr === 'NAME' || hStr === 'ADMIN NAME' || hStr === 'ADMIN' || hStr === 'USERNAME' || hStr === 'ชื่อ') nameIdx = c;
          if (hStr === 'PIN' || hStr === 'PASSWORD' || hStr === 'PASSCODE' || hStr === 'รหัสผ่าน') pinIdx = c;
        }
        for (var i = 1; i < adminValues.length; i++) {
          var rowAdminName = cleanStr_(adminValues[i][nameIdx]);
          var storedPin = cleanStr_(adminValues[i][pinIdx]);
          if (rowAdminName.toLowerCase() === targetName && storedPin && storedPin === inputPin) {
            actualAdminName = rowAdminName;
            break;
          }
        }
      }
      if (!actualAdminName) {
        recordLoginFailure_('ADMIN', targetName);
        writeLog_('ADMIN', adminName, 'LOGIN_FAILED', 'บัญชี Admin หรือรหัสผ่านไม่ถูกต้อง', ipAddress, deviceInfo);
        return { success: false, message: 'บัญชี Admin หรือรหัสผ่านไม่ถูกต้อง' };
      }
      adminName = actualAdminName;
      clearLoginFailures_('ADMIN', targetName);
    }

    var allData = (typeof getAllMonthData_ === 'function') ? getAllMonthData_(selectedMonth) : null;
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

    var attRes = (typeof getAttendanceRecords_ === 'function') ? getAttendanceRecords_(selectedMonth, null) : { success: false };
    var attendanceList = (attRes && attRes.success) ? attRes.records : [];

    var logsRes = (typeof getLogs_ === 'function') ? getLogs_(selectedMonth) : { success: false };
    var logsList = (logsRes && logsRes.success) ? logsRes.logs : [];

    var annRes = (typeof getActiveAnnouncements_ === 'function') ? getActiveAnnouncements_(adminName, 'ADMIN') : { success: false };
    var announcements = (annRes && annRes.success) ? annRes.list : [];

    var annStatsRes = (typeof getAnnouncementStats_ === 'function') ? getAnnouncementStats_() : null;
    var announcementStats = (annStatsRes && annStatsRes.success) ? annStatsRes : { stats: [], totalAnnouncements: 0, totalViews: 0 };

    var accessRes = getAccessControlList_();
    var accessList = accessRes.success ? accessRes.list : [];

    if (isNewLogin) {
      sessionToken = issueSession_('ADMIN', adminName);
      writeLog_('ADMIN', adminName, 'LOGIN_SUCCESS', 'เข้าสู่ระบบ Admin Dashboard เดือน: ' + (selectedMonth || '-'), ipAddress, deviceInfo);
    }

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
      systemVersionInfo: getSystemVersionInfo(),
      sessionToken: sessionToken
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

function getUserMappingDatabase_() {
  return SpreadsheetApp.openById(USER_MAPPING_DATABASE_ID);
}

function getOrCreateUserMappingSheet_(spreadsheet) {
  var ss = spreadsheet || getUserMappingDatabase_();
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

function getOrCreateUserMappingHistorySheet_(spreadsheet) {
  var ss = spreadsheet || getUserMappingDatabase_();
  var sheet = ss.getSheetByName(USER_MAPPING_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USER_MAPPING_HISTORY_SHEET);
    sheet.getRange(1, 1, 1, USER_MAPPING_HISTORY_HEADERS.length).setValues([USER_MAPPING_HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, USER_MAPPING_HISTORY_HEADERS.length).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function userMappingSnapshotDate_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
}

function userMappingHistoryRow_(snapshotDate, row) {
  return [snapshotDate, cleanStr_(row[0]), cleanStr_(row[1]), cleanStr_(row[2]), cleanStr_(row[3]), cleanStr_(row[4]), row[5], row[6], row[7], cleanStr_(row[8]), row[9]];
}

// Archive every current mapping once per day. Subsequent edits update that day's snapshot,
// while earlier dates remain immutable and A:J in User_Mapping remain unchanged.
function archiveUserMappingDailySnapshot_(spreadsheet, liveRows, historyData) {
  var lock = LockService.getScriptLock();
  // Snapshotting is only a first-load fallback; do not make every reader wait
  // behind a script-wide lock when another user is opening the same page.
  if (!lock.tryLock(1000)) return null;
  try {
    var history = getOrCreateUserMappingHistorySheet_(spreadsheet);
    historyData = historyData || history.getDataRange().getValues();
    var snapshotDate = userMappingSnapshotDate_();
    var existing = Object.create(null);
    for (var i = 1; i < historyData.length; i++) {
      var id = cleanStr_(historyData[i][1]).toLowerCase();
      if (id) existing[cleanStr_(historyData[i][0]) + '|' + id] = i + 1;
    }
    var newRows = [];
    for (var r = 1; r < liveRows.length; r++) {
      var row = liveRows[r];
      var id = cleanStr_(row[0]);
      if (!id) continue;
      var key = snapshotDate + '|' + id.toLowerCase();
      if (!existing[key]) newRows.push(userMappingHistoryRow_(snapshotDate, row));
    }
    if (newRows.length) {
      history.getRange(history.getLastRow() + 1, 1, newRows.length, USER_MAPPING_HISTORY_HEADERS.length).setValues(newRows);
      historyData = historyData.concat(newRows);
    }
    return historyData;
  } finally {
    lock.releaseLock();
  }
}

function archiveUserMappingDailySnapshotJob_() {
  var database = getUserMappingDatabase_();
  var sheet = getOrCreateUserMappingSheet_(database);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) archiveUserMappingDailySnapshot_(database, data);
}

function installUserMappingDailyArchiveTrigger_() {
  var handler = 'archiveUserMappingDailySnapshotJob_';
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function(trigger) { return trigger.getHandlerFunction() === handler; });
  if (!exists) ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(1).inTimezone('Asia/Bangkok').create();
  return { success: true, message: exists ? 'มี Trigger สำรองรายวันอยู่แล้ว' : 'ติดตั้ง Trigger สำรอง User Mapping รายวันเวลา 01:00 แล้ว' };
}

function upsertUserMappingDailySnapshot_(spreadsheet, mappingData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var history = getOrCreateUserMappingHistorySheet_(spreadsheet);
    var snapshotDate = userMappingSnapshotDate_();
    var id = cleanStr_(mappingData.tiktokId);
    if (!id) return;
    var values = history.getDataRange().getValues();
    var key = snapshotDate + '|' + id.toLowerCase();
    var target = -1;
    for (var i = 1; i < values.length; i++) {
      if (cleanStr_(values[i][0]) + '|' + cleanStr_(values[i][1]).toLowerCase() === key) { target = i + 1; break; }
    }
    var snapshot = [snapshotDate, id, cleanStr_(mappingData.userName), cleanStr_(mappingData.tiktokUrl), cleanStr_(mappingData.assignedVJ), cleanStr_(mappingData.team), mappingData.coinsToday || 0, mappingData.coinsMonth || 0, Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'), cleanStr_(mappingData.uploadedBy), mappingData.totalDays || 0];
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

function getOrCreateUserMappingAvatarSheet_(spreadsheet) {
  var ss = spreadsheet || getUserMappingDatabase_();
  var sheet = ss.getSheetByName(USER_MAPPING_AVATAR_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USER_MAPPING_AVATAR_SHEET);
    sheet.getRange(1, 1, 1, USER_MAPPING_AVATAR_HEADERS.length).setValues([USER_MAPPING_AVATAR_HEADERS]);
    sheet.getRange(1, 1, 1, USER_MAPPING_AVATAR_HEADERS.length).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getUserMappingAvatarFolder_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(USER_MAPPING_AVATAR_FOLDER_PROPERTY);
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { props.deleteProperty(USER_MAPPING_AVATAR_FOLDER_PROPERTY); }
  }
  var folder = DriveApp.createFolder('ACE User Mapping Avatar Cache');
  props.setProperty(USER_MAPPING_AVATAR_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function authorizeUserMappingAvatarCache_() {
  getUserMappingDatabase_();
  getOrCreateUserMappingAvatarSheet_();
  getUserMappingAvatarFolder_();
  UrlFetchApp.fetch('https://unavatar.io/', { muteHttpExceptions: true, followRedirects: true });
  return 'Avatar cache authorization ready. Drive folder is private by default.';
}

function getUserMappingAvatarData(avatarIds, userName, userType, sessionToken, ipAddress) {
  try {
    var auth = requireSession_(sessionToken, null, ipAddress);
    if (!auth.success) return { success: false, message: auth.message, avatars: {} };
    userType = auth.session.userType;
    userName = auth.session.userName;
    var ids = Array.isArray(avatarIds) ? avatarIds.map(function(id) { return cleanStr_(id); }).filter(Boolean).slice(0, 10) : [];
    if (!ids.length) return { success: true, avatars: {} };
    var database = getUserMappingDatabase_();
    var mappingSheet = getOrCreateUserMappingSheet_(database);
    var mappingRows = mappingSheet.getDataRange().getValues();
    var cleanType = cleanStr_(userType).toUpperCase();
    var cleanUser = cleanStr_(userName).toLowerCase();
    if (cleanType !== 'ADMIN' && cleanType !== 'VJ') return { success: false, message: 'ไม่สามารถตรวจสอบสิทธิ์โหลดภาพได้', avatars: {} };
    var allowed = Object.create(null);
    for (var r = 1; r < mappingRows.length; r++) {
      var id = cleanStr_(mappingRows[r][0]);
      var assigned = cleanStr_(mappingRows[r][3]).toLowerCase();
      if (id && (cleanType === 'ADMIN' || assigned === cleanUser)) allowed[id.toLowerCase()] = id;
    }
    var historySheetForAccess = database.getSheetByName(USER_MAPPING_HISTORY_SHEET);
    if (cleanType === 'VJ' && historySheetForAccess) {
      var historicalRows = historySheetForAccess.getDataRange().getValues();
      for (var hr = 1; hr < historicalRows.length; hr++) {
        var historicalId = cleanStr_(historicalRows[hr][1]);
        var historicalVj = cleanStr_(historicalRows[hr][4]).toLowerCase();
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
        var h = cleanStr_(headers[c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (h === 'STGENAME' || h === 'STAGENAME' || h === 'ชื่อการแสดง') stageCol = c;
        if (h === 'IDTIKTOK' || h === 'TIKTOKID' || h === 'TIKTOK') idCol = c;
      }
      for (var v = 1; v < vjValues.length; v++) {
        var stage = cleanStr_(vjValues[v][stageCol]);
        var vjId = cleanStr_(vjValues[v][idCol]);
        if (vjId && (cleanType === 'ADMIN' || stage.toLowerCase() === cleanUser)) allowed[vjId.toLowerCase()] = vjId;
      }
    }
    var requested = Object.create(null);
    ids.forEach(function(id) { if (allowed[id.toLowerCase()]) requested[id.toLowerCase()] = allowed[id.toLowerCase()]; });
    var cacheSheet = getOrCreateUserMappingAvatarSheet_(database);
    var cacheRows = cacheSheet.getDataRange().getValues();
    var cacheMap = Object.create(null);
    for (var k = 1; k < cacheRows.length; k++) {
      var cacheId = cleanStr_(cacheRows[k][0]).toLowerCase();
      if (cacheId) cacheMap[cacheId] = { rowIndex: k + 1, fileId: cleanStr_(cacheRows[k][1]), mimeType: cleanStr_(cacheRows[k][2]) };
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
        var sourceUrl = getProfilePicUrl_(canonicalId, canonicalId);
        if (!sourceUrl || sourceUrl.indexOf('ui-avatars.com') !== -1) return;
        var response = UrlFetchApp.fetch(sourceUrl, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 ACEContentPortal/1.0' } });
        var status = response.getResponseCode();
        var mimeType = cleanStr_(response.getHeaders()['Content-Type'] || response.getHeaders()['content-type']).split(';')[0].toLowerCase();
        if (status < 200 || status >= 300 || !/^image\/(jpeg|png|webp|gif)$/.test(mimeType)) return;
        blob = response.getBlob().setContentType(mimeType);
        if (blob.getBytes().length > 1500000) return;
        if (!folder) folder = getUserMappingAvatarFolder_();
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

function getUserMappingData(userName, userType, sessionToken, ipAddress, historyOffset) {
  try {
    var auth = requireSession_(sessionToken, null, ipAddress);
    if (!auth.success) return { success: false, message: auth.message, list: [], historyList: [] };
    userName = auth.session.userName;
    userType = auth.session.userType;
    var database = getUserMappingDatabase_();
    var sheet = getOrCreateUserMappingSheet_(database);
    var data = sheet.getDataRange().getValues();
    if (!data || !data.length) data = [[]];
    var historySheet = getOrCreateUserMappingHistorySheet_(database);
    var historyData = historySheet.getDataRange().getValues();
    var todayStr = userMappingSnapshotDate_();
    // The daily archive is already maintained by writes and the time trigger.
    // Only perform the full-sheet fallback once per day (or after cache expiry).
    var archiveCache = CacheService.getScriptCache();
    var archiveCacheKey = 'ACE_USER_MAPPING_ARCHIVED_' + todayStr;
    if (data.length > 1 && !archiveCache.get(archiveCacheKey)) {
      var archivedData = archiveUserMappingDailySnapshot_(database, data, historyData);
      if (archivedData) {
        historyData = archivedData;
        archiveCache.put(archiveCacheKey, '1', 21600);
      }
    }
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    var vjSheet = activeSs.getSheetByName('vj mc all') || activeSs.getSheetByName('VJ MC ALL') || activeSs.getSheetByName('vj_mc_all');
    var vjProfileMap = Object.create(null);
    if (vjSheet) {
      var vjValues = vjSheet.getDataRange().getValues();
      var hRow = vjValues[0] || [];
      var stageIdx = 1, tiktokIdx = 7;
      for (var c = 0; c < hRow.length; c++) {
        var hStr = cleanStr_(hRow[c]).toUpperCase().replace(/[\s_-]+/g, '');
        if (hStr === 'STGENAME' || hStr === 'STAGENAME' || hStr === 'ชื่อการแสดง') stageIdx = c;
        if (hStr === 'IDTIKTOK' || hStr === 'TIKTOKID' || hStr === 'TIKTOK') tiktokIdx = c;
      }
      for (var v = 1; v < vjValues.length; v++) {
        var stage = cleanStr_(vjValues[v][stageIdx]);
        if (stage) vjProfileMap[stage.toLowerCase()] = tiktokIdx >= 0 ? cleanStr_(vjValues[v][tiktokIdx]) : '';
      }
    }

    var cleanUser = cleanStr_(userName).toLowerCase();
    var cleanType = cleanStr_(userType).toUpperCase();
    function makeItem(row, index, isHistory) {
      var offset = isHistory ? 1 : 0;
      var tiktokId = cleanStr_(row[offset]);
      var uName = cleanStr_(row[offset + 1]);
      var tiktokUrl = cleanStr_(row[offset + 2]);
      var assignedVJ = cleanStr_(row[offset + 3]);
      var team = cleanStr_(row[offset + 4]);
      var updatedRaw = row[offset + 7];
      var updatedText = '-';
      var updatedDateOnly = '';
      if (updatedRaw) {
        if (updatedRaw instanceof Date) {
          updatedText = Utilities.formatDate(updatedRaw, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
          updatedDateOnly = Utilities.formatDate(updatedRaw, 'Asia/Bangkok', 'yyyy-MM-dd');
        } else {
          updatedText = cleanStr_(updatedRaw);
          var match = updatedText.match(/^(\d{4}-\d{2}-\d{2})/);
          if (match) updatedDateOnly = match[1];
        }
      }
      var snapshotDate = isHistory ? cleanStr_(row[0]) : updatedDateOnly;
      return {
        rowIndex: isHistory ? null : index + 1,
        snapshotDate: snapshotDate,
        tiktokId: tiktokId, userName: uName, tiktokUrl: tiktokUrl,
        assignedVJ: assignedVJ, team: team,
        coinsToday: parseDiamondVal_(row[offset + 5]), coinsMonth: parseDiamondVal_(row[offset + 6]),
        updatedDate: updatedText, recordedBy: cleanStr_(row[offset + 8]), totalDays: parseDiamondVal_(row[offset + 9]),
        userAvatar: getProfilePicUrl_(tiktokId, uName || tiktokId),
        vjTikTokId: vjProfileMap[assignedVJ.toLowerCase()] || '',
        vjAvatar: getProfilePicUrl_(vjProfileMap[assignedVJ.toLowerCase()] || '', assignedVJ),
        isUpdatedToday: (isHistory ? snapshotDate : updatedDateOnly) === todayStr
      };
    }

    var list = [], historyList = [], todayUpdatedCount = 0;
    for (var i = 1; i < data.length; i++) {
      var liveRow = data[i];
      // Filter VJ rows before avatar/date formatting work. This is especially
      // important for mobile requests where the sheet may contain many VJs.
      if (cleanType === 'VJ' && cleanStr_(liveRow[3]).toLowerCase() !== cleanUser) continue;
      var item = makeItem(liveRow, i, false);
      if (!item.tiktokId && !item.userName && !item.assignedVJ) continue;
      list.push(item);
      if (item.isUpdatedToday) todayUpdatedCount++;
    }
    var matchingHistoryRows = [];
    for (var h = 1; h < historyData.length; h++) {
      var historyRow = historyData[h];
      if (!cleanStr_(historyRow[1]) && !cleanStr_(historyRow[2]) && !cleanStr_(historyRow[4])) continue;
      if (cleanType === 'VJ' && cleanStr_(historyRow[4]).toLowerCase() !== cleanUser) continue;
      matchingHistoryRows.push({ row: historyRow, index: h });
    }
    var safeHistoryOffset = Math.max(0, Math.floor(Number(historyOffset) || 0));
    var historyPageSize = 100;
    // Snapshot rows are appended by date. Read the newest page from the end of
    // the sheet instead of sorting the full archive on every mobile request.
    var historyEnd = Math.max(0, matchingHistoryRows.length - safeHistoryOffset);
    var historyStart = Math.max(0, historyEnd - historyPageSize);
    var historyPage = matchingHistoryRows.slice(historyStart, historyEnd).reverse();
    historyPage.forEach(function(entry) {
      historyList.push(makeItem(entry.row, entry.index, true));
    });
    historyList.sort(function(a,b) { return a.snapshotDate === b.snapshotDate ? a.team.localeCompare(b.team) || a.assignedVJ.localeCompare(b.assignedVJ) || a.userName.localeCompare(b.userName) : (a.snapshotDate < b.snapshotDate ? 1 : -1); });
    return {
      success: true,
      userType: cleanType,
      list: list,
      historyList: historyList,
      historyOffset: safeHistoryOffset,
      historyNextOffset: safeHistoryOffset + historyList.length,
      historyHasMore: historyStart > 0,
      historyTotalCount: matchingHistoryRows.length,
      totalCount: list.length,
      todayUpdatedCount: todayUpdatedCount,
      todayStr: todayStr
    };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในระบบ User Mapping: ' + err.toString(), list: [], historyList: [] };
  }
}

/**
 * บันทึกหรืออัปเดตข้อมูล User Mapping (สำหรับ Admin)
 */
function saveUserMappingRecord(mappingData, adminName, sessionToken, ipAddress) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    if (!mappingData || typeof mappingData !== 'object') return { success: false, message: 'ข้อมูล User Mapping ไม่ถูกต้อง' };
    var database = getUserMappingDatabase_();
    var sheet = getOrCreateUserMappingSheet_(database);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var recorder = adminName || 'Admin';

    var tiktokId = safeSheetText_(mappingData.tiktokId);
    var uName = safeSheetText_(mappingData.userName);
    var tiktokUrl = cleanStr_(mappingData.tiktokUrl);
    var assignedVJ = safeSheetText_(mappingData.assignedVJ);
    var team = safeSheetText_(mappingData.team);
    var coinsToday = mappingData.coinsToday;
    var coinsMonth = mappingData.coinsMonth;
    var totalDays = mappingData.totalDays;

    if (!tiktokId || !uName || !assignedVJ) {
      return { success: false, message: 'โปรดระบุ ID TikTok, User Name และ Assigned VJ ให้ครบถ้วน' };
    }
    if (tiktokId.length > 100 || uName.length > 150 || assignedVJ.length > 100 || team.length > 100) return { success: false, message: 'ข้อมูล User Mapping ยาวเกินกำหนด' };
    if (tiktokUrl && !/^https:\/\/(?:[a-z0-9-]+\.)*tiktok\.com(?:\/|$)/i.test(tiktokUrl)) return { success: false, message: 'ลิงก์โปรไฟล์ต้องเป็น HTTPS ของ TikTok' };

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;

    if (mappingData.rowIndex && Number(mappingData.rowIndex) > 1) {
      targetRow = Number(mappingData.rowIndex);
      if (!Number.isInteger(targetRow) || targetRow > sheet.getLastRow()) return { success: false, message: 'ไม่พบบรรทัด User Mapping ที่ต้องการแก้ไข' };
    } else {
      for (var i = 1; i < data.length; i++) {
        if (cleanStr_(data[i][0]).toLowerCase() === tiktokId.toLowerCase()) {
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

    upsertUserMappingDailySnapshot_(database, { tiktokId: tiktokId, userName: uName, tiktokUrl: tiktokUrl, assignedVJ: assignedVJ, team: team, coinsToday: coinsToday, coinsMonth: coinsMonth, uploadedBy: recorder, totalDays: totalDays });
    writeLog_('ADMIN', adminName, 'USER_MAPPING_SAVE', 'บันทึก User Mapping ID TikTok: ' + tiktokId + ' ให้ VJ: ' + assignedVJ);

    return { success: true, message: 'บันทึกข้อมูล User Mapping เรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + err.toString() };
  }
}

/**
 * ลบข้อมูล User Mapping (สำหรับ Admin)
 */
function deleteUserMappingRecord(rowIndex, adminName, sessionToken, ipAddress) {
  try {
    var auth = requireSession_(sessionToken, 'ADMIN', ipAddress);
    if (!auth.success) return auth;
    adminName = auth.session.userName;
    var database = getUserMappingDatabase_();
    var sheet = getOrCreateUserMappingSheet_(database);
    var row = Number(rowIndex);
    if (Number.isInteger(row) && row > 1 && row <= sheet.getLastRow()) {
      sheet.deleteRow(row);
      writeLog_('ADMIN', adminName, 'USER_MAPPING_DELETE', 'ลบข้อมูลบรรทัดที่ ' + row);
      return { success: true, message: 'ลบรายการ User Mapping เรียบร้อยแล้ว' };
    }
    return { success: false, message: 'ไม่พบบรรทัดข้อมูลที่ต้องการลบ' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบ: ' + err.toString() };
  }
}
