// ============================================================
//  BioAttend – Google Apps Script Backend (CORS Fixed)
//  Deploy: Web App → Execute as Me → Access: Anyone
// ============================================================

var SHEET_USERS      = 'Users';
var SHEET_ATTENDANCE = 'Attendance';
var SHEET_SESSIONS   = 'Sessions';

var COLLEGE_LAT    = 13.3318;
var COLLEGE_LNG    = 77.1274;
var FENCE_RADIUS_M = 100;
var COOLDOWN_HOURS = 3;

var CAMPUS_SSIDS = [
  'SIT-WiFi', 'SIT_Campus', 'SIT-Student', 'SIT-Staff',
  'Siddaganga', 'SIT_Tumkur', 'sit-wifi', 'sit_campus',
  'SIT-Wireless', 'Airtel_Vodka'
];

function doOptions(e) {
  var output = ContentService.createTextOutput('');
  output.addHeader('Access-Control-Allow-Origin', '*');
  output.addHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  output.addHeader('Access-Control-Allow-Headers', 'Content-Type');
  return output;
}

function jsonOut(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  output.addHeader('Access-Control-Allow-Origin', '*');
  return output;
}

function doGet(e) {
  try {
    var param = (e && e.parameter) ? e.parameter : {};
    if (!param.data) {
      return jsonOut({ status: 'BioAttend API running', time: new Date().toString() });
    }
    var body = JSON.parse(decodeURIComponent(param.data));
    return jsonOut(route(body));
  } catch (err) {
    return jsonOut({ success: false, message: 'doGet error: ' + err.toString() });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return jsonOut(route(body));
  } catch (err) {
    return jsonOut({ success: false, message: 'doPost error: ' + err.toString() });
  }
}

function route(body) {
  switch (body.action) {
    case 'register':        return registerUser(body);
    case 'signIn':          return signInUser(body);
    case 'markAttendance':  return markAttendance(body);
    case 'saveBiometric':   return saveBiometric(body);
    case 'getBiometric':    return getBiometric(body);
    case 'createSession':   return createSession(body);
    case 'getActiveSession':return getActiveSession(body);
    case 'getSessions':     return getSessions(body);
    case 'getAttendance':   return getAttendance(body);
    case 'registerDevice':  return registerDevice(body);
    case 'checkDevice':     return checkDevice(body);
    case 'checkWifi':       return checkWifi(body);
    case 'debug':           return debugInfo();
    default: return { success: false, message: 'Unknown action: ' + body.action };
  }
}

var USER_COLS = ['UserID', 'FullName', 'Email', 'PasswordHash', 'DOB', 'Mobile', 'Institution', 'Department', 'MarkFromAnywhere', 'BiometricCredentialId', 'CreatedAt', 'Role', 'DeviceId'];

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var headers = {
      Users: USER_COLS,
      Attendance: ['AttendanceID','UserID','FullName','Email','SessionID','Subject','Timestamp','Date','Time','Method','Lat','Lng','DistanceFromCollege'],
      Sessions: ['SessionID','TeacherID','TeacherName','Subject','Date','StartTime','EndTime','Status','WindowMinutes']
    };
    if (headers[name]) {
      sheet.appendRow(headers[name]);
      sheet.getRange(1,1,1,headers[name].length).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    }
  } else if (name === 'Users') {
    var existingHeaders = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (existingHeaders.indexOf('Role') === -1) sheet.getRange(1,sheet.getLastColumn()+1).setValue('Role').setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    if (existingHeaders.indexOf('DeviceId') === -1) sheet.getRange(1,sheet.getLastColumn()+1).setValue('DeviceId').setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
  }
  return sheet;
}

function generateId(prefix) {
  return (prefix || 'id') + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,5);
}

function hashPassword(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ('0'+(b&0xFF).toString(16)).slice(-2); }).join('');
}

function haversineMetres(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var dLat = (lat2-lat1) * Math.PI/180;
  var dLng = (lng2-lng1) * Math.PI/180;
  var a = Math.sin(dLat/2)*Math.sin(dLat/2) + Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.sin(dLng/2)*Math.sin(dLng/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function getRows(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function registerUser(body) {
  try {
    if (!body.name || !body.email || !body.password) return { success: false, message: 'Name, email and password required' };
    var sheet = getSheet(SHEET_USERS);
    var rows = getRows(sheet);
    for (var i = 0; i < rows.length; i++) if (String(rows[i].Email).toLowerCase() === String(body.email).toLowerCase()) return { success: false, message: 'Email already registered' };
    var userId = generateId('u');
    var role = body.role || 'student';
    sheet.appendRow([userId, body.name, body.email, hashPassword(body.password), body.dob || '', body.mobile || '', 'SIT Tumkur', body.department || '', 'NO', '', new Date().toISOString(), role, body.deviceId || '']);
    return { success: true, userId: userId, role: role, message: 'Account created' };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function signInUser(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var rows = getRows(sheet);
    var hash = hashPassword(body.password || '');
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r.Email).toLowerCase() === String(body.email).toLowerCase() && r.PasswordHash === hash) {
        var role = String(r.Role || '').trim() || 'student';
        return { success: true, userId: r.UserID, name: r.FullName, role: role };
      }
    }
    return { success: false, message: 'Invalid email or password' };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function markAttendance(body) {
  try {
    var lat = parseFloat(body.lat), lng = parseFloat(body.lng);
    if (isNaN(lat) || isNaN(lng)) return { success: false, message: 'Location required' };
    var dist = Math.round(haversineMetres(lat, lng, COLLEGE_LAT, COLLEGE_LNG));
    if (dist > FENCE_RADIUS_M) return { success: false, message: 'Outside geofence (' + dist + 'm)', distance: dist };
    
    var sessionSheet = getSheet(SHEET_SESSIONS);
    var sessions = getRows(sessionSheet);
    var now = new Date();
    var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var activeSession = null;
    for (var i = 0; i < sessions.length; i++) {
      var s = sessions[i];
      if (s.Status === 'open' && s.Date === todayStr) {
        var start = new Date(s.Date + 'T' + s.StartTime);
        var end = new Date(s.Date + 'T' + s.EndTime);
        if (now >= start && now <= end) { activeSession = s; break; }
      }
    }
    if (!activeSession) return { success: false, message: 'No active session' };
    
    var attSheet = getSheet(SHEET_ATTENDANCE);
    var existing = getRows(attSheet);
    for (var j = 0; j < existing.length; j++) if (existing[j].UserID === body.userId && existing[j].SessionID === activeSession.SessionID) return { success: false, message: 'Already marked for this session' };
    
    var userSheet = getSheet(SHEET_USERS);
    var users = getRows(userSheet);
    var user = null;
    for (var k = 0; k < users.length; k++) if (users[k].UserID === body.userId) { user = users[k]; break; }
    if (!user) return { success: false, message: 'User not found' };
    
    var tz = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    attSheet.appendRow([generateId('a'), body.userId, user.FullName, user.Email, activeSession.SessionID, activeSession.Subject, now.toISOString(), dateStr, timeStr, body.method || 'password', lat, lng, dist]);
    return { success: true, message: 'Attendance marked: ' + activeSession.Subject };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function createSession(body) {
  try {
    if (body.role !== 'teacher') return { success: false, message: 'Teacher only' };
    var sheet = getSheet(SHEET_SESSIONS);
    var now = new Date();
    var tz = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var startStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    var end = new Date(now.getTime() + parseInt(body.windowMinutes) * 60000);
    var endStr = Utilities.formatDate(end, tz, 'HH:mm:ss');
    var sessId = generateId('sess');
    // Close previous open sessions
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) if (data[i][1] === body.userId && data[i][7] === 'open') sheet.getRange(i+1, 8).setValue('closed');
    sheet.appendRow([sessId, body.userId, body.teacherName || '', body.subject, dateStr, startStr, endStr, 'open', body.windowMinutes]);
    return { success: true, sessionId: sessId, message: 'Session opened' };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function getActiveSession(body) {
  try {
    var sheet = getSheet(SHEET_SESSIONS);
    var rows = getRows(sheet);
    var now = new Date();
    var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    for (var i = 0; i < rows.length; i++) {
      var s = rows[i];
      if (s.Status === 'open' && s.Date === todayStr) {
        var start = new Date(s.Date + 'T' + s.StartTime);
        var end = new Date(s.Date + 'T' + s.EndTime);
        if (now >= start && now <= end) {
          var secsLeft = Math.max(0, Math.round((end - now) / 1000));
          return { success: true, active: true, session: s, secondsLeft: secsLeft };
        }
      }
    }
    return { success: true, active: false };
  } catch(err) { return { success: false, message: err.toString() }; }
}

// Simplified other functions...
function getSessions(body) {
  try {
    var sheet = getSheet(SHEET_SESSIONS);
    var rows = getRows(sheet).filter(r => r.TeacherID === body.userId);
    return { success: true, sessions: rows };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function getAttendance(body) {
  try {
    var sheet = getSheet(SHEET_ATTENDANCE);
    var rows = getRows(sheet).filter(r => r.SessionID === body.sessionId);
    return { success: true, attendance: rows };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function saveBiometric(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var bioCol = headers.indexOf('BiometricCredentialId') + 1;
    for (var i = 1; i < data.length; i++) if (data[i][0] === body.userId) {
      sheet.getRange(i+1, bioCol).setValue(body.credentialId);
      return { success: true };
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function getBiometric(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var rows = getRows(sheet);
    for (var i = 0; i < rows.length; i++) if (String(rows[i].Email).toLowerCase() === String(body.email).toLowerCase()) {
      if (!rows[i].BiometricCredentialId) return { success: false, message: 'No biometric' };
      return { success: true, credentialId: rows[i].BiometricCredentialId, userId: rows[i].UserID };
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function checkWifi(body) {
  try {
    if (!body.ssid) return { success: true, onCampusWifi: false };
    var clientSsid = String(body.ssid).trim().toLowerCase();
    for (var i = 0; i < CAMPUS_SSIDS.length; i++) if (CAMPUS_SSIDS[i].toLowerCase() === clientSsid) return { success: true, onCampusWifi: true };
    return { success: true, onCampusWifi: false, message: 'Not campus WiFi' };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function debugInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return { success: true, spreadsheetName: ss.getName(), sheets: ss.getSheets().map(s => s.getName()) };
}
