// ============================================================
//  BioAttend – Google Apps Script Backend
//  College : Siddaganga Institute of Technology, Tumkur
//  Deploy  : Web App → Execute as Me → Access: Anyone
// ============================================================

var SHEET_USERS      = 'Users';
var SHEET_ATTENDANCE = 'Attendance';
var SHEET_SESSIONS   = 'Sessions';

// ── Geofence config (SIT Tumkur) ─────────────────────────────
var COLLEGE_LAT    = 13.3318;
var COLLEGE_LNG    = 77.1274;
var FENCE_RADIUS_M = 100; // strict – inside building only

// ── Anti-cheat config ─────────────────────────────────────────
var COOLDOWN_HOURS = 3;   // students can't re-mark within 3 hours

// Known SIT campus WiFi SSIDs (add more from your IT dept)
// The client sends the SSID; server validates it's a campus network.
// Note: SSID check is a soft layer — spoofing SSID is possible but
// combined with GPS + device binding it makes proxy attendance very hard.
var CAMPUS_SSIDS = [
  'SIT-WiFi',
  'SIT_Campus',
  'SIT-Student',
  'SIT-Staff',
  'Siddaganga',
  'SIT_Tumkur',
  'sit-wifi',
  'sit_campus'
];

// ── JSON output ───────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Entry points ──────────────────────────────────────────────
function doGet(e) {
  try {
    var param = (e && e.parameter) ? e.parameter : {};
    if (!param.data) return jsonOut({ status: 'BioAttend API running', time: new Date().toString() });
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

// ── Router ────────────────────────────────────────────────────
function route(body) {
  switch (body.action) {
    case 'register':        return registerUser(body);
    case 'signIn':          return signInUser(body);
    case 'markAttendance':  return markAttendance(body);
    case 'saveBiometric':   return saveBiometric(body);
    case 'getBiometric':    return getBiometric(body);
    // Session management (teacher)
    case 'createSession':   return createSession(body);
    case 'closeSession':    return closeSession(body);
    case 'getActiveSession':return getActiveSession(body);
    case 'getSessions':     return getSessions(body);
    // Admin / reports
    case 'getAttendance':   return getAttendance(body);
    case 'getStudents':     return getStudents(body);
    case 'registerDevice':  return registerDevice(body);
    case 'checkDevice':     return checkDevice(body);
    case 'checkWifi':       return checkWifi(body);
    case 'debug':           return debugInfo();
    default: return { success: false, message: 'Unknown action: ' + body.action };
  }
}

// ── Helpers ───────────────────────────────────────────────────
function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var headers = {
      Users:      ['UserID','FullName','Email','PasswordHash','DOB','Mobile','Institution','Department','Role','BiometricCredentialId','DeviceId','CreatedAt'],
      Attendance: ['AttendanceID','UserID','FullName','Email','SessionID','Subject','Timestamp','Date','Time','Method','Lat','Lng','DistanceFromCollege'],
      Sessions:   ['SessionID','TeacherID','TeacherName','Subject','Date','StartTime','EndTime','Status','WindowMinutes']
    };
    if (headers[name]) {
      var h = headers[name];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    }
  }
  return sheet;
}

function generateId(prefix) {
  return (prefix || 'id') + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,5);
}

function hashPassword(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
}

// Haversine distance in metres between two lat/lng points
function haversineMetres(lat1, lng1, lat2, lng2) {
  var R  = 6371000;
  var dL = (lat2 - lat1) * Math.PI / 180;
  var dN = (lng2 - lng1) * Math.PI / 180;
  var a  = Math.sin(dL/2)*Math.sin(dL/2) +
           Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*
           Math.sin(dN/2)*Math.sin(dN/2);
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

// ── 1. Register ───────────────────────────────────────────────
function registerUser(body) {
  try {
    if (!body.name || !body.email || !body.password)
      return { success: false, message: 'Name, email and password are required' };

    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].Email).toLowerCase() === String(body.email).toLowerCase())
        return { success: false, message: 'Email already registered' };
    }

    var userId = generateId('u');
    var role   = body.role || 'student'; // 'student' or 'teacher'
    sheet.appendRow([
      userId, body.name, body.email, hashPassword(body.password),
      body.dob || '', body.mobile || '', 'SIT Tumkur',
      body.department || '', role, '', body.deviceId || '',
      new Date().toISOString()
    ]);
    return { success: true, userId: userId, role: role, message: 'Account created' };
  } catch(err) { return { success: false, message: 'register error: ' + err.toString() }; }
}

// ── 2. Sign In ────────────────────────────────────────────────
function signInUser(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet);
    var hash  = hashPassword(body.password || '');
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r.Email).toLowerCase() === String(body.email).toLowerCase() && r.PasswordHash === hash) {
        return { success: true, userId: r.UserID, name: r.FullName, role: r.Role || 'student' };
      }
    }
    return { success: false, message: 'Invalid email or password' };
  } catch(err) { return { success: false, message: 'signIn error: ' + err.toString() }; }
}

// ── 3. Mark Attendance (with geofence + session check) ────────
function markAttendance(body) {
  try {
    // ── A. Geofence check ──
    var lat = parseFloat(body.lat);
    var lng = parseFloat(body.lng);
    if (isNaN(lat) || isNaN(lng))
      return { success: false, message: 'Location not provided. Enable GPS and try again.' };

    var dist = Math.round(haversineMetres(lat, lng, COLLEGE_LAT, COLLEGE_LNG));
    if (dist > FENCE_RADIUS_M)
      return {
        success: false,
        message: 'You are ' + dist + 'm away from SIT campus. You must be within ' + FENCE_RADIUS_M + 'm to mark attendance.',
        distance: dist
      };

    // ── B. Device binding check ──
    if (body.deviceId) {
      var userSheet2 = getSheet(SHEET_USERS);
      var users2     = getRows(userSheet2);
      for (var di = 0; di < users2.length; di++) {
        if (users2[di].UserID === body.userId) {
          var storedDevice = String(users2[di].DeviceId || '').trim();
          if (storedDevice && storedDevice !== String(body.deviceId).trim()) {
            return {
              success: false,
              message: 'This account is bound to a different device. Attendance can only be marked from your registered phone.'
            };
          }
          break;
        }
      }
    }

    // ── C. WiFi SSID check (soft layer — warns if not on campus network) ──
    if (body.ssid) {
      var ssidMatch = false;
      var clientSsid = String(body.ssid).trim().toLowerCase();
      for (var si = 0; si < CAMPUS_SSIDS.length; si++) {
        if (CAMPUS_SSIDS[si].toLowerCase() === clientSsid) { ssidMatch = true; break; }
      }
      if (!ssidMatch) {
        return {
          success: false,
          code: 'WIFI_MISMATCH',
          message: 'You must be connected to SIT campus WiFi to mark attendance. Current network: "' + body.ssid + '"'
        };
      }
    }

    // ── D. Active session check ──
    var sessionSheet = getSheet(SHEET_SESSIONS);
    var sessions     = getRows(sessionSheet);
    var now          = new Date();
    var todayStr     = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var activeSession = null;

    for (var i = 0; i < sessions.length; i++) {
      var s = sessions[i];
      if (s.Status === 'open' && String(s.Date) === todayStr) {
        // Check time window
        var start = new Date(s.Date + 'T' + s.StartTime);
        var end   = new Date(s.Date + 'T' + s.EndTime);
        if (now >= start && now <= end) { activeSession = s; break; }
      }
    }
    if (!activeSession)
      return { success: false, message: 'No active attendance session right now. Wait for your teacher to open one.' };

    // ── E. Duplicate check (same session) ──
    var attSheet = getSheet(SHEET_ATTENDANCE);
    var existing = getRows(attSheet);
    for (var j = 0; j < existing.length; j++) {
      var a = existing[j];
      if (a.UserID === body.userId && a.SessionID === activeSession.SessionID)
        return { success: false, message: 'You have already marked attendance for this session.' };
    }

    // ── F. Cooldown check ──
    var COOLDOWN_MS = COOLDOWN_HOURS * 60 * 60 * 1000;
    for (var ci = 0; ci < existing.length; ci++) {
      var ca = existing[ci];
      if (ca.UserID === body.userId && ca.Timestamp) {
        var lastTime = new Date(ca.Timestamp);
        var diffMs   = now.getTime() - lastTime.getTime();
        if (diffMs < COOLDOWN_MS) {
          var minsLeft = Math.ceil((COOLDOWN_MS - diffMs) / 60000);
          return {
            success: false,
            code: 'COOLDOWN',
            minutesLeft: minsLeft,
            message: 'Cooldown active. You can mark attendance again in ' + minsLeft + ' minute(s).'
          };
        }
      }
    }

    // ── G. Get user info ──
    var userSheet = getSheet(SHEET_USERS);
    var users     = getRows(userSheet);
    var user      = null;
    for (var k = 0; k < users.length; k++) {
      if (users[k].UserID === body.userId) { user = users[k]; break; }
    }
    if (!user) return { success: false, message: 'User not found' };

    // ── H. Record ──
    var tz      = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    attSheet.appendRow([
      generateId('a'), body.userId, user.FullName, user.Email,
      activeSession.SessionID, activeSession.Subject,
      now.toISOString(), dateStr, timeStr,
      body.method || 'password',
      lat, lng, dist
    ]);

    return {
      success: true,
      message: '✓ Attendance marked for ' + activeSession.Subject + ' at ' + timeStr,
      subject: activeSession.Subject,
      distance: dist
    };
  } catch(err) { return { success: false, message: 'markAttendance error: ' + err.toString() }; }
}

// ── 4. Create Session (teacher) ───────────────────────────────
function createSession(body) {
  try {
    if (body.role !== 'teacher' && body.role !== 'admin')
      return { success: false, message: 'Only teachers can create sessions' };
    if (!body.subject || !body.windowMinutes)
      return { success: false, message: 'Subject and window duration required' };

    var sheet    = getSheet(SHEET_SESSIONS);
    var now      = new Date();
    var tz       = Session.getScriptTimeZone();
    var dateStr  = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var startStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    var end      = new Date(now.getTime() + parseInt(body.windowMinutes) * 60000);
    var endStr   = Utilities.formatDate(end, tz, 'HH:mm:ss');
    var sessId   = generateId('sess');

    // Close any other open sessions by this teacher
    var rows = getRows(sheet);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === body.userId && data[i][7] === 'open') {
        sheet.getRange(i+1, 8).setValue('closed');
      }
    }

    sheet.appendRow([sessId, body.userId, body.teacherName || '', body.subject, dateStr, startStr, endStr, 'open', body.windowMinutes]);
    return { success: true, sessionId: sessId, subject: body.subject, startTime: startStr, endTime: endStr, message: 'Session opened for ' + body.windowMinutes + ' minutes' };
  } catch(err) { return { success: false, message: 'createSession error: ' + err.toString() }; }
}

// ── 5. Close Session (teacher) ────────────────────────────────
function closeSession(body) {
  try {
    var sheet = getSheet(SHEET_SESSIONS);
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.sessionId) {
        sheet.getRange(i+1, 8).setValue('closed');
        return { success: true, message: 'Session closed' };
      }
    }
    return { success: false, message: 'Session not found' };
  } catch(err) { return { success: false, message: 'closeSession error: ' + err.toString() }; }
}

// ── 6. Get Active Session ─────────────────────────────────────
function getActiveSession(body) {
  try {
    var sheet    = getSheet(SHEET_SESSIONS);
    var rows     = getRows(sheet);
    var now      = new Date();
    var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    for (var i = 0; i < rows.length; i++) {
      var s = rows[i];
      if (s.Status === 'open' && String(s.Date) === todayStr) {
        var start = new Date(s.Date + 'T' + s.StartTime);
        var end   = new Date(s.Date + 'T' + s.EndTime);
        if (now >= start && now <= end) {
          var secsLeft = Math.max(0, Math.round((end - now) / 1000));
          return { success: true, active: true, session: s, secondsLeft: secsLeft };
        }
      }
    }
    return { success: true, active: false };
  } catch(err) { return { success: false, message: 'getActiveSession error: ' + err.toString() }; }
}

// ── 7. Get Sessions list (teacher dashboard) ──────────────────
function getSessions(body) {
  try {
    var sheet = getSheet(SHEET_SESSIONS);
    var rows  = getRows(sheet);
    // Filter by teacher if provided
    if (body.userId) rows = rows.filter(function(r){ return r.TeacherID === body.userId; });
    // Get attendance count per session
    var attSheet = getSheet(SHEET_ATTENDANCE);
    var attRows  = getRows(attSheet);
    rows.forEach(function(s) {
      s.presentCount = attRows.filter(function(a){ return a.SessionID === s.SessionID; }).length;
    });
    rows.reverse(); // newest first
    return { success: true, sessions: rows.slice(0, 20) };
  } catch(err) { return { success: false, message: 'getSessions error: ' + err.toString() }; }
}

// ── 8. Get Attendance list (teacher dashboard) ────────────────
function getAttendance(body) {
  try {
    var sheet = getSheet(SHEET_ATTENDANCE);
    var rows  = getRows(sheet);
    if (body.sessionId) rows = rows.filter(function(r){ return r.SessionID === body.sessionId; });
    if (body.date)      rows = rows.filter(function(r){ return r.Date === body.date; });
    return { success: true, attendance: rows };
  } catch(err) { return { success: false, message: 'getAttendance error: ' + err.toString() }; }
}

// ── 9. Get Students list ──────────────────────────────────────
function getStudents(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet).filter(function(r){ return r.Role === 'student'; });
    // strip password hash for safety
    rows.forEach(function(r){ delete r.PasswordHash; delete r.BiometricCredentialId; });
    return { success: true, students: rows };
  } catch(err) { return { success: false, message: 'getStudents error: ' + err.toString() }; }
}

// ── 10. Biometric helpers ─────────────────────────────────────
function saveBiometric(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.userId) {
        sheet.getRange(i+1, 10).setValue(body.credentialId);
        return { success: true };
      }
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: 'saveBiometric error: ' + err.toString() }; }
}

function getBiometric(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet);
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r.Email).toLowerCase() === String(body.email).toLowerCase()) {
        if (!r.BiometricCredentialId) return { success: false, message: 'No biometric registered' };
        return { success: true, credentialId: r.BiometricCredentialId, userId: r.UserID };
      }
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: 'getBiometric error: ' + err.toString() }; }
}

// ── 12. Register Device ───────────────────────────────────────
// Called after account creation. Saves the device fingerprint to col 11.
// If already bound to a different device, blocks the change.
function registerDevice(body) {
  try {
    if (!body.userId || !body.deviceId)
      return { success: false, message: 'userId and deviceId required' };

    var sheet = getSheet(SHEET_USERS);
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.userId) {
        var existing = String(data[i][10] || '').trim(); // col 11 = DeviceId (0-indexed: 10)

        // Already bound to a DIFFERENT device → block
        if (existing && existing !== String(body.deviceId).trim()) {
          return {
            success: false,
            alreadyBound: true,
            message: 'This account is already bound to another device. Contact your admin to reset it.'
          };
        }

        // Not yet bound OR same device → save/confirm
        sheet.getRange(i + 1, 11).setValue(body.deviceId);
        return {
          success: true,
          alreadyBound: false,
          message: existing ? 'Device confirmed' : 'Device registered successfully'
        };
      }
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: 'registerDevice error: ' + err.toString() }; }
}

// ── 13. Check Device (lightweight, called on sign-in) ─────────
function checkDevice(body) {
  try {
    if (!body.userId || !body.deviceId)
      return { success: false, message: 'userId and deviceId required' };

    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet);

    for (var i = 0; i < rows.length; i++) {
      if (rows[i].UserID === body.userId) {
        var stored = String(rows[i].DeviceId || '').trim();
        if (!stored) return { success: true, status: 'unbound' };
        if (stored === String(body.deviceId).trim()) return { success: true, status: 'match' };
        return { success: false, status: 'mismatch', message: 'This account is registered to a different device.' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: 'checkDevice error: ' + err.toString() }; }
}

// ── 15. WiFi SSID check (called from frontend before markAttendance) ──
// Returns whether the given SSID is a known SIT campus network.
function checkWifi(body) {
  try {
    if (!body.ssid) return { success: true, onCampusWifi: false, message: 'No SSID provided' };
    var clientSsid = String(body.ssid).trim().toLowerCase();
    for (var i = 0; i < CAMPUS_SSIDS.length; i++) {
      if (CAMPUS_SSIDS[i].toLowerCase() === clientSsid)
        return { success: true, onCampusWifi: true, ssid: body.ssid };
    }
    return {
      success: true,
      onCampusWifi: false,
      ssid: body.ssid,
      message: 'Not on a recognised SIT campus network. Connect to SIT WiFi and try again.'
    };
  } catch(err) { return { success: false, message: 'checkWifi error: ' + err.toString() }; }
}

// ── 14. Debug ─────────────────────────────────────────────────
function debugInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return { success: true, spreadsheetName: ss.getName(), sheets: ss.getSheets().map(function(s){ return s.getName(); }), id: ss.getId() };
  } catch(err) { return { success: false, message: 'Spreadsheet error: ' + err.toString() }; }
}