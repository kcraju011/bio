// ============================================================
//  BioAttend – Google Apps Script Backend
//  College : Siddaganga Institute of Technology, Tumkur
//  Deploy  : Web App → Execute as Me → Access: Anyone
//  Version : 4 (fixed column mapping, no setHeader, dev mode)
// ============================================================

// ── Sheet names ───────────────────────────────────────────────
var SHEET_USERS      = 'Users';
var SHEET_ATTENDANCE = 'Attendance';
var SHEET_SESSIONS   = 'Sessions';

// ── Location config ───────────────────────────────────────────
// DEVELOPMENT MODE: using home location for testing
// Change to SIT coords (13.3318, 77.1274) when deploying for real
var COLLEGE_LAT    = 13.3280233;
var COLLEGE_LNG    = 77.1198344;
var FENCE_RADIUS_M = 500; // relaxed for home testing — set to 100 for production

// ── Anti-cheat config ─────────────────────────────────────────
var COOLDOWN_HOURS = 3;

// Allowed WiFi SSIDs (dev: home network included for testing)
var CAMPUS_SSIDS = [
  'Airtel_Vodka',   // dev/home — REMOVE before production
  'SIT-WiFi',
  'SIT_Campus',
  'SIT-Student',
  'SIT-Staff',
  'Siddaganga',
  'SIT_Tumkur',
  'sit-wifi',
  'sit_campus'
];

// ── DEV MODE: set true to skip WiFi + geofence checks ────────
var DEV_MODE = true; // CHANGE TO false before production deployment

// ── JSON output (no setHeader — not supported in Apps Script) ─
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Entry points ──────────────────────────────────────────────
function doGet(e) {
  try {
    var param = (e && e.parameter) ? e.parameter : {};
    if (!param.data) return jsonOut({ status: 'BioAttend API v4 running', time: new Date().toString(), devMode: DEV_MODE });
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
    case 'register':         return registerUser(body);
    case 'signIn':           return signInUser(body);
    case 'markAttendance':   return markAttendance(body);
    case 'saveBiometric':    return saveBiometric(body);
    case 'getBiometric':     return getBiometric(body);
    case 'createSession':    return createSession(body);
    case 'closeSession':     return closeSession(body);
    case 'getActiveSession': return getActiveSession(body);
    case 'getSessions':      return getSessions(body);
    case 'getAttendance':    return getAttendance(body);
    case 'getStudents':      return getStudents(body);
    case 'getDashboard':     return getDashboard(body);
    case 'exportAttendance': return exportAttendance(body);
    case 'registerDevice':   return registerDevice(body);
    case 'checkDevice':      return checkDevice(body);
    case 'checkWifi':        return checkWifi(body);
    case 'fixRows':          return fixExistingRows();
    case 'resetUsersSheet':  return resetUsersSheet();
    case 'debug':            return debugInfo();
    default: return { success: false, message: 'Unknown action: ' + body.action };
  }
}

// ── Sheet helpers ─────────────────────────────────────────────
// Clean 9-column Users sheet — no MarkFromAnywhere, no DeviceId confusion.
// DeviceId is stored by registerDevice() dynamically after registration.
// Col:  1        2          3       4              5     6        7              8             9
// Name: UserID   FullName   Email   PasswordHash   DOB   Mobile   Institution   Department   Role
var USER_HEADERS = ['UserID','FullName','Email','PasswordHash','DOB','Mobile','Institution','Department','Role'];

function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var initHeaders = {
      Users:      USER_HEADERS,
      Attendance: ['AttendanceID','UserID','FullName','Email','SessionID','Subject','Timestamp','Date','Time','Method','Lat','Lng','DistanceFromCollege'],
      Sessions:   ['SessionID','TeacherID','TeacherName','Subject','Date','StartTime','EndTime','Status','WindowMinutes']
    };
    if (initHeaders[name]) {
      var h = initHeaders[name];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    }
  } else if (name === SHEET_USERS) {
    // Auto-add missing columns to existing sheets without touching existing data
    var existing = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    ['BiometricCredentialId','DeviceId'].forEach(function(col) {
      if (existing.indexOf(col) === -1) {
        var c = sheet.getLastColumn() + 1;
        sheet.getRange(1,c).setValue(col).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
      }
    });
  }
  return sheet;
}

// Maps rows to objects using the ACTUAL header row (immune to column order)
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

// Finds 1-indexed column number by header name
function colIndex(sheet, headerName) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var idx = headers.indexOf(headerName);
  return idx === -1 ? -1 : idx + 1;
}

function generateId(prefix) {
  return (prefix||'id') + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,5);
}

function hashPassword(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
}

function haversineMetres(lat1,lng1,lat2,lng2) {
  var R=6371000, dL=(lat2-lat1)*Math.PI/180, dN=(lng2-lng1)*Math.PI/180;
  var a=Math.sin(dL/2)*Math.sin(dL/2)+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dN/2)*Math.sin(dN/2);
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

// ── 1. Register ───────────────────────────────────────────────
function registerUser(body) {
  try {
    if (!body.name || !body.email || !body.password)
      return { success:false, message:'Name, email and password required' };

    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet);
    for (var i=0; i<rows.length; i++) {
      if (String(rows[i].Email).toLowerCase() === String(body.email).toLowerCase())
        return { success:false, message:'Email already registered' };
    }

    var userId = generateId('u');
    var role   = body.role || 'student';

    // Write exactly 9 columns matching USER_HEADERS:
    // UserID, FullName, Email, PasswordHash, DOB, Mobile, Institution, Department, Role
    // BiometricCredentialId and DeviceId are added as extra columns by getSheet()
    // and written separately by saveBiometric() and registerDevice()
    sheet.appendRow([
      userId,                       // 1 UserID
      body.name,                    // 2 FullName
      body.email,                   // 3 Email
      hashPassword(body.password),  // 4 PasswordHash
      body.dob     || '',           // 5 DOB
      body.mobile  || '',           // 6 Mobile
      'SIT Tumkur',                 // 7 Institution
      body.department || '',        // 8 Department
      role                          // 9 Role
    ]);
    return { success:true, userId:userId, role:role, message:'Account created' };
  } catch(err) { return { success:false, message:'register: '+err.toString() }; }
}

// ── 2. Sign In ────────────────────────────────────────────────
function signInUser(body) {
  try {
    var sheet = getSheet(SHEET_USERS);
    var rows  = getRows(sheet);
    var hash  = hashPassword(body.password || '');
    for (var i=0; i<rows.length; i++) {
      var r = rows[i];
      if (String(r.Email).toLowerCase() === String(body.email).toLowerCase() && r.PasswordHash === hash) {
        var role = String(r.Role || '').trim() || 'student';
        return { success:true, userId:r.UserID, name:r.FullName, role:role };
      }
    }
    return { success:false, message:'Invalid email or password' };
  } catch(err) { return { success:false, message:'signIn: '+err.toString() }; }
}

// ── 3. Mark Attendance ────────────────────────────────────────
function markAttendance(body) {
  try {
    var now = new Date();
    var tz  = Session.getScriptTimeZone();

    // A. Geofence (skip in DEV_MODE)
    var dist = 0;
    if (!DEV_MODE) {
      var lat = parseFloat(body.lat), lng = parseFloat(body.lng);
      if (isNaN(lat)||isNaN(lng)) return { success:false, message:'GPS required. Enable location and try again.' };
      dist = Math.round(haversineMetres(lat, lng, COLLEGE_LAT, COLLEGE_LNG));
      if (dist > FENCE_RADIUS_M)
        return { success:false, message:'You are '+dist+'m from campus. Must be within '+FENCE_RADIUS_M+'m.', distance:dist };
    } else {
      // DEV_MODE: still compute distance for logging but don't block
      var lat = parseFloat(body.lat)||COLLEGE_LAT, lng = parseFloat(body.lng)||COLLEGE_LNG;
      dist = Math.round(haversineMetres(lat, lng, COLLEGE_LAT, COLLEGE_LNG));
    }

    // B. Device binding (skip in DEV_MODE)
    if (!DEV_MODE && body.deviceId) {
      var uRows = getRows(getSheet(SHEET_USERS));
      for (var di=0; di<uRows.length; di++) {
        if (uRows[di].UserID === body.userId) {
          var sd = String(uRows[di].DeviceId||'').trim();
          if (sd && sd !== String(body.deviceId).trim())
            return { success:false, message:'Account bound to a different device.' };
          break;
        }
      }
    }

    // C. WiFi check (skip in DEV_MODE)
    if (!DEV_MODE && body.ssid) {
      var ssidOk = false, cs = String(body.ssid).trim().toLowerCase();
      for (var si=0; si<CAMPUS_SSIDS.length; si++) {
        if (CAMPUS_SSIDS[si].toLowerCase()===cs){ ssidOk=true; break; }
      }
      if (!ssidOk) return { success:false, code:'WIFI_MISMATCH', message:'Connect to SIT campus WiFi. Current: "'+body.ssid+'"' };
    }

    // D. Active session check
    var sessions = getRows(getSheet(SHEET_SESSIONS));
    var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var activeSession = null;
    for (var i=0; i<sessions.length; i++) {
      var s=sessions[i];
      if (s.Status==='open' && String(s.Date)===todayStr) {
        var st=new Date(s.Date+'T'+s.StartTime), en=new Date(s.Date+'T'+s.EndTime);
        if (now>=st && now<=en){ activeSession=s; break; }
      }
    }
    if (!activeSession) return { success:false, message:'No active session right now. Wait for your teacher to open one.' };

    // E. Duplicate check
    var attSheet = getSheet(SHEET_ATTENDANCE);
    var existing = getRows(attSheet);
    for (var j=0; j<existing.length; j++) {
      if (existing[j].UserID===body.userId && existing[j].SessionID===activeSession.SessionID)
        return { success:false, message:'Already marked for this session.' };
    }

    // F. Cooldown check
    var coolMs = COOLDOWN_HOURS*3600000;
    for (var ci=0; ci<existing.length; ci++) {
      if (existing[ci].UserID===body.userId && existing[ci].Timestamp) {
        var diff = now.getTime() - new Date(existing[ci].Timestamp).getTime();
        if (diff < coolMs) {
          var ml = Math.ceil((coolMs-diff)/60000);
          return { success:false, code:'COOLDOWN', minutesLeft:ml, message:'Cooldown active. Try again in '+ml+' min.' };
        }
      }
    }

    // G. Get user
    var users = getRows(getSheet(SHEET_USERS));
    var user  = null;
    for (var k=0; k<users.length; k++) { if (users[k].UserID===body.userId){ user=users[k]; break; } }
    if (!user) return { success:false, message:'User not found' };

    // H. Record
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    attSheet.appendRow([
      generateId('a'), body.userId, user.FullName, user.Email,
      activeSession.SessionID, activeSession.Subject,
      now.toISOString(), dateStr, timeStr,
      body.method||'password', lat, lng, dist
    ]);
    return {
      success:true,
      message:(DEV_MODE?'[DEV] ':'')+'\u2713 Attendance marked for '+activeSession.Subject+' at '+timeStr,
      subject:activeSession.Subject, distance:dist
    };
  } catch(err) { return { success:false, message:'markAttendance: '+err.toString() }; }
}

// ── 4. Create Session ─────────────────────────────────────────
function createSession(body) {
  try {
    if (body.role!=='teacher'&&body.role!=='admin')
      return { success:false, message:'Only teachers can create sessions' };
    if (!body.subject||!body.windowMinutes)
      return { success:false, message:'Subject and window required' };

    var sheet=getSheet(SHEET_SESSIONS), now=new Date(), tz=Session.getScriptTimeZone();
    var dateStr=Utilities.formatDate(now,tz,'yyyy-MM-dd');
    var startStr=Utilities.formatDate(now,tz,'HH:mm:ss');
    var end=new Date(now.getTime()+parseInt(body.windowMinutes)*60000);
    var endStr=Utilities.formatDate(end,tz,'HH:mm:ss');
    var sessId=generateId('sess');

    // Close existing open sessions by this teacher
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++){
      if (data[i][1]===body.userId&&data[i][7]==='open')
        sheet.getRange(i+1,8).setValue('closed');
    }
    sheet.appendRow([sessId,body.userId,body.teacherName||'',body.subject,dateStr,startStr,endStr,'open',body.windowMinutes]);
    return { success:true, sessionId:sessId, subject:body.subject, startTime:startStr, endTime:endStr, message:'Session opened for '+body.windowMinutes+' min' };
  } catch(err) { return { success:false, message:'createSession: '+err.toString() }; }
}

// ── 5. Close Session ──────────────────────────────────────────
function closeSession(body) {
  try {
    var sheet=getSheet(SHEET_SESSIONS), data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++){
      if (data[i][0]===body.sessionId){ sheet.getRange(i+1,8).setValue('closed'); return { success:true }; }
    }
    return { success:false, message:'Session not found' };
  } catch(err) { return { success:false, message:'closeSession: '+err.toString() }; }
}

// ── 6. Get Active Session ─────────────────────────────────────
function getActiveSession(body) {
  try {
    var sheet=getSheet(SHEET_SESSIONS), rows=getRows(sheet), now=new Date();
    var todayStr=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd');
    for (var i=0;i<rows.length;i++){
      var s=rows[i];
      if (s.Status==='open'&&String(s.Date)===todayStr){
        var st=new Date(s.Date+'T'+s.StartTime), en=new Date(s.Date+'T'+s.EndTime);
        if (now>=st&&now<=en){
          return { success:true, active:true, session:s, secondsLeft:Math.max(0,Math.round((en-now)/1000)) };
        }
      }
    }
    return { success:true, active:false };
  } catch(err) { return { success:false, message:'getActiveSession: '+err.toString() }; }
}

// ── 7. Get Sessions list ──────────────────────────────────────
function getSessions(body) {
  try {
    var sheet=getSheet(SHEET_SESSIONS), rows=getRows(sheet);
    if (body.userId) rows=rows.filter(function(r){ return r.TeacherID===body.userId; });
    var attRows=getRows(getSheet(SHEET_ATTENDANCE));
    rows.forEach(function(s){ s.presentCount=attRows.filter(function(a){ return a.SessionID===s.SessionID; }).length; });
    rows.reverse();
    return { success:true, sessions:rows.slice(0,20) };
  } catch(err) { return { success:false, message:'getSessions: '+err.toString() }; }
}

// ── 8. Get Attendance list ────────────────────────────────────
function getAttendance(body) {
  try {
    var rows=getRows(getSheet(SHEET_ATTENDANCE));
    if (body.sessionId) rows=rows.filter(function(r){ return r.SessionID===body.sessionId; });
    if (body.date)      rows=rows.filter(function(r){ return r.Date===body.date; });
    return { success:true, attendance:rows };
  } catch(err) { return { success:false, message:'getAttendance: '+err.toString() }; }
}

// ── 9. Get Students ───────────────────────────────────────────
function getStudents(body) {
  try {
    var rows=getRows(getSheet(SHEET_USERS)).filter(function(r){ return String(r.Role||'student')==='student'; });
    rows.forEach(function(r){ delete r.PasswordHash; delete r.BiometricCredentialId; });
    return { success:true, students:rows };
  } catch(err) { return { success:false, message:'getStudents: '+err.toString() }; }
}

// ── 9b. Get Dashboard (live stats for a session) ─────────────
// Returns present students + absent students for a session
function getDashboard(body) {
  try {
    if (!body.sessionId) return { success:false, message:'sessionId required' };

    // All students enrolled (all users with role=student)
    var allStudents = getRows(getSheet(SHEET_USERS)).filter(function(r){
      return String(r.Role||'student').trim() === 'student';
    });

    // Who marked attendance for this session
    var attRows = getRows(getSheet(SHEET_ATTENDANCE)).filter(function(r){
      return r.SessionID === body.sessionId;
    });

    var presentIds = {};
    attRows.forEach(function(a){ presentIds[a.UserID] = a; });

    var present = [];
    var absent  = [];

    allStudents.forEach(function(s) {
      if (presentIds[s.UserID]) {
        present.push({
          userId:     s.UserID,
          name:       s.FullName,
          email:      s.Email,
          department: s.Department,
          time:       presentIds[s.UserID].Time,
          method:     presentIds[s.UserID].Method,
          distance:   presentIds[s.UserID].DistanceFromCollege
        });
      } else {
        absent.push({
          userId:     s.UserID,
          name:       s.FullName,
          email:      s.Email,
          department: s.Department
        });
      }
    });

    return {
      success:      true,
      total:        allStudents.length,
      presentCount: present.length,
      absentCount:  absent.length,
      present:      present,
      absent:       absent
    };
  } catch(err) { return { success:false, message:'getDashboard: '+err.toString() }; }
}

// ── 9c. Export Attendance as CSV data ─────────────────────────
function exportAttendance(body) {
  try {
    var rows = getRows(getSheet(SHEET_ATTENDANCE));
    if (body.sessionId) rows = rows.filter(function(r){ return r.SessionID === body.sessionId; });
    if (body.date)      rows = rows.filter(function(r){ return r.Date === body.date; });
    if (body.subject)   rows = rows.filter(function(r){ return r.Subject === body.subject; });

    // Build CSV string
    var headers = ['Name','Email','Department','Date','Time','Subject','Method','Distance(m)'];
    var lines   = [headers.join(',')];

    // Get department from users sheet
    var userMap = {};
    getRows(getSheet(SHEET_USERS)).forEach(function(u){ userMap[u.UserID] = u.Department || ''; });

    rows.forEach(function(r) {
      lines.push([
        '"'+(r.FullName||'')+'"',
        '"'+(r.Email||'')+'"',
        '"'+(userMap[r.UserID]||'')+'"',
        r.Date || '',
        r.Time || '',
        '"'+(r.Subject||'')+'"',
        r.Method || '',
        r.DistanceFromCollege || ''
      ].join(','));
    });

    return { success:true, csv:lines.join('\n'), rowCount:rows.length };
  } catch(err) { return { success:false, message:'exportAttendance: '+err.toString() }; }
}
function saveBiometric(body) {
  try {
    var sheet=getSheet(SHEET_USERS), data=sheet.getDataRange().getValues(), headers=data[0];
    var col=headers.indexOf('BiometricCredentialId')+1;
    if (col<1) return { success:false, message:'BiometricCredentialId column missing' };
    for (var i=1;i<data.length;i++){
      if (data[i][0]===body.userId){ sheet.getRange(i+1,col).setValue(body.credentialId); return { success:true }; }
    }
    return { success:false, message:'User not found' };
  } catch(err) { return { success:false, message:'saveBiometric: '+err.toString() }; }
}

// ── 11. Get Biometric ─────────────────────────────────────────
function getBiometric(body) {
  try {
    var rows=getRows(getSheet(SHEET_USERS));
    for (var i=0;i<rows.length;i++){
      var r=rows[i];
      if (String(r.Email).toLowerCase()===String(body.email).toLowerCase()){
        if (!r.BiometricCredentialId) return { success:false, message:'No biometric registered' };
        return { success:true, credentialId:r.BiometricCredentialId, userId:r.UserID };
      }
    }
    return { success:false, message:'User not found' };
  } catch(err) { return { success:false, message:'getBiometric: '+err.toString() }; }
}

// ── 12. Register Device ───────────────────────────────────────
function registerDevice(body) {
  try {
    if (!body.userId||!body.deviceId) return { success:false, message:'userId and deviceId required' };
    var sheet=getSheet(SHEET_USERS), data=sheet.getDataRange().getValues(), headers=data[0];
    var col=headers.indexOf('DeviceId')+1;
    if (col<1) return { success:false, message:'DeviceId column not found' };
    for (var i=1;i<data.length;i++){
      if (data[i][0]===body.userId){
        var existing=String(data[i][col-1]||'').trim();
        if (existing && existing!==String(body.deviceId).trim())
          return { success:false, alreadyBound:true, message:'Account already bound to another device.' };
        sheet.getRange(i+1,col).setValue(body.deviceId);
        return { success:true, alreadyBound:false, message:existing?'Device confirmed':'Device registered' };
      }
    }
    return { success:false, message:'User not found' };
  } catch(err) { return { success:false, message:'registerDevice: '+err.toString() }; }
}

// ── 13. Check Device ─────────────────────────────────────────
function checkDevice(body) {
  try {
    if (!body.userId||!body.deviceId) return { success:false, message:'userId and deviceId required' };
    var rows=getRows(getSheet(SHEET_USERS));
    for (var i=0;i<rows.length;i++){
      if (rows[i].UserID===body.userId){
        var stored=String(rows[i].DeviceId||'').trim();
        if (!stored) return { success:true, status:'unbound' };
        if (stored===String(body.deviceId).trim()) return { success:true, status:'match' };
        return { success:false, status:'mismatch', message:'Account registered to a different device.' };
      }
    }
    return { success:false, message:'User not found' };
  } catch(err) { return { success:false, message:'checkDevice: '+err.toString() }; }
}

// ── 14. Check WiFi ────────────────────────────────────────────
function checkWifi(body) {
  try {
    if (!body.ssid) return { success:true, onCampusWifi:false, message:'No SSID provided' };
    var cs=String(body.ssid).trim().toLowerCase();
    for (var i=0;i<CAMPUS_SSIDS.length;i++){
      if (CAMPUS_SSIDS[i].toLowerCase()===cs) return { success:true, onCampusWifi:true, ssid:body.ssid };
    }
    return { success:true, onCampusWifi:false, ssid:body.ssid, message:'Not on a recognised campus network.' };
  } catch(err) { return { success:false, message:'checkWifi: '+err.toString() }; }
}

// ── 15. Reset Users Sheet (run once from editor to fix corruption) ──
// Deletes the old Users sheet and creates a fresh clean one.
// ALL existing user rows are deleted — re-register after running this.
function resetUsersSheet() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var old   = ss.getSheetByName(SHEET_USERS);
    if (old) ss.deleteSheet(old);
    // Re-create with clean headers
    var sheet = ss.insertSheet(SHEET_USERS);
    var h = USER_HEADERS.concat(['BiometricCredentialId','DeviceId']);
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    return { success:true, message:'Users sheet reset. Headers: ' + h.join(', ') + '. Please re-register all users.' };
  } catch(err) { return { success:false, message:'resetUsersSheet: '+err.toString() }; }
}

// ── 16. Fix existing bad rows (repair without deleting) ───────
function fixExistingRows() {
  try {
    var sheet   = getSheet(SHEET_USERS);
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var roleCol = headers.indexOf('Role')+1;
    var mfaCol  = headers.indexOf('MarkFromAnywhere')+1;
    if (roleCol < 1) return { success:false, message:'Role column not found' };
    var fixed = 0;
    for (var i=1; i<data.length; i++) {
      var role = String(data[i][roleCol-1]||'').trim();
      var mfa  = mfaCol>0 ? String(data[i][mfaCol-1]||'').trim() : '';
      // If Role is blank but MarkFromAnywhere has a role value
      if (!role && (mfa==='teacher'||mfa==='student')) {
        sheet.getRange(i+1,roleCol).setValue(mfa);
        if (mfaCol>0) sheet.getRange(i+1,mfaCol).setValue('NO');
        fixed++;
      }
      // If Role is blank entirely
      if (!role && mfa!=='teacher' && mfa!=='student') {
        sheet.getRange(i+1,roleCol).setValue('student');
        fixed++;
      }
    }
    return { success:true, message:'Fixed '+fixed+' rows' };
  } catch(err) { return { success:false, message:'fixExistingRows: '+err.toString() }; }
}

// ── 16. Debug ─────────────────────────────────────────────────
function debugInfo() {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var userSheet=getSheet(SHEET_USERS);
    var headers=userSheet.getRange(1,1,1,userSheet.getLastColumn()).getValues()[0];
    return {
      success:true,
      spreadsheetName:ss.getName(),
      sheets:ss.getSheets().map(function(s){ return s.getName(); }),
      userHeaders:headers,
      devMode:DEV_MODE,
      college:{ lat:COLLEGE_LAT, lng:COLLEGE_LNG, fenceM:FENCE_RADIUS_M }
    };
  } catch(err) { return { success:false, message:'debug: '+err.toString() }; }
}