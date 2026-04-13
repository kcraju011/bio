// ============================================================
//  BioAttend v5 – Google Apps Script Backend
//  College : Siddaganga Institute of Technology, Tumkur
//  Schema  : Normalized, production-grade, optimized for Sheets
//  Deploy  : Web App → Execute as Me → Access: Anyone
// ============================================================
//
//  SHEET STRUCTURE (normalized schema):
//  ┌──────────────────────────────────────────────────────┐
//  │  Roles         | role_id, name                       │
//  │  Departments   | dept_id, name, incharge, email      │
//  │  Locations     | loc_id, name, lat, lng              │
//  │  Users         | user_id, inst_id, dept_id, role_id, │
//  │                │ full_name, dob, mobile, email,       │
//  │                │ password_hash, biometric_code,       │
//  │                │ device_id                           │
//  │  UserLocMap    | map_id, user_id, loc_id, allowed_m  │
//  │  Sessions      | sess_id, teacher_id, subject, date, │
//  │                │ start_time, end_time, status, win_m │
//  │  Attendance    | att_id, user_id, sess_id,           │
//  │                │ att_type, att_date, att_time,       │
//  │                │ login_method, loc_id, lat, lng,     │
//  │                │ address, dist_from_centre           │
//  │  _Cache        | key, value, expires_at              │
//  └──────────────────────────────────────────────────────┘
//
//  KEY OPTIMIZATIONS:
//  • full_name removed from Attendance (JOIN via user_id)
//  • Duplicate att_time column eliminated
//  • Entry+Exit merged via att_type ('entry'|'exit')
//  • _Cache sheet used for short-lived server-side caching
//  • Single getDataRange() per operation (no repeated reads)
//  • In-memory maps replace O(n) loops where possible
//  • All date cells stored as yyyy-MM-dd strings (no Date objects)
// ============================================================

// ── Sheet names ───────────────────────────────────────────────
var S_ROLES     = 'Roles';
var S_DEPTS     = 'Departments';
var S_LOCS      = 'Locations';
var S_USERS     = 'Users';
var S_USERLOCMAP= 'UserLocMap';
var S_SESSIONS  = 'Sessions';
var S_ATTENDANCE= 'Attendance';
var S_CACHE     = '_Cache';

// ── Default college GPS anchor (fallback) ────────────────────
var DEFAULT_LAT = 13.32603;
var DEFAULT_LNG = 77.12621;
var DEFAULT_LOC_ID = 'loc_default';

// ── Header definitions ────────────────────────────────────────
var HEADERS = {};
HEADERS[S_ROLES]      = ['role_id','name'];
HEADERS[S_DEPTS]      = ['dept_id','name','incharge','email'];
HEADERS[S_LOCS]       = ['loc_id','name','lat','lng'];
HEADERS[S_USERS]      = ['user_id','inst_id','dept_id','role_id','full_name','dob','mobile','email','password_hash','biometric_code','device_id'];
HEADERS[S_USERLOCMAP] = ['map_id','user_id','loc_id','allowed_distance_m'];
HEADERS[S_SESSIONS]   = ['sess_id','teacher_id','subject','date','start_time','end_time','status','window_minutes'];
HEADERS[S_ATTENDANCE] = ['att_id','user_id','sess_id','att_type','att_date','att_time','login_method','loc_id','lat','lng','address','dist_from_centre'];
HEADERS[S_CACHE]      = ['cache_key','value','expires_at'];

// ── Haversine distance (metres) ───────────────────────────────
function haversineMeters(lat1, lng1, lat2, lng2) {
  var R  = 6371000;
  var p1 = lat1 * Math.PI / 180, p2 = lat2 * Math.PI / 180;
  var dp = (lat2 - lat1) * Math.PI / 180;
  var dl = (lng2 - lng1) * Math.PI / 180;
  var a  = Math.sin(dp/2)*Math.sin(dp/2) + Math.cos(p1)*Math.cos(p2)*Math.sin(dl/2)*Math.sin(dl/2);
  return Math.round(6371000 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
}

// ── JSON output ───────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Entry points ──────────────────────────────────────────────
function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (!p.data) return jsonOut({ status:'BioAttend API v5', time: new Date().toISOString() });
    return jsonOut(route(JSON.parse(decodeURIComponent(p.data))));
  } catch(err) { return jsonOut({ success:false, message:'doGet: '+err }); }
}
function doPost(e) {
  try { return jsonOut(route(JSON.parse(e.postData.contents))); }
  catch(err) { return jsonOut({ success:false, message:'doPost: '+err }); }
}

// ── Router ────────────────────────────────────────────────────
function route(body) {
  switch (body.action) {
    // Auth
    case 'register':          return registerUser(body);
    case 'signIn':            return signInUser(body);
    // Biometric & Device
    case 'saveBiometric':     return saveBiometric(body);
    case 'getBiometric':      return getBiometric(body);
    case 'registerDevice':    return registerDevice(body);
    case 'checkDevice':       return checkDevice(body);
    // Attendance
    case 'markAttendance':    return markAttendance(body);
    case 'markExit':          return markExit(body);
    // Sessions
    case 'createSession':     return createSession(body);
    case 'closeSession':      return closeSession(body);
    case 'getActiveSession':  return getActiveSession(body);
    case 'getSessions':       return getSessions(body);
    // Reporting
    case 'getAttendance':     return getAttendance(body);
    case 'getDashboard':      return getDashboard(body);
    case 'getStudents':       return getStudents(body);
    case 'exportAttendance':  return exportAttendance(body);
    // Lookup tables (for front-end dropdowns)
    case 'getDepartments':    return getDepartments();
    case 'getRoles':          return getRoles();
    case 'getLocations':      return getLocations();
    // Admin / Setup
    case 'setupSheets':       return setupAllSheets();
    case 'seedDefaults':      return seedDefaultData();
    case 'clearCache':        return clearCache();
    case 'debug':             return debugInfo();
    default: return { success:false, message:'Unknown action: '+body.action };
  }
}

// ════════════════════════════════════════════════════════════
//  SHEET / DATA HELPERS
// ════════════════════════════════════════════════════════════

// Get or create sheet with correct headers
function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var h = HEADERS[name] || [];
    if (h.length) {
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length)
        .setFontWeight('bold')
        .setBackground('#1a2a52')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// Returns array of row objects keyed by header.
// OPTIMIZATION: single getDataRange() call per sheet per operation.
function getRows(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var h = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    h.forEach(function(col, i) { obj[col] = row[i]; });
    return obj;
  });
}

// Build an in-memory index map: { fieldValue -> row }
// Use instead of O(n) loops when looking up single records.
function buildIndex(rows, field) {
  var map = {};
  rows.forEach(function(r) { map[String(r[field]||'').trim()] = r; });
  return map;
}

// Find 1-based column index by header name
function colIndex(sheet, name) {
  var h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var i = h.indexOf(name);
  return i === -1 ? -1 : i + 1;
}

// Update a single cell in a row where keyField === keyValue
function updateCell(sheet, keyField, keyValue, targetField, newValue) {
  var data = sheet.getDataRange().getValues();
  var h    = data[0];
  var keyCol = h.indexOf(keyField);
  var tgtCol = h.indexOf(targetField);
  if (keyCol < 0 || tgtCol < 0) return false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]) === String(keyValue)) {
      sheet.getRange(i+1, tgtCol+1).setValue(newValue);
      return true;
    }
  }
  return false;
}

// ID generator
function genId(prefix) {
  return (prefix||'id')+'_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);
}

// SHA-256 password hash
function hashPw(pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8)
    .map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
}

// Normalize any date value to yyyy-MM-dd string
function normDate(raw) {
  if (!raw && raw !== 0) return '';
  var tz = Session.getScriptTimeZone();
  if (raw instanceof Date) return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
  var s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (s.length > 10 && s.indexOf('T') > 0) return s.substring(0,10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return s.substring(0,10);
}

// ════════════════════════════════════════════════════════════
//  SERVER-SIDE CACHE  (_Cache sheet)
//  TTL-based, keyed strings. Keeps dashboard fast.
// ════════════════════════════════════════════════════════════

function cacheGet(key) {
  try {
    var sheet = getSheet(S_CACHE);
    var rows  = getRows(sheet);
    var now   = Date.now();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].cache_key === key) {
        if (rows[i].expires_at && Number(rows[i].expires_at) > now) {
          return JSON.parse(rows[i].value);
        }
        return null; // expired
      }
    }
  } catch(e) {}
  return null;
}

function cacheSet(key, value, ttlSeconds) {
  try {
    var sheet   = getSheet(S_CACHE);
    var data    = sheet.getDataRange().getValues();
    var h       = data[0];
    var keyCol  = h.indexOf('cache_key');
    var expires = Date.now() + (ttlSeconds||30)*1000;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][keyCol]) === key) {
        sheet.getRange(i+1, 2).setValue(JSON.stringify(value));
        sheet.getRange(i+1, 3).setValue(expires);
        return;
      }
    }
    sheet.appendRow([key, JSON.stringify(value), expires]);
  } catch(e) {}
}

function clearCache() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s  = ss.getSheetByName(S_CACHE);
    if (s) {
      var last = s.getLastRow();
      if (last > 1) s.deleteRows(2, last-1);
    }
    return { success:true, message:'Cache cleared' };
  } catch(e) { return { success:false, message:e.toString() }; }
}

// ════════════════════════════════════════════════════════════
//  LOOKUP HELPERS — resolve IDs to names in-memory
// ════════════════════════════════════════════════════════════

function loadUserMap() {
  return buildIndex(getRows(getSheet(S_USERS)), 'user_id');
}

function loadDeptMap() {
  return buildIndex(getRows(getSheet(S_DEPTS)), 'dept_id');
}

function loadLocMap() {
  return buildIndex(getRows(getSheet(S_LOCS)), 'loc_id');
}

// Find allowed locations for a user (from UserLocMap)
// Returns array of { loc_id, name, lat, lng, allowed_distance_m }
function getUserLocations(userId) {
  var maps  = getRows(getSheet(S_USERLOCMAP)).filter(function(r){ return r.user_id === userId; });
  var locMap= loadLocMap();
  if (!maps.length) {
    // Fallback: use default lab location
    var defLoc = locMap[DEFAULT_LOC_ID];
    if (defLoc) return [{ loc_id:DEFAULT_LOC_ID, name:defLoc.name, lat:parseFloat(defLoc.lat), lng:parseFloat(defLoc.lng), allowed_distance_m:200 }];
    return [{ loc_id:DEFAULT_LOC_ID, name:'Lab', lat:DEFAULT_LAT, lng:DEFAULT_LNG, allowed_distance_m:200 }];
  }
  return maps.map(function(m) {
    var loc = locMap[m.loc_id] || {};
    return { loc_id:m.loc_id, name:loc.name||m.loc_id, lat:parseFloat(loc.lat||0), lng:parseFloat(loc.lng||0), allowed_distance_m:parseInt(m.allowed_distance_m)||200 };
  });
}

// ════════════════════════════════════════════════════════════
//  1. REGISTER USER
// ════════════════════════════════════════════════════════════
function registerUser(body) {
  try {
    if (!body.full_name || !body.email || !body.password)
      return { success:false, message:'full_name, email and password required' };

    var sheet = getSheet(S_USERS);
    var rows  = getRows(sheet);
    // O(n) email dupe check — acceptable for registration
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].email).toLowerCase() === String(body.email).toLowerCase())
        return { success:false, message:'Email already registered' };
    }

    // Resolve role_id
    var roleRows = getRows(getSheet(S_ROLES));
    var roleMap  = buildIndex(roleRows, 'name');
    var roleName = body.role || 'student';
    var roleRow  = roleMap[roleName];
    if (!roleRow) return { success:false, message:'Invalid role: '+roleName };

    var userId = genId('u');
    sheet.appendRow([
      userId,
      body.inst_id    || 'SIT_Tumkur',
      body.dept_id    || '',
      roleRow.role_id,
      body.full_name,
      body.dob        || '',
      body.mobile     || '',
      body.email,
      hashPw(body.password),
      '',   // biometric_code (set later)
      ''    // device_id (set later)
    ]);

    return { success:true, user_id:userId, role:roleName, message:'Account created' };
  } catch(e) { return { success:false, message:'register: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  2. SIGN IN
// ════════════════════════════════════════════════════════════
function signInUser(body) {
  try {
    var rows = getRows(getSheet(S_USERS));
    var hash = hashPw(body.password || '');
    var roleMap = buildIndex(getRows(getSheet(S_ROLES)), 'role_id');

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r.email).toLowerCase() === String(body.email||'').toLowerCase() && r.password_hash === hash) {
        var roleName = (roleMap[r.role_id]||{}).name || 'student';
        return { success:true, user_id:r.user_id, name:r.full_name, role:roleName, dept_id:r.dept_id };
      }
    }
    return { success:false, message:'Invalid email or password' };
  } catch(e) { return { success:false, message:'signIn: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  3. MARK ATTENDANCE  (att_type = 'entry')
//  OPTIMIZATION: resolves geofence against user's assigned
//  locations from UserLocMap — falls back to default if none.
// ════════════════════════════════════════════════════════════
function markAttendance(body) {
  try {
    if (!body.user_id) return { success:false, message:'user_id required' };

    var now   = new Date();
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    // Load user — single pass
    var userMap = loadUserMap();
    var user    = userMap[body.user_id];
    if (!user) return { success:false, message:'User not found' };

    // Geofence check against user's assigned locations
    var lat = body.lat !== undefined && body.lat !== '' ? parseFloat(body.lat) : null;
    var lng = body.lng !== undefined && body.lng !== '' ? parseFloat(body.lng) : null;
    var matchedLoc = null;
    var distFromCentre = '';

    if (lat !== null && lng !== null && !isNaN(lat) && !isNaN(lng)) {
      var allowedLocs = getUserLocations(body.user_id);
      var minDist = Infinity, closestLoc = null;
      for (var k = 0; k < allowedLocs.length; k++) {
        var d = haversineMeters(lat, lng, allowedLocs[k].lat, allowedLocs[k].lng);
        if (d < minDist) { minDist = d; closestLoc = allowedLocs[k]; }
        if (d <= allowedLocs[k].allowed_distance_m) { matchedLoc = allowedLocs[k]; distFromCentre = d; break; }
      }
      if (!matchedLoc) {
        return {
          success:  false,
          code:     'TOO_FAR',
          distance: minDist,
          location: closestLoc ? closestLoc.name : 'lab',
          message:  'You are '+minDist+'m from '+((closestLoc||{}).name||'the lab')+'. Must be within '+(closestLoc ? closestLoc.allowed_distance_m : 200)+'m.'
        };
      }
      distFromCentre = minDist;
    }

    // Duplicate check for today — only read Attendance once
    var attSheet = getSheet(S_ATTENDANCE);
    var existing = getRows(attSheet);
    for (var j = 0; j < existing.length; j++) {
      var row = existing[j];
      if (String(row.user_id) === String(body.user_id) &&
          row.att_type === 'entry' &&
          normDate(row.att_date) === today) {
        return { success:false, message:'Entry already marked today at '+row.att_time };
      }
    }

    // Determine sess_id from active session (optional)
    var sessId = body.sess_id || '';
    if (!sessId) {
      var active = _getActiveSessionData();
      if (active) sessId = active.sess_id;
    }

    var locId = (matchedLoc && matchedLoc.loc_id) ? matchedLoc.loc_id : (body.loc_id || DEFAULT_LOC_ID);

    attSheet.appendRow([
      genId('att'),            // att_id
      body.user_id,            // user_id (FK — no full_name redundancy)
      sessId,                  // sess_id
      'entry',                 // att_type
      today,                   // att_date  (string, never Date object)
      timeStr,                 // att_time
      body.login_method || 'biometric', // login_method
      locId,                   // loc_id (FK)
      lat !== null ? lat : '', // lat
      lng !== null ? lng : '', // lng
      body.address || '',      // address
      distFromCentre           // dist_from_centre
    ]);

    return {
      success:           true,
      message:           '✓ Entry marked at '+timeStr,
      name:              user.full_name,
      date:              today,
      time:              timeStr,
      location:          body.address || (matchedLoc ? matchedLoc.name : ''),
      dist_from_centre:  distFromCentre
    };
  } catch(e) { return { success:false, message:'markAttendance: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  4. MARK EXIT  (att_type = 'exit')
//  Inserts a NEW row instead of updating — cleaner audit trail
//  and avoids slow row-scanning updates.
// ════════════════════════════════════════════════════════════
function markExit(body) {
  try {
    if (!body.user_id) return { success:false, message:'user_id required' };

    var now     = new Date();
    var tz      = Session.getScriptTimeZone();
    var today   = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    var attSheet = getSheet(S_ATTENDANCE);
    var rows     = getRows(attSheet);

    var entryRow = null, exitExists = false;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r.user_id) !== String(body.user_id)) continue;
      if (normDate(r.att_date) !== today) continue;
      if (r.att_type === 'entry') entryRow = r;
      if (r.att_type === 'exit')  exitExists = true;
    }

    if (!entryRow)   return { success:false, message:'No entry record found for today. Mark entry first.' };
    if (exitExists)  return { success:false, message:'Exit already recorded today.' };

    // Duration calculation
    var entryDt  = new Date(today+'T'+entryRow.att_time);
    var diffMins = Math.max(0, Math.round((now - entryDt) / 60000));
    var duration = Math.floor(diffMins/60)+'h '+( diffMins%60)+'m';

    var lat = body.lat !== undefined && body.lat !== '' ? parseFloat(body.lat) : '';
    var lng = body.lng !== undefined && body.lng !== '' ? parseFloat(body.lng) : '';
    var dist = (lat !== '' && lng !== '') ? haversineMeters(lat, lng, DEFAULT_LAT, DEFAULT_LNG) : '';

    attSheet.appendRow([
      genId('att'),
      body.user_id,
      entryRow.sess_id || '',
      'exit',
      today,
      timeStr,
      body.login_method || 'biometric',
      entryRow.loc_id || DEFAULT_LOC_ID,
      lat, lng,
      body.address || '',
      dist
    ]);

    return { success:true, message:'✓ Exit recorded at '+timeStr+' · Duration: '+duration, exit_time:timeStr, duration:duration };
  } catch(e) { return { success:false, message:'markExit: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  5. SESSIONS
// ════════════════════════════════════════════════════════════
function createSession(body) {
  try {
    if (!body.subject || !body.window_minutes)
      return { success:false, message:'subject and window_minutes required' };

    // Role check via role lookup
    var userRows = getRows(getSheet(S_USERS));
    var roleMap  = buildIndex(getRows(getSheet(S_ROLES)), 'role_id');
    var user     = buildIndex(userRows,'user_id')[body.user_id];
    if (!user) return { success:false, message:'User not found' };
    var roleName = (roleMap[user.role_id]||{}).name||'student';
    if (roleName !== 'teacher' && roleName !== 'admin')
      return { success:false, message:'Only teachers can create sessions' };

    var now   = new Date(), tz = Session.getScriptTimeZone();
    var date  = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var start = Utilities.formatDate(now, tz, 'HH:mm:ss');
    var end   = Utilities.formatDate(new Date(now.getTime()+parseInt(body.window_minutes)*60000), tz, 'HH:mm:ss');
    var sessId= genId('sess');

    // Close any open sessions by this teacher
    var sessSheet = getSheet(S_SESSIONS);
    var sessData  = sessSheet.getDataRange().getValues();
    var statusCol = sessData[0].indexOf('status');
    var tidCol    = sessData[0].indexOf('teacher_id');
    for (var i = 1; i < sessData.length; i++) {
      if (sessData[i][tidCol] === body.user_id && sessData[i][statusCol] === 'open')
        sessSheet.getRange(i+1, statusCol+1).setValue('closed');
    }

    sessSheet.appendRow([sessId, body.user_id, body.subject, date, start, end, 'open', body.window_minutes]);
    return { success:true, sess_id:sessId, subject:body.subject, start_time:start, end_time:end, message:'Session opened for '+body.window_minutes+' min' };
  } catch(e) { return { success:false, message:'createSession: '+e }; }
}

function closeSession(body) {
  try {
    var sheet = getSheet(S_SESSIONS);
    var data  = sheet.getDataRange().getValues();
    var h     = data[0], idCol = h.indexOf('sess_id'), stCol = h.indexOf('status');
    for (var i = 1; i < data.length; i++) {
      if (data[i][idCol] === body.sess_id) {
        sheet.getRange(i+1, stCol+1).setValue('closed');
        return { success:true };
      }
    }
    return { success:false, message:'Session not found' };
  } catch(e) { return { success:false, message:'closeSession: '+e }; }
}

// Internal helper — avoids double sheet load
function _getActiveSessionData() {
  var now    = new Date(), tz = Session.getScriptTimeZone();
  var today  = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var rows   = getRows(getSheet(S_SESSIONS));
  for (var i = 0; i < rows.length; i++) {
    var s = rows[i];
    if (s.status === 'open' && normDate(s.date) === today) {
      var st = new Date(s.date+'T'+s.start_time);
      var en = new Date(s.date+'T'+s.end_time);
      if (now >= st && now <= en) return s;
    }
  }
  return null;
}

function getActiveSession(body) {
  try {
    var s = _getActiveSessionData();
    if (!s) return { success:true, active:false };
    var en   = new Date(s.date+'T'+s.end_time);
    var secs = Math.max(0, Math.round((en - new Date()) / 1000));
    return { success:true, active:true, session:s, seconds_left:secs };
  } catch(e) { return { success:false, message:'getActiveSession: '+e }; }
}

function getSessions(body) {
  try {
    var rows = getRows(getSheet(S_SESSIONS));
    if (body.user_id) rows = rows.filter(function(r){ return r.teacher_id === body.user_id; });

    // Count present students per session — batch in-memory
    var attRows = getRows(getSheet(S_ATTENDANCE)).filter(function(r){ return r.att_type === 'entry'; });
    var sessCounts = {};
    attRows.forEach(function(a){ sessCounts[a.sess_id] = (sessCounts[a.sess_id]||0)+1; });
    rows.forEach(function(s){ s.present_count = sessCounts[s.sess_id]||0; });

    rows.reverse();
    return { success:true, sessions:rows.slice(0,30) };
  } catch(e) { return { success:false, message:'getSessions: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  6. ATTENDANCE QUERIES
// ════════════════════════════════════════════════════════════
function getAttendance(body) {
  try {
    var rows = getRows(getSheet(S_ATTENDANCE));
    if (body.sess_id)   rows = rows.filter(function(r){ return r.sess_id   === body.sess_id; });
    if (body.att_date)  rows = rows.filter(function(r){ return normDate(r.att_date) === body.att_date; });
    if (body.user_id)   rows = rows.filter(function(r){ return r.user_id   === body.user_id; });
    if (body.att_type)  rows = rows.filter(function(r){ return r.att_type  === body.att_type; });

    // JOIN full_name & dept in-memory (no redundant storage)
    var userMap = loadUserMap(), deptMap = loadDeptMap(), locMap = loadLocMap();
    rows.forEach(function(r) {
      var u = userMap[r.user_id]||{};
      r.full_name = u.full_name||'';
      r.email     = u.email||'';
      r.dept_name = (deptMap[u.dept_id]||{}).name||'';
      r.loc_name  = (locMap[r.loc_id]||{}).name||'';
    });

    return { success:true, attendance:rows };
  } catch(e) { return { success:false, message:'getAttendance: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  7. DASHBOARD  — cached for 30 seconds to handle rapid
//     refreshes without hammering the sheet API.
// ════════════════════════════════════════════════════════════
function getDashboard(body) {
  try {
    if (!body.sess_id) return { success:false, message:'sess_id required' };

    var cacheKey = 'dashboard_'+body.sess_id;
    var cached   = cacheGet(cacheKey);
    if (cached)  return cached;

    // All students (role = student)
    var roleRows = getRows(getSheet(S_ROLES));
    var roleMap  = buildIndex(roleRows, 'name');
    var studentRoleId = (roleMap['student']||{}).role_id || '';

    var userRows = getRows(getSheet(S_USERS));
    var deptMap  = loadDeptMap();
    var students = userRows.filter(function(r){ return r.role_id === studentRoleId; });

    // Entry records for this session — one pass
    var attRows = getRows(getSheet(S_ATTENDANCE));
    var presentMap = {};
    attRows.forEach(function(r){
      if (r.sess_id === body.sess_id && r.att_type === 'entry')
        presentMap[r.user_id] = r;
    });
    // Exit records
    var exitMap = {};
    attRows.forEach(function(r){
      if (r.sess_id === body.sess_id && r.att_type === 'exit')
        exitMap[r.user_id] = r;
    });

    var present = [], absent = [];
    students.forEach(function(s) {
      var deptName = (deptMap[s.dept_id]||{}).name||'';
      if (presentMap[s.user_id]) {
        var pr  = presentMap[s.user_id];
        var ex  = exitMap[s.user_id];
        present.push({
          user_id:         s.user_id,
          name:            s.full_name,
          email:           s.email,
          dept:            deptName,
          att_time:        pr.att_time,
          login_method:    pr.login_method,
          dist_from_centre:pr.dist_from_centre,
          lat:             pr.lat,
          lng:             pr.lng,
          exit_time:       ex ? ex.att_time : null
        });
      } else {
        absent.push({ user_id:s.user_id, name:s.full_name, email:s.email, dept:deptName });
      }
    });

    var result = {
      success:       true,
      total:         students.length,
      present_count: present.length,
      absent_count:  absent.length,
      pct:           students.length ? Math.round(present.length/students.length*100) : 0,
      present:       present,
      absent:        absent
    };

    cacheSet(cacheKey, result, 30); // cache for 30 seconds
    return result;
  } catch(e) { return { success:false, message:'getDashboard: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  8. STUDENTS LIST
// ════════════════════════════════════════════════════════════
function getStudents(body) {
  try {
    var roleMap  = buildIndex(getRows(getSheet(S_ROLES)), 'name');
    var studRoleId = (roleMap['student']||{}).role_id||'';
    var deptMap  = loadDeptMap();
    var students = getRows(getSheet(S_USERS))
      .filter(function(r){ return r.role_id === studRoleId; })
      .map(function(r) {
        return {
          user_id:   r.user_id,
          full_name: r.full_name,
          email:     r.email,
          dept:      (deptMap[r.dept_id]||{}).name||'',
          mobile:    r.mobile,
          has_bio:   !!r.biometric_code,
          has_device:!!r.device_id
        };
      });
    return { success:true, students:students };
  } catch(e) { return { success:false, message:'getStudents: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  9. EXPORT CSV
// ════════════════════════════════════════════════════════════
function exportAttendance(body) {
  try {
    var rows = getRows(getSheet(S_ATTENDANCE));
    if (body.sess_id)  rows = rows.filter(function(r){ return r.sess_id === body.sess_id; });
    if (body.att_date) rows = rows.filter(function(r){ return normDate(r.att_date) === body.att_date; });

    var userMap = loadUserMap(), deptMap = loadDeptMap(), locMap = loadLocMap();
    var header  = ['Name','Email','Department','Date','Time','Type','Method','Location','Distance(m)'];
    var lines   = [header.join(',')];

    rows.forEach(function(r) {
      var u = userMap[r.user_id]||{};
      lines.push([
        '"'+(u.full_name||'')+'"',
        '"'+(u.email||'')+'"',
        '"'+((deptMap[u.dept_id]||{}).name||'')+'"',
        r.att_date||'',
        r.att_time||'',
        r.att_type||'',
        r.login_method||'',
        '"'+((locMap[r.loc_id]||{}).name||'')+'"',
        r.dist_from_centre||''
      ].join(','));
    });

    return { success:true, csv:lines.join('\n'), row_count:rows.length };
  } catch(e) { return { success:false, message:'exportAttendance: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  10. BIOMETRIC & DEVICE
// ════════════════════════════════════════════════════════════
function saveBiometric(body) {
  try {
    if (!body.user_id||!body.credential_id) return { success:false, message:'user_id and credential_id required' };
    var ok = updateCell(getSheet(S_USERS),'user_id',body.user_id,'biometric_code',body.credential_id);
    return ok ? { success:true } : { success:false, message:'User not found' };
  } catch(e) { return { success:false, message:'saveBiometric: '+e }; }
}

function getBiometric(body) {
  try {
    var rows = getRows(getSheet(S_USERS));
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].email).toLowerCase() === String(body.email||'').toLowerCase()) {
        if (!rows[i].biometric_code) return { success:false, message:'No biometric registered' };
        return { success:true, credential_id:rows[i].biometric_code, user_id:rows[i].user_id, name:rows[i].full_name };
      }
    }
    return { success:false, message:'User not found' };
  } catch(e) { return { success:false, message:'getBiometric: '+e }; }
}

function registerDevice(body) {
  try {
    if (!body.user_id||!body.device_id) return { success:false, message:'user_id and device_id required' };
    var sheet = getSheet(S_USERS);
    var data  = sheet.getDataRange().getValues();
    var h     = data[0];
    var uidCol= h.indexOf('user_id'), devCol = h.indexOf('device_id');
    if (uidCol<0||devCol<0) return { success:false, message:'Column not found' };
    for (var i = 1; i < data.length; i++) {
      if (data[i][uidCol] === body.user_id) {
        var existing = String(data[i][devCol]||'').trim();
        if (existing && existing !== String(body.device_id).trim())
          return { success:false, already_bound:true, message:'Account bound to another device.' };
        sheet.getRange(i+1, devCol+1).setValue(body.device_id);
        return { success:true, already_bound:false, message: existing?'Device confirmed':'Device registered' };
      }
    }
    return { success:false, message:'User not found' };
  } catch(e) { return { success:false, message:'registerDevice: '+e }; }
}

function checkDevice(body) {
  try {
    if (!body.user_id||!body.device_id) return { success:false, message:'user_id and device_id required' };
    var rows = getRows(getSheet(S_USERS));
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].user_id === body.user_id) {
        var stored = String(rows[i].device_id||'').trim();
        if (!stored)                              return { success:true,  status:'unbound' };
        if (stored === String(body.device_id).trim()) return { success:true,  status:'match' };
        return { success:false, status:'mismatch', message:'Account registered to a different device.' };
      }
    }
    return { success:false, message:'User not found' };
  } catch(e) { return { success:false, message:'checkDevice: '+e }; }
}

// ════════════════════════════════════════════════════════════
//  11. LOOKUP TABLES (for front-end dropdowns)
// ════════════════════════════════════════════════════════════
function getDepartments() {
  try { return { success:true, departments:getRows(getSheet(S_DEPTS)) }; }
  catch(e) { return { success:false, message:e.toString() }; }
}
function getRoles() {
  try { return { success:true, roles:getRows(getSheet(S_ROLES)) }; }
  catch(e) { return { success:false, message:e.toString() }; }
}
function getLocations() {
  try { return { success:true, locations:getRows(getSheet(S_LOCS)) }; }
  catch(e) { return { success:false, message:e.toString() }; }
}

// ════════════════════════════════════════════════════════════
//  12. SETUP — creates all sheets with correct headers
//  Run once from Apps Script editor: setupAllSheets()
// ════════════════════════════════════════════════════════════
function setupAllSheets() {
  try {
    [S_ROLES, S_DEPTS, S_LOCS, S_USERS, S_USERLOCMAP, S_SESSIONS, S_ATTENDANCE, S_CACHE]
      .forEach(function(name){ getSheet(name); });
    return { success:true, message:'All sheets created/verified.' };
  } catch(e) { return { success:false, message:e.toString() }; }
}

// ════════════════════════════════════════════════════════════
//  13. SEED DEFAULT DATA  (run once after setupSheets)
// ════════════════════════════════════════════════════════════
function seedDefaultData() {
  try {
    // Roles
    var rSheet = getSheet(S_ROLES);
    if (getRows(rSheet).length === 0) {
      rSheet.appendRow(['role_student', 'student']);
      rSheet.appendRow(['role_teacher', 'teacher']);
      rSheet.appendRow(['role_admin',   'admin']);
    }

    // Default location (lab)
    var lSheet = getSheet(S_LOCS);
    if (getRows(lSheet).length === 0) {
      lSheet.appendRow([DEFAULT_LOC_ID, 'Main Lab – SIT Tumkur', DEFAULT_LAT, DEFAULT_LNG]);
    }

    // Sample departments
    var dSheet = getSheet(S_DEPTS);
    if (getRows(dSheet).length === 0) {
      var depts = [
        ['dept_cse','Computer Science & Engineering','',''],
        ['dept_ise','Information Science & Engineering','',''],
        ['dept_ece','Electronics & Communication','',''],
        ['dept_eee','Electrical & Electronics','',''],
        ['dept_mech','Mechanical Engineering','',''],
        ['dept_civil','Civil Engineering','',''],
        ['dept_chem','Chemical Engineering','',''],
        ['dept_bio','Biotechnology','',''],
        ['dept_mba','MBA','',''],
        ['dept_mca','MCA','','']
      ];
      depts.forEach(function(d){ dSheet.appendRow(d); });
    }

    return { success:true, message:'Default data seeded.' };
  } catch(e) { return { success:false, message:e.toString() }; }
}

// ════════════════════════════════════════════════════════════
//  14. DEBUG
// ════════════════════════════════════════════════════════════
function debugInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      success:    true,
      name:       ss.getName(),
      sheets:     ss.getSheets().map(function(s){ return { name:s.getName(), rows:s.getLastRow()-1 }; }),
      tz:         Session.getScriptTimeZone(),
      schema_ver: 'v5'
    };
  } catch(e) { return { success:false, message:'debug: '+e }; }
}