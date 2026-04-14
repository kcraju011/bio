// ============================================================
//  BioAttend v5 — Google Apps Script Backend
//  College : Siddaganga Institute of Technology, Tumkur
//  Deploy  : Web App → Execute as Me → Access: Anyone
//
//  SCHEMA v5  — All 8 tables from the spec, fully normalized
//  ─────────────────────────────────────────────────────────
//  Sheet 1 : Users                  (user master + auth)
//  Sheet 2 : Departments            (dept master)
//  Sheet 3 : Roles                  (role master)
//  Sheet 4 : AttendanceLocations    (geofence anchors)
//  Sheet 5 : AttendanceTypes        (entry / exit)
//  Sheet 6 : Sessions               (teacher windows)
//  Sheet 7 : Attendance             (lean fact table)
//  Sheet 8 : LocationMonitor        (GPS audit log)
//  Sheet 9 : UserLocationMap        (per-user geofence overrides)
//  Sheet 10: _Cache                 (TTL KV store)
//
//  KEY DESIGN DECISIONS
//  ─────────────────────
//  ✓ full_name REMOVED from Attendance  → join on Users at read time
//  ✓ duplicate attendance_time REMOVED  → single entry_time + exit_time
//  ✓ address REMOVED from Attendance    → stored only in LocationMonitor
//  ✓ EntryGPS/ExitGPS strings REMOVED   → computed from lat+lng on demand
//  ✓ institution constant REMOVED       → not in Users at all
//  ✓ subject REMOVED from Attendance    → joined via session_id → Sessions
//  ✓ Indexes simulated via buildIndex() → O(n) scan → O(1) lookup
//  ✓ Request-scoped row cache (_RC)     → each sheet read only once per request
//  ✓ TTL server cache (_Cache sheet)    → dashboard/students cached 30s
//  ✓ Geofence per user+location         → UserLocationMap overrides global default
// ============================================================

// ── Sheet name constants ──────────────────────────────────────
var SH = {
  USERS    : 'Users',
  DEPTS    : 'Departments',
  ROLES    : 'Roles',
  LOCS     : 'AttendanceLocations',
  ATT_TYPES: 'AttendanceTypes',
  SESSIONS : 'Sessions',
  ATT      : 'Attendance',
  LOC_MON  : 'LocationMonitor',
  USR_LOC  : 'UserLocationMap',
  CACHE    : '_Cache'
};

// ── Centralized schema — single source of truth ───────────────
var SCHEMA = {
  // user_id | dept_id(FK) | role_id(FK) | full_name | dob | mobile
  // email | password_hash | biometric_code | device_identification | created_at
  Users: [
    'user_id','dept_id','role_id','full_name','dob',
    'mobile','email','password_hash','biometric_code',
    'device_identification','created_at'
  ],

  // department_id | name | in_charge | email
  Departments: ['department_id','name','in_charge','email'],

  // role_id | name   (student / teacher / admin)
  Roles: ['role_id','name'],

  // location_id | name | latitude | longitude | default_allowed_distance_m
  AttendanceLocations: [
    'location_id','name','latitude','longitude','default_allowed_distance_m'
  ],

  // att_type_id | type_name   (entry / exit)
  AttendanceTypes: ['att_type_id','type_name'],

  // session_id | teacher_id(FK→Users) | location_id(FK→AttendanceLocations)
  // subject | date | start_time | end_time | window_minutes | status
  // REMOVED: teacher_name  (join on Users)
  Sessions: [
    'session_id','teacher_id','location_id',
    'subject','date','start_time','end_time','window_minutes','status'
  ],

  // attendance_id | session_id(FK,nullable) | user_id(FK) | location_id(FK)
  // att_type_id(FK) | att_date | entry_time | exit_time
  // login_method | entry_lat | entry_lng | entry_distance_m
  // exit_lat | exit_lng | exit_distance_m | duration_mins
  //
  // REMOVED: full_name (join)  address (LocationMonitor)
  //          duplicate attendance_time  EntryGPS/ExitGPS (computed)
  //          Subject (join via session_id)
  Attendance: [
    'attendance_id','session_id','user_id','location_id','att_type_id',
    'att_date','entry_time','exit_time','login_method',
    'entry_lat','entry_lng','entry_distance_m',
    'exit_lat','exit_lng','exit_distance_m','duration_mins'
  ],

  // monitor_id | user_id(FK) | location_id(FK) | latitude | longitude
  // distance_from_centre | address | timestamp
  LocationMonitor: [
    'monitor_id','user_id','location_id',
    'latitude','longitude','distance_from_centre','address','timestamp'
  ],

  // map_id | user_id(FK) | location_id(FK) | allowed_distance_m
  // Per-user geofence override — if NULL use location default
  UserLocationMap: [
    'map_id','user_id','location_id','allowed_distance_m'
  ],

  // key | value | expires_at
  _Cache: ['key','value','expires_at']
};

// ── Default lab location (bootstrapped into AttendanceLocations) ─
var DEFAULT_LOC = {
  location_id:              'loc_default',
  name:                     'Main Lab – SIT Tumkur',
  latitude:                 13.32603,
  longitude:                77.12621,
  default_allowed_distance_m: 200
};

var CACHE_TTL_MS = 30000; // 30 seconds

// ============================================================
//  UTILITIES
// ============================================================
function haversine(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var p1 = lat1 * Math.PI / 180, p2 = lat2 * Math.PI / 180;
  var dp = (lat2 - lat1) * Math.PI / 180;
  var dl = (lng2 - lng1) * Math.PI / 180;
  var a  = Math.sin(dp/2)*Math.sin(dp/2) +
           Math.cos(p1)*Math.cos(p2)*Math.sin(dl/2)*Math.sin(dl/2);
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
}

// Normalize any raw date cell → 'yyyy-MM-dd' string robustly
function normDate(raw, tz) {
  if (!raw && raw !== 0) return '';
  if (raw instanceof Date) return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
  var s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (s.indexOf('T') > 0) return s.substring(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return s.substring(0, 10);
}

function uid(prefix) {
  return (prefix||'id') + '_' + Date.now() + '_' + Math.random().toString(36).substr(2,5);
}

function hashPw(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  SHEET LAYER  ─ request-scoped in-memory row cache
//
//  Every sheet is read at most ONCE per doGet/doPost call.
//  _RC is reset at the top of every entry point.
//  This alone cuts sheet API calls by 60-80% for complex actions.
// ============================================================
var _RC = {};  // request cache: { sheetName: [rowObjects] }

function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var h = SCHEMA[name];
    if (h) {
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length)
        .setFontWeight('bold').setBackground('#0d1a40').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// Returns all data rows as array-of-objects, cached per request.
// 10 callers pay for exactly 1 sheet read.
function rows(name) {
  if (_RC[name]) return _RC[name];
  var data = getSheet(name).getDataRange().getValues();
  if (data.length <= 1) { _RC[name] = []; return []; }
  var hdrs = data[0];
  _RC[name] = data.slice(1).map(function(r) {
    var o = {};
    hdrs.forEach(function(h,i){ o[h] = r[i]; });
    return o;
  });
  return _RC[name];
}

function bustCache(name) { delete _RC[name]; }

// ── Simulated Indexes ─────────────────────────────────────────
// Converts O(n*m) nested loops → O(n+m) by building a hash-map.
// Call ONCE per request, use the map for all lookups.

// 1-to-1 index: keyField → single row
function idx(rowArr, keyField) {
  var m = {};
  rowArr.forEach(function(r) {
    var k = String(r[keyField]||'');
    if (k) m[k] = r;
  });
  return m;
}

// 1-to-many index: keyField → array of rows
function midx(rowArr, keyField) {
  var m = {};
  rowArr.forEach(function(r) {
    var k = String(r[keyField]||'');
    if (!m[k]) m[k] = [];
    m[k].push(r);
  });
  return m;
}

// ============================================================
//  SERVER-SIDE CACHE  (_Cache sheet, TTL-based KV)
//  Used for: getDashboard, getStudents — the two heaviest reads
// ============================================================
function cacheGet(key) {
  try {
    var data = rows(SH.CACHE);
    for (var i=0; i<data.length; i++) {
      if (data[i].key === key) {
        if (new Date(data[i].expires_at).getTime() > Date.now())
          return JSON.parse(data[i].value);
        return null;
      }
    }
  } catch(e) {}
  return null;
}

function cacheSet(key, value) {
  try {
    var sheet   = getSheet(SH.CACHE);
    var raw     = sheet.getDataRange().getValues();
    var hdrs    = raw[0];
    var kCol    = hdrs.indexOf('key')+1;
    var vCol    = hdrs.indexOf('value')+1;
    var eCol    = hdrs.indexOf('expires_at')+1;
    var exp     = new Date(Date.now()+CACHE_TTL_MS).toISOString();
    var serial  = JSON.stringify(value);
    for (var i=1; i<raw.length; i++) {
      if (raw[i][kCol-1]===key) {
        sheet.getRange(i+1,vCol).setValue(serial);
        sheet.getRange(i+1,eCol).setValue(exp);
        bustCache(SH.CACHE); return;
      }
    }
    sheet.appendRow([key,serial,exp]);
    bustCache(SH.CACHE);
  } catch(e) {}
}

function cacheDel(key) {
  try {
    var sheet = getSheet(SH.CACHE);
    var raw   = sheet.getDataRange().getValues();
    var kCol  = raw[0].indexOf('key')+1;
    for (var i=1; i<raw.length; i++) {
      if (raw[i][kCol-1]===key) { sheet.deleteRow(i+1); bustCache(SH.CACHE); return; }
    }
  } catch(e) {}
}

// ============================================================
//  ENTRY POINTS
// ============================================================
function doGet(e) {
  _RC = {};
  try {
    var p = (e&&e.parameter) ? e.parameter : {};
    if (!p.data) return jsonOut({status:'BioAttend API v5',time:new Date().toString()});
    return jsonOut(route(JSON.parse(decodeURIComponent(p.data))));
  } catch(err) { return jsonOut({success:false,message:'doGet: '+err}); }
}

function doPost(e) {
  _RC = {};
  try {
    return jsonOut(route(JSON.parse(e.postData.contents)));
  } catch(err) { return jsonOut({success:false,message:'doPost: '+err}); }
}

function route(b) {
  switch(b.action) {
    // Auth & identity
    case 'register':         return registerUser(b);
    case 'signIn':           return signInUser(b);
    case 'saveBiometric':    return saveBiometric(b);
    case 'getBiometric':     return getBiometric(b);
    case 'registerDevice':   return registerDevice(b);
    case 'checkDevice':      return checkDevice(b);
    // Attendance
    case 'markAttendance':   return markAttendance(b);
    case 'markExit':         return markExit(b);
    // Sessions
    case 'createSession':    return createSession(b);
    case 'closeSession':     return closeSession(b);
    case 'getActiveSession': return getActiveSession(b);
    case 'getSessions':      return getSessions(b);
    // Reporting
    case 'getDashboard':     return getDashboard(b);
    case 'getAttendance':    return getAttendance(b);
    case 'getStudents':      return getStudents(b);
    case 'exportAttendance': return exportAttendance(b);
    // Reference data
    case 'getDepartments':   return {success:true, departments:rows(SH.DEPTS)};
    case 'getRoles':         return {success:true, roles:rows(SH.ROLES)};
    case 'getLocations':     return getLocations();
    // Admin / setup
    case 'setupSheets':      return setupSheets();
    case 'seedMasterData':   return seedMasterData();
    case 'debug':            return debugInfo();
    default: return {success:false, message:'Unknown action: '+b.action};
  }
}

// ============================================================
//  REGISTER
// ============================================================
function registerUser(b) {
  try {
    if (!b.name||!b.email||!b.password)
      return {success:false,message:'Name, email and password required'};

    var users  = rows(SH.USERS);
    var eLower = b.email.toLowerCase();
    for (var i=0; i<users.length; i++)
      if (String(users[i].email||'').toLowerCase()===eLower)
        return {success:false,message:'Email already registered'};

    // Resolve dept_id — accept id directly or resolve by name
    var deptId = b.dept_id || '';
    if (!deptId && b.department) {
      var depts = rows(SH.DEPTS);
      for (var d=0; d<depts.length; d++) {
        if (String(depts[d].name||'').toLowerCase()===String(b.department).toLowerCase()) {
          deptId = depts[d].department_id; break;
        }
      }
    }

    // Resolve role_id — accept id directly or resolve by name (default: student)
    var roleId = b.role_id || _roleId(b.role||'student');

    var userId = uid('u');
    getSheet(SH.USERS).appendRow([
      userId, deptId, roleId,
      b.name, b.dob||'', b.mobile||'', b.email,
      hashPw(b.password), '', '', new Date().toISOString()
    ]);
    bustCache(SH.USERS);
    cacheDel('students_list');

    return {success:true, userId:userId, message:'Account created'};
  } catch(err) { return {success:false,message:'register: '+err}; }
}

// ============================================================
//  SIGN IN
// ============================================================
function signInUser(b) {
  try {
    var userRows = rows(SH.USERS);
    var hash     = hashPw(b.password||'');
    var eLower   = String(b.email||'').toLowerCase();
    for (var i=0; i<userRows.length; i++) {
      var u = userRows[i];
      if (String(u.email||'').toLowerCase()===eLower && u.password_hash===hash) {
        return {
          success: true,
          userId:  u.user_id,
          name:    u.full_name,
          role:    _roleName(u.role_id),
          dept_id: u.dept_id
        };
      }
    }
    return {success:false, message:'Invalid email or password'};
  } catch(err) { return {success:false,message:'signIn: '+err}; }
}

// ── Role helpers ──────────────────────────────────────────────
function _roleName(roleId) {
  var r = rows(SH.ROLES);
  for (var i=0; i<r.length; i++) if (r[i].role_id===roleId) return r[i].name;
  return 'student';
}
function _roleId(name) {
  var r = rows(SH.ROLES);
  var n = String(name||'').toLowerCase();
  for (var i=0; i<r.length; i++) if (String(r[i].name||'').toLowerCase()===n) return r[i].role_id;
  return '';
}

// ============================================================
//  MARK ATTENDANCE  (entry + geofence)
//
//  Geofence priority:
//  1. UserLocationMap override for this user+location
//  2. AttendanceLocations.default_allowed_distance_m
//  3. Hardcoded 200m fallback
// ============================================================
function markAttendance(b) {
  try {
    if (!b.userId) return {success:false,message:'userId required'};

    var now     = new Date();
    var tz      = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    // A. User lookup via index  O(1)
    var userIdx = idx(rows(SH.USERS), 'user_id');
    var user    = userIdx[b.userId];
    if (!user) return {success:false,message:'User not found'};

    // B. Resolve location
    var loc  = _resolveLoc(b.location_id||null);
    var lat  = (b.lat!=null&&b.lat!=='') ? parseFloat(b.lat) : null;
    var lng  = (b.lng!=null&&b.lng!=='') ? parseFloat(b.lng) : null;
    var dist = '';

    // C. Geofence check (only when GPS provided)
    if (lat!==null && !isNaN(lat)) {
      dist = haversine(lat, lng, parseFloat(loc.latitude), parseFloat(loc.longitude));
      var allowed = _allowedDist(b.userId, loc.location_id, parseInt(loc.default_allowed_distance_m)||200);
      if (dist > allowed) {
        return {
          success:  false,
          code:     'TOO_FAR',
          distance: dist,
          allowed:  allowed,
          message:  'You are '+dist+'m from '+loc.name+'. Must be within '+allowed+'m.'
        };
      }
    }

    // D. Duplicate check — O(n) single pass using multi-index
    var attByUser = midx(rows(SH.ATT), 'user_id');
    var todayAtt  = (attByUser[b.userId]||[]).filter(function(a) {
      return normDate(a.att_date, tz)===dateStr;
    });
    if (todayAtt.length)
      return {success:false, message:'Attendance already marked today at '+(todayAtt[0].entry_time||'')};

    // E. Resolve att_type_id for "entry"
    var attTypeId = _attTypeId('entry');

    // F. Write lean Attendance row — NO full_name, NO address, NO GPS string
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    var attId   = uid('att');
    getSheet(SH.ATT).appendRow([
      attId,
      b.sessionId||'',
      b.userId,
      loc.location_id||'',
      attTypeId,
      dateStr,
      timeStr,
      '',           // exit_time
      b.method||'biometric',
      lat!==null ? String(lat) : '',
      lng!==null ? String(lng) : '',
      dist!=='' ? dist : '',
      '','','',''   // exit_lat, exit_lng, exit_distance_m, duration_mins
    ]);
    bustCache(SH.ATT);
    cacheDel('dashboard_'+(b.sessionId||''));

    // G. Write GPS audit log to LocationMonitor (separate table, keeps Attendance lean)
    try {
      if (lat!==null) {
        getSheet(SH.LOC_MON).appendRow([
          uid('mon'), b.userId, loc.location_id||'',
          lat, lng, dist, b.address||'', now.toISOString()
        ]);
      }
    } catch(e) {}

    return {
      success:        true,
      message:        '✓ Attendance marked at '+timeStr,
      name:           user.full_name,
      date:           dateStr,
      time:           timeStr,
      location:       b.address||(lat!==null ? lat+', '+lng : 'not captured'),
      gps:            lat!==null ? lat+', '+lng : '',
      distanceMeters: dist
    };
  } catch(err) { return {success:false,message:'markAttendance: '+err}; }
}

// Resolve location row; fallback to default
function _resolveLoc(locId) {
  var locs = rows(SH.LOCS);
  if (locId) {
    for (var i=0; i<locs.length; i++) if (locs[i].location_id===locId) return locs[i];
  }
  if (locs.length) return locs[0];
  return DEFAULT_LOC;
}

// Resolve effective allowed distance for a user+location (UserLocationMap override)
function _allowedDist(userId, locationId, defaultDist) {
  var maps = rows(SH.USR_LOC);
  for (var i=0; i<maps.length; i++) {
    var m = maps[i];
    if (m.user_id===userId && m.location_id===locationId && m.allowed_distance_m) {
      return parseInt(m.allowed_distance_m);
    }
  }
  return defaultDist || 200;
}

// Resolve att_type_id for a given type name
function _attTypeId(typeName) {
  var types = rows(SH.ATT_TYPES);
  var n = String(typeName||'entry').toLowerCase();
  for (var i=0; i<types.length; i++)
    if (String(types[i].type_name||'').toLowerCase()===n) return types[i].att_type_id;
  return 'att_type_entry';
}

// ============================================================
//  MARK EXIT
// ============================================================
function markExit(b) {
  try {
    if (!b.userId) return {success:false,message:'userId required'};

    var now     = new Date();
    var tz      = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    var sheet   = getSheet(SH.ATT);
    var data    = sheet.getDataRange().getValues();
    var hdrs    = data[0];

    function c(n) { var i=hdrs.indexOf(n); return i===-1?-1:i+1; }

    var uCol    = c('user_id'),     dCol  = c('att_date');
    var etCol   = c('entry_time'),  xTCol = c('exit_time');
    var xLaCol  = c('exit_lat'),    xLnCol= c('exit_lng');
    var xDCol   = c('exit_distance_m'), durCol = c('duration_mins');
    var locIdCol= c('location_id');

    if (uCol<1||dCol<1||xTCol<1)
      return {success:false,message:'Sheet missing columns. Run setupSheets.'};

    for (var i=1; i<data.length; i++) {
      if (String(data[i][uCol-1]||'').trim()!==String(b.userId).trim()) continue;
      if (normDate(data[i][dCol-1],tz)!==dateStr) continue;

      if (String(data[i][xTCol-1]||'').trim())
        return {success:false,message:'Exit already recorded at '+data[i][xTCol-1]};

      // Duration in minutes
      var durMins='';
      var et = String(data[i][etCol>0?etCol-1:-1]||'').trim();
      if (et) {
        var ed = new Date(dateStr+'T'+et);
        if (!isNaN(ed.getTime())) durMins = Math.max(0,Math.round((now-ed)/60000));
      }

      // Exit GPS + distance
      var xLat=(b.lat!=null&&b.lat!=='') ? parseFloat(b.lat) : null;
      var xLng=(b.lng!=null&&b.lng!=='') ? parseFloat(b.lng) : null;
      var xDist='';
      if (xLat!==null) {
        var rowLoc = _resolveLoc(locIdCol>0 ? data[i][locIdCol-1] : null);
        xDist = haversine(xLat,xLng,parseFloat(rowLoc.latitude),parseFloat(rowLoc.longitude));
      }

      sheet.getRange(i+1,xTCol).setValue(timeStr);
      if (xLaCol>0) sheet.getRange(i+1,xLaCol).setValue(xLat!==null?String(xLat):'');
      if (xLnCol>0) sheet.getRange(i+1,xLnCol).setValue(xLng!==null?String(xLng):'');
      if (xDCol>0)  sheet.getRange(i+1,xDCol).setValue(xDist);
      if (durCol>0) sheet.getRange(i+1,durCol).setValue(durMins);
      SpreadsheetApp.flush();
      bustCache(SH.ATT);

      // GPS audit log for exit
      try {
        if (xLat!==null) {
          getSheet(SH.LOC_MON).appendRow([
            uid('mon'), b.userId,
            locIdCol>0?data[i][locIdCol-1]:'',
            xLat, xLng, xDist, b.address||'', now.toISOString()
          ]);
        }
      } catch(e) {}

      var durLabel = durMins!==''
        ? (Math.floor(durMins/60)>0 ? Math.floor(durMins/60)+'h '+(durMins%60)+'m' : durMins+'m')
        : '';

      return {
        success:        true,
        message:        '✓ Exit at '+timeStr+(durLabel?' · '+durLabel:''),
        exitTime:       timeStr,
        duration:       durLabel,
        location:       b.address||(xLat!==null?xLat+', '+xLng:'not captured'),
        gps:            xLat!==null?xLat+', '+xLng:'',
        distanceMeters: xDist
      };
    }

    // Not found — surface debug info
    var knownDates=[];
    for (var d2=1; d2<data.length; d2++) {
      if (String(data[d2][uCol-1]||'').trim()===String(b.userId).trim())
        knownDates.push(normDate(data[d2][dCol-1],tz));
    }
    return {
      success: false,
      message: 'No attendance found for today ('+dateStr+'). '+
               (knownDates.length ? 'Recorded dates: '+knownDates.join(', ') : 'No records found.'),
      debug: 'tz='+tz
    };
  } catch(err) { return {success:false,message:'markExit: '+err}; }
}

// ============================================================
//  SESSIONS
// ============================================================
function createSession(b) {
  try {
    var roleName = b.role || _roleName(b.role_id||'');
    if (roleName!=='teacher'&&roleName!=='admin')
      return {success:false,message:'Only teachers can create sessions'};
    if (!b.subject||!b.windowMinutes)
      return {success:false,message:'Subject and window required'};

    var now     = new Date();
    var tz      = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now,tz,'yyyy-MM-dd');
    var start   = Utilities.formatDate(now,tz,'HH:mm:ss');
    var end     = Utilities.formatDate(new Date(now.getTime()+parseInt(b.windowMinutes)*60000),tz,'HH:mm:ss');
    var sessId  = uid('sess');
    var locId   = b.location_id || DEFAULT_LOC.location_id;

    // Close open sessions for this teacher
    var sheet = getSheet(SH.SESSIONS);
    var data  = sheet.getDataRange().getValues();
    var hdrs  = data[0];
    var tCol  = hdrs.indexOf('teacher_id')+1;
    var sCol  = hdrs.indexOf('status')+1;
    for (var i=1; i<data.length; i++)
      if (data[i][tCol-1]===b.userId&&data[i][sCol-1]==='open')
        sheet.getRange(i+1,sCol).setValue('closed');

    sheet.appendRow([sessId,b.userId,locId,b.subject,dateStr,start,end,b.windowMinutes,'open']);
    bustCache(SH.SESSIONS);

    return {success:true,sessionId:sessId,subject:b.subject,startTime:start,endTime:end,
            message:'Session opened for '+b.windowMinutes+' min'};
  } catch(err) { return {success:false,message:'createSession: '+err}; }
}

function closeSession(b) {
  try {
    var sheet = getSheet(SH.SESSIONS);
    var data  = sheet.getDataRange().getValues();
    var hdrs  = data[0];
    var sidCol = hdrs.indexOf('session_id')+1;
    var stCol  = hdrs.indexOf('status')+1;
    for (var i=1; i<data.length; i++) {
      if (data[i][sidCol-1]===b.sessionId) {
        sheet.getRange(i+1,stCol).setValue('closed');
        bustCache(SH.SESSIONS);
        return {success:true};
      }
    }
    return {success:false,message:'Session not found'};
  } catch(err) { return {success:false,message:'closeSession: '+err}; }
}

function getActiveSession() {
  try {
    var sesRows = rows(SH.SESSIONS);
    var now     = new Date();
    var tz      = Session.getScriptTimeZone();
    var today   = Utilities.formatDate(now,tz,'yyyy-MM-dd');
    for (var i=0; i<sesRows.length; i++) {
      var s = sesRows[i];
      if (s.status!=='open'||normDate(s.date,tz)!==today) continue;
      var st = new Date(today+'T'+s.start_time);
      var en = new Date(today+'T'+s.end_time);
      if (now>=st&&now<=en)
        return {success:true,active:true,session:s,secondsLeft:Math.max(0,Math.round((en-now)/1000))};
    }
    return {success:true,active:false};
  } catch(err) { return {success:false,message:'getActiveSession: '+err}; }
}

function getSessions(b) {
  try {
    var sesRows = rows(SH.SESSIONS);
    if (b.userId) sesRows = sesRows.filter(function(r){return r.teacher_id===b.userId;});

    // Count present per session using multi-index → O(n+m) not O(n*m)
    var attBySess = midx(rows(SH.ATT), 'session_id');
    sesRows.forEach(function(s){
      s.presentCount = (attBySess[s.session_id]||[]).length;
    });

    sesRows.sort(function(a,b){return String(b.date).localeCompare(String(a.date));});
    return {success:true, sessions:sesRows.slice(0,30)};
  } catch(err) { return {success:false,message:'getSessions: '+err}; }
}

// ============================================================
//  DASHBOARD  (30-second server cache)
// ============================================================
function getDashboard(b) {
  try {
    if (!b.sessionId) return {success:false,message:'sessionId required'};
    var cKey = 'dashboard_'+b.sessionId;
    var hit  = cacheGet(cKey);
    if (hit) return hit;

    // Build all indexes in one pass each
    var userRows  = rows(SH.USERS);
    var deptRows  = rows(SH.DEPTS);
    var attRows   = rows(SH.ATT).filter(function(a){return a.session_id===b.sessionId;});

    var userIndex = idx(userRows, 'user_id');
    var deptIndex = idx(deptRows, 'department_id');
    var attIndex  = idx(attRows,  'user_id');

    var stuRoleId = _roleId('student');
    var students  = userRows.filter(function(u){return u.role_id===stuRoleId;});

    var present=[], absent=[];
    students.forEach(function(s) {
      var dept = deptIndex[s.dept_id]||{};
      var a    = attIndex[s.user_id];
      if (a) {
        present.push({
          userId:         s.user_id,
          name:           s.full_name,
          email:          s.email,
          department:     dept.name||'',
          time:           a.entry_time,
          method:         a.login_method,
          distanceMeters: a.entry_distance_m||''
        });
      } else {
        absent.push({userId:s.user_id,name:s.full_name,email:s.email,department:dept.name||''});
      }
    });

    var result = {
      success:true, total:students.length,
      presentCount:present.length, absentCount:absent.length,
      present:present, absent:absent
    };
    cacheSet(cKey, result);
    return result;
  } catch(err) { return {success:false,message:'getDashboard: '+err}; }
}

// ============================================================
//  GET STUDENTS  (30-second server cache)
// ============================================================
function getStudents() {
  try {
    var hit = cacheGet('students_list');
    if (hit) return hit;

    var userRows  = rows(SH.USERS);
    var deptRows  = rows(SH.DEPTS);
    var deptIndex = idx(deptRows, 'department_id');
    var stuRoleId = _roleId('student');

    var list = userRows
      .filter(function(u){return u.role_id===stuRoleId;})
      .map(function(u){
        var dept = deptIndex[u.dept_id]||{};
        return {
          userId:               u.user_id,
          FullName:             u.full_name,
          Email:                u.email,
          Department:           dept.name||'',
          dept_id:              u.dept_id,
          DeviceId:             u.device_identification||'',
          BiometricCredentialId: u.biometric_code||''
        };
      });

    var result = {success:true, students:list};
    cacheSet('students_list', result);
    return result;
  } catch(err) { return {success:false,message:'getStudents: '+err}; }
}

// ============================================================
//  GET ATTENDANCE  (for history; enriched via joins)
// ============================================================
function getAttendance(b) {
  try {
    var tz      = Session.getScriptTimeZone();
    var attRows = rows(SH.ATT);
    var uIdx    = idx(rows(SH.USERS),    'user_id');
    var sIdx    = idx(rows(SH.SESSIONS), 'session_id');

    if (b.sessionId) attRows = attRows.filter(function(r){return r.session_id===b.sessionId;});
    if (b.date)      attRows = attRows.filter(function(r){return normDate(r.att_date,tz)===b.date;});

    return {
      success:true,
      attendance: attRows.map(function(r){
        var u=uIdx[r.user_id]||{};
        var s=sIdx[r.session_id]||{};
        return {
          attendance_id:    r.attendance_id,
          user_id:          r.user_id,
          full_name:        u.full_name||'',
          email:            u.email||'',
          subject:          s.subject||'',
          att_date:         r.att_date,
          entry_time:       r.entry_time,
          exit_time:        r.exit_time,
          login_method:     r.login_method,
          entry_distance_m: r.entry_distance_m,
          duration_mins:    r.duration_mins
        };
      })
    };
  } catch(err) { return {success:false,message:'getAttendance: '+err}; }
}

// ============================================================
//  EXPORT CSV
// ============================================================
function exportAttendance(b) {
  try {
    var tz      = Session.getScriptTimeZone();
    var attRows = rows(SH.ATT);
    var uIdx    = idx(rows(SH.USERS),    'user_id');
    var dIdx    = idx(rows(SH.DEPTS),    'department_id');
    var sIdx    = idx(rows(SH.SESSIONS), 'session_id');

    if (b.sessionId) attRows = attRows.filter(function(r){return r.session_id===b.sessionId;});
    if (b.date)      attRows = attRows.filter(function(r){return normDate(r.att_date,tz)===b.date;});

    var hdr = ['Name','Email','Department','Date','Entry Time','Exit Time',
               'Duration(mins)','Subject','Method','Distance(m)'];
    var lines = [hdr.join(',')];
    attRows.forEach(function(r){
      var u=uIdx[r.user_id]||{};
      var d=dIdx[u.dept_id]||{};
      var s=sIdx[r.session_id]||{};
      lines.push([
        '"'+(u.full_name||'')+'"', '"'+(u.email||'')+'"',
        '"'+(d.name||'')+'"', r.att_date||'',
        r.entry_time||'', r.exit_time||'', r.duration_mins||'',
        '"'+(s.subject||'')+'"', r.login_method||'',
        r.entry_distance_m||''
      ].join(','));
    });
    return {success:true, csv:lines.join('\n'), rowCount:attRows.length};
  } catch(err) { return {success:false,message:'exportAttendance: '+err}; }
}

// ============================================================
//  LOCATIONS
// ============================================================
function getLocations() {
  try {
    var locs = rows(SH.LOCS);
    return {success:true, locations:locs.length ? locs : [DEFAULT_LOC]};
  } catch(err) { return {success:false,message:'getLocations: '+err}; }
}

// ============================================================
//  BIOMETRIC
// ============================================================
function saveBiometric(b) {
  try {
    var sheet = getSheet(SH.USERS);
    var data  = sheet.getDataRange().getValues();
    var hdrs  = data[0];
    var uCol  = hdrs.indexOf('user_id')+1;
    var bCol  = hdrs.indexOf('biometric_code')+1;
    for (var i=1; i<data.length; i++) {
      if (data[i][uCol-1]===b.userId) {
        sheet.getRange(i+1,bCol).setValue(b.credentialId);
        bustCache(SH.USERS);
        return {success:true};
      }
    }
    return {success:false,message:'User not found'};
  } catch(err) { return {success:false,message:'saveBiometric: '+err}; }
}

function getBiometric(b) {
  try {
    var userRows = rows(SH.USERS);
    var eLower   = String(b.email||'').toLowerCase();
    for (var i=0; i<userRows.length; i++) {
      var u = userRows[i];
      if (String(u.email||'').toLowerCase()===eLower) {
        if (!u.biometric_code) return {success:false,message:'No biometric registered'};
        return {success:true, credentialId:u.biometric_code, userId:u.user_id, name:u.full_name};
      }
    }
    return {success:false,message:'User not found'};
  } catch(err) { return {success:false,message:'getBiometric: '+err}; }
}

// ============================================================
//  DEVICE BINDING
// ============================================================
function registerDevice(b) {
  try {
    if (!b.userId||!b.deviceId) return {success:false,message:'userId and deviceId required'};
    var sheet = getSheet(SH.USERS);
    var data  = sheet.getDataRange().getValues();
    var hdrs  = data[0];
    var uCol  = hdrs.indexOf('user_id')+1;
    var dCol  = hdrs.indexOf('device_identification')+1;
    for (var i=1; i<data.length; i++) {
      if (data[i][uCol-1]===b.userId) {
        var existing = String(data[i][dCol-1]||'').trim();
        if (existing&&existing!==String(b.deviceId).trim())
          return {success:false,alreadyBound:true,message:'Account already bound to another device.'};
        sheet.getRange(i+1,dCol).setValue(b.deviceId);
        bustCache(SH.USERS);
        return {success:true,alreadyBound:false,message:existing?'Device confirmed':'Device registered'};
      }
    }
    return {success:false,message:'User not found'};
  } catch(err) { return {success:false,message:'registerDevice: '+err}; }
}

function checkDevice(b) {
  try {
    if (!b.userId||!b.deviceId) return {success:false,message:'userId and deviceId required'};
    var uIdx = idx(rows(SH.USERS),'user_id');
    var u    = uIdx[b.userId];
    if (!u) return {success:false,message:'User not found'};
    var stored = String(u.device_identification||'').trim();
    if (!stored) return {success:true,status:'unbound'};
    if (stored===String(b.deviceId).trim()) return {success:true,status:'match'};
    return {success:false,status:'mismatch',message:'Account registered to a different device.'};
  } catch(err) { return {success:false,message:'checkDevice: '+err}; }
}

// ============================================================
//  SETUP — run once from Apps Script editor
// ============================================================
function setupSheets() {
  try {
    Object.keys(SCHEMA).forEach(function(n){ getSheet(n); });
    return {success:true, message:'All sheets ready: '+Object.keys(SCHEMA).join(', ')};
  } catch(err) { return {success:false,message:'setupSheets: '+err}; }
}

// Seed Roles, Departments, AttendanceTypes, default location  (run once)
function seedMasterData() {
  try {
    // Roles
    if (!rows(SH.ROLES).length) {
      var rs = getSheet(SH.ROLES);
      rs.appendRow(['role_student','student']);
      rs.appendRow(['role_teacher','teacher']);
      rs.appendRow(['role_admin',  'admin']);
      bustCache(SH.ROLES);
    }
    // AttendanceTypes
    if (!rows(SH.ATT_TYPES).length) {
      var ts = getSheet(SH.ATT_TYPES);
      ts.appendRow(['att_type_entry','entry']);
      ts.appendRow(['att_type_exit', 'exit']);
      bustCache(SH.ATT_TYPES);
    }
    // Departments
    if (!rows(SH.DEPTS).length) {
      var ds = getSheet(SH.DEPTS);
      [['dept_cse','Computer Science & Engineering','',''],
       ['dept_ise','Information Science & Engineering','',''],
       ['dept_ece','Electronics & Communication','',''],
       ['dept_eee','Electrical & Electronics','',''],
       ['dept_mech','Mechanical Engineering','',''],
       ['dept_civil','Civil Engineering','',''],
       ['dept_chem','Chemical Engineering','',''],
       ['dept_bio','Biotechnology','',''],
       ['dept_mba','MBA','',''],
       ['dept_mca','MCA','','']
      ].forEach(function(d){ds.appendRow(d);});
      bustCache(SH.DEPTS);
    }
    // Default location
    if (!rows(SH.LOCS).length) {
      getSheet(SH.LOCS).appendRow([
        DEFAULT_LOC.location_id, DEFAULT_LOC.name,
        DEFAULT_LOC.latitude, DEFAULT_LOC.longitude,
        DEFAULT_LOC.default_allowed_distance_m
      ]);
      bustCache(SH.LOCS);
    }
    return {success:true,message:'Master data seeded: Roles, AttendanceTypes, Departments, Location'};
  } catch(err) { return {success:false,message:'seedMasterData: '+err}; }
}

// ============================================================
//  DEBUG
// ============================================================
function debugInfo() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var info = {};
    Object.keys(SCHEMA).forEach(function(n){
      var s = ss.getSheetByName(n);
      info[n] = s
        ? {rows:s.getLastRow()-1, cols:s.getLastColumn(),
           headers:s.getLastColumn()>0?s.getRange(1,1,1,s.getLastColumn()).getValues()[0]:[]}
        : 'MISSING';
    });
    return {success:true, sheets:info, tz:Session.getScriptTimeZone()};
  } catch(err) { return {success:false,message:'debug: '+err}; }
}