// ============================================================
//  BioAttend – Google Apps Script Backend  v5
//  College  : Siddaganga Institute of Technology, Tumkur
//  Schema   : Optimized 7-sheet design (normalised, no redundancy)
//  Deploy   : Web App → Execute as Me → Access: Anyone
// ============================================================

// ── Sheet names ───────────────────────────────────────────────
var SH_USERS       = 'Users';
var SH_ATTENDANCE  = 'Attendance';
var SH_SESSIONS    = 'Sessions';
var SH_DEPARTMENTS = 'Departments';
var SH_ROLES       = 'Roles';
var SH_LOCATIONS   = 'AttendanceLocations';
var SH_USER_INDEX  = 'UserIndex';   // simulated index — row pointer map

// ── Column constants (1-indexed) ──────────────────────────────
// Keep these in sync with sheet headers — never hardcode numbers elsewhere.
var UC = { ID:1, NAME:2, EMAIL:3, HASH:4, DOB:5, MOBILE:6, INST:7, DEPT:8, ROLE:9, BIO:10, DEVICE:11 };
var AC = { ID:1, UID:2, LOC:3, DATE:4, ETIME:5, ELAT:6, ELNG:7, EADDR:8, EDIST:9, METHOD:10,
           XTIME:11, XLAT:12, XLNG:13, XADDR:14, XDIST:15, DURATION:16, TIMESTAMP:17 };
var SC = { ID:1, TID:2, TNAME:3, SUBJ:4, DATE:5, START:6, END:7, STATUS:8, WINDOW:9 };

// ── GPS anchor (SIT campus lab) ───────────────────────────────
var LAB_LAT = 13.32603;
var LAB_LNG = 77.12621;
var GEOFENCE_M = 200;

// ── Cache TTLs (seconds) ──────────────────────────────────────
var TTL_LOOKUP   = 600;  // Departments, Roles, Locations — 10 min
var TTL_DASHBOARD = 60;  // getDashboard result — 1 min
var TTL_SESSION   = 30;  // active session — 30s

// ============================================================
//  TRANSPORT
// ============================================================

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (!p.data) return jsonOut({ status:'BioAttend v5', time:new Date().toString() });
    return jsonOut(route(JSON.parse(decodeURIComponent(p.data))));
  } catch(err) { return jsonOut({ success:false, message:'doGet: '+err }); }
}

function doPost(e) {
  try { return jsonOut(route(JSON.parse(e.postData.contents))); }
  catch(err) { return jsonOut({ success:false, message:'doPost: '+err }); }
}

// ============================================================
//  ROUTER
// ============================================================

function route(b) {
  switch(b.action) {
    // Auth
    case 'register':           return registerUser(b);
    case 'signIn':             return signInUser(b);
    case 'getBiometric':       return getBiometric(b);
    case 'saveBiometric':      return saveBiometric(b);
    case 'registerDevice':     return registerDevice(b);
    case 'checkDevice':        return checkDevice(b);
    // Attendance
    case 'markEntry':          return markEntry(b);
    case 'markExit':           return markExit(b);
    case 'getMyAttendance':    return getMyAttendance(b);
    case 'exportAttendance':   return exportAttendance(b);
    // Sessions
    case 'createSession':      return createSession(b);
    case 'closeSession':       return closeSession(b);
    case 'getActiveSession':   return getActiveSession(b);
    case 'getSessions':        return getSessions(b);
    // Dashboard
    case 'getDashboard':       return getDashboard(b);
    case 'getStudents':        return getStudents(b);
    // Admin (gated)
    case 'setupSheets':        return setupSheets();
    case 'debug':              return debugInfo(b);
    default: return { success:false, message:'Unknown action: '+b.action };
  }
}

// ============================================================
//  SHEET HELPERS
// ============================================================

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name) {
  var sheet = ss().getSheetByName(name);
  if (!sheet) {
    sheet = ss().insertSheet(name);
    var HEADERS = {
      Users: [
        'user_id','full_name','email','password_hash',
        'dob','mobile','institute_id','department_id','role_id',
        'biometric_credential_id','device_id'
      ],
      Attendance: [
        'attendance_id','user_id','location_id',
        'date','entry_time','entry_lat','entry_lng',
        'entry_address','entry_distance_m','login_method',
        'exit_time','exit_lat','exit_lng','exit_address',
        'exit_distance_m','duration_mins','entry_timestamp'
      ],
      Sessions: [
        'session_id','teacher_id','teacher_name',
        'subject','date','start_time','end_time','status','window_minutes'
      ],
      Departments: ['department_id','name','in_charge','email'],
      Roles:       ['role_id','name'],
      AttendanceLocations: ['location_id','name','latitude','longitude','allowed_radius_m'],
      UserIndex:   ['user_id','email','row_number']
    };
    if (HEADERS[name]) {
      var h = HEADERS[name];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length)
        .setFontWeight('bold').setBackground('#0f172a').setFontColor('#f1f5f9');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// getRows — maps rows to objects by header name
// Only reads filled rows (getLastRow instead of getDataRange)
function getRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var data    = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var o = {};
    headers.forEach(function(h,i){ o[h] = row[i]; });
    return o;
  });
}

// getCachedRows — CacheService wrapper for read-heavy lookup sheets
function getCachedRows(sheetName, ttl) {
  var cache = CacheService.getScriptCache();
  var key   = 'rows_' + sheetName;
  var hit   = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch(e){} }
  var rows = getRows(getSheet(sheetName));
  try { cache.put(key, JSON.stringify(rows), ttl || TTL_LOOKUP); } catch(e){}
  return rows;
}

function invalidateCache(sheetName) {
  try { CacheService.getScriptCache().remove('rows_' + sheetName); } catch(e){}
}

// buildMap — O(1) lookup map from array of objects
function buildMap(rows, keyField) {
  var m = {};
  rows.forEach(function(r){ m[r[keyField]] = r; });
  return m;
}

// ── Simulated index: UserIndex sheet ─────────────────────────
// 3 columns: user_id | email | row_number (in Users sheet)
// Written on register, used on every login — avoids full Users scan.

function getUserByEmail(email) {
  var idxSheet = getSheet(SH_USER_INDEX);
  var idxData  = idxSheet.getDataRange().getValues();
  var emailLow = String(email).toLowerCase();
  for (var i = 1; i < idxData.length; i++) {
    if (String(idxData[i][1]).toLowerCase() === emailLow) {
      var rowNum = idxData[i][2];
      var cols   = Object.keys(UC).length;
      // Direct fetch — no loop through all users
      var uRow = getSheet(SH_USERS).getRange(rowNum, 1, 1, cols).getValues()[0];
      var headers = getSheet(SH_USERS).getRange(1, 1, 1, cols).getValues()[0];
      var obj = {};
      headers.forEach(function(h,i){ obj[h] = uRow[i]; });
      return obj;
    }
  }
  return null;
}

function getUserById(userId) {
  var idxSheet = getSheet(SH_USER_INDEX);
  var idxData  = idxSheet.getDataRange().getValues();
  for (var i = 1; i < idxData.length; i++) {
    if (String(idxData[i][0]) === String(userId)) {
      var rowNum  = idxData[i][2];
      var cols    = Object.keys(UC).length;
      var uRow    = getSheet(SH_USERS).getRange(rowNum, 1, 1, cols).getValues()[0];
      var headers = getSheet(SH_USERS).getRange(1, 1, 1, cols).getValues()[0];
      var obj = {};
      headers.forEach(function(h,i){ obj[h] = uRow[i]; });
      return obj;
    }
  }
  return null;
}

function addToUserIndex(userId, email) {
  var userSheet = getSheet(SH_USERS);
  var rowNum    = userSheet.getLastRow(); // just appended
  getSheet(SH_USER_INDEX).appendRow([userId, email, rowNum]);
}

// ── Helpers ───────────────────────────────────────────────────

function generateId(prefix) {
  return (prefix||'id')+'_'+Date.now()+'_'+Math.random().toString(36).substr(2,5);
}

function hashPw(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
}

function haversine(lat1, lng1, lat2, lng2) {
  var R=6371000, d2r=Math.PI/180;
  var dLat=(lat2-lat1)*d2r, dLng=(lng2-lng1)*d2r;
  var a=Math.sin(dLat/2)*Math.sin(dLat/2)+
        Math.cos(lat1*d2r)*Math.cos(lat2*d2r)*
        Math.sin(dLng/2)*Math.sin(dLng/2);
  return Math.round(R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a)));
}

function normDate(raw, tz) {
  if (!raw && raw!==0) return '';
  if (raw instanceof Date) return Utilities.formatDate(raw,tz,'yyyy-MM-dd');
  var s=String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (s.length>10&&s.indexOf('T')>0) return s.substring(0,10);
  var d=new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d,tz,'yyyy-MM-dd');
  return s.substring(0,10);
}

// ============================================================
//  1. REGISTER
// ============================================================

function registerUser(b) {
  try {
    if (!b.name||!b.email||!b.password)
      return { success:false, message:'Name, email and password required' };

    // Check duplicate via UserIndex (fast path)
    var existing = getUserByEmail(b.email);
    if (existing) return { success:false, message:'Email already registered' };

    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var userId = generateId('u');
      var role   = b.role || 'student';
      var sheet  = getSheet(SH_USERS);
      sheet.appendRow([
        userId,               // user_id
        b.name,               // full_name
        b.email,              // email
        hashPw(b.password),   // password_hash
        b.dob      || '',     // dob
        b.mobile   || '',     // mobile
        'SIT001',             // institute_id
        b.department || '',   // department_id
        role,                 // role_id
        '',                   // biometric_credential_id
        b.deviceId || ''      // device_id
      ]);
      // Update UserIndex
      addToUserIndex(userId, b.email);
      invalidateCache(SH_USERS);
      return { success:true, userId:userId, role:role, message:'Account created' };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success:false, message:'register: '+err }; }
}

// ============================================================
//  2. SIGN IN
// ============================================================

function signInUser(b) {
  try {
    var user = getUserByEmail(b.email);
    if (!user) return { success:false, message:'No account found for this email' };
    if (user.password_hash !== hashPw(b.password||''))
      return { success:false, message:'Invalid password' };
    return {
      success: true,
      userId:  user.user_id,
      name:    user.full_name,
      role:    user.role_id || 'student'
    };
  } catch(err) { return { success:false, message:'signIn: '+err }; }
}

// ============================================================
//  3. MARK ENTRY
// ============================================================

function markEntry(b) {
  try {
    if (!b.userId) return { success:false, message:'userId required' };

    var tz  = Session.getScriptTimeZone();
    var now = new Date();

    // A. Get user via index
    var user = getUserById(b.userId);
    if (!user) return { success:false, message:'User not found' };

    // B. Geofence check (200m) — only when GPS provided
    var lat = b.lat ? parseFloat(b.lat) : null;
    var lng = b.lng ? parseFloat(b.lng) : null;
    var dist = '';
    if (lat !== null && !isNaN(lat)) {
      dist = haversine(lat, lng, LAB_LAT, LAB_LNG);
      if (dist > GEOFENCE_M) {
        return {
          success:  false,
          code:     'TOO_FAR',
          distance: dist,
          message:  'You are '+dist+'m from campus. Must be within '+GEOFENCE_M+'m.'
        };
      }
    }

    // C. Duplicate check — read only user_id + date columns (narrow read)
    var attSheet  = getSheet(SH_ATTENDANCE);
    var lastRow   = attSheet.getLastRow();
    var dateStr   = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    if (lastRow >= 2) {
      var uids  = attSheet.getRange(2, AC.UID,  lastRow-1, 1).getValues();
      var dates = attSheet.getRange(2, AC.DATE, lastRow-1, 1).getValues();
      for (var j=0; j<uids.length; j++) {
        if (String(uids[j][0]).trim() === String(b.userId).trim() &&
            normDate(dates[j][0], tz) === dateStr) {
          return { success:false, message:'Attendance already marked for today.' };
        }
      }
    }

    // D. Write — LockService prevents concurrent corruption
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      var timeStr  = Utilities.formatDate(now, tz, 'HH:mm:ss');
      var latStr   = lat !== null ? String(lat) : '';
      var lngStr   = lng !== null ? String(lng) : '';
      var addr     = b.address || '';
      var distStr  = dist !== '' ? String(dist) : '';

      attSheet.appendRow([
        generateId('a'),          // attendance_id
        b.userId,                 // user_id
        b.locationId || 'DEFAULT',// location_id
        dateStr,                  // date — always string, never Date object
        timeStr,                  // entry_time
        latStr,                   // entry_lat
        lngStr,                   // entry_lng
        addr,                     // entry_address
        distStr,                  // entry_distance_m
        b.method || 'biometric',  // login_method
        '',                       // exit_time
        '','','','',              // exit_lat, exit_lng, exit_address, exit_distance_m
        '',                       // duration_mins
        now.toISOString()         // entry_timestamp (for duration calc)
      ]);
      SpreadsheetApp.flush();
      invalidateCache(SH_ATTENDANCE);

      return {
        success:        true,
        message:        '\u2713 Attendance marked at '+timeStr,
        name:           user.full_name,
        date:           dateStr,
        time:           timeStr,
        location:       addr || (latStr ? latStr+', '+lngStr : 'not captured'),
        gps:            latStr && lngStr ? latStr+', '+lngStr : '',
        distanceMeters: distStr
      };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success:false, message:'markEntry: '+err }; }
}

// ============================================================
//  4. MARK EXIT
// ============================================================

function markExit(b) {
  try {
    if (!b.userId) return { success:false, message:'userId required' };

    var tz      = Session.getScriptTimeZone();
    var now     = new Date();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    var sheet   = getSheet(SH_ATTENDANCE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success:false, message:'No attendance records found' };

    // Narrow read — only cols we need for matching + exit writing
    var data = sheet.getRange(1, 1, lastRow, 17).getValues();
    var headers = data[0];

    for (var i=1; i<data.length; i++) {
      var row = data[i];
      if (String(row[AC.UID-1]  ).trim() !== String(b.userId).trim()) continue;
      if (normDate(row[AC.DATE-1], tz)   !== dateStr) continue;

      // Found — check already exited
      if (String(row[AC.XTIME-1]||'').trim()) {
        return { success:false, message:'Exit already recorded at '+row[AC.XTIME-1] };
      }

      // Compute duration in minutes (integer)
      var durationMins = '';
      var tsRaw = row[AC.TIMESTAMP-1];
      if (tsRaw) {
        var entryDt = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
        if (!isNaN(entryDt.getTime())) {
          durationMins = Math.max(0, Math.round((now-entryDt)/60000));
        }
      }

      var xlat  = b.lat     ? String(parseFloat(b.lat)) : '';
      var xlng  = b.lng     ? String(parseFloat(b.lng)) : '';
      var xaddr = b.address || '';
      var xdist = (xlat&&xlng) ? haversine(parseFloat(xlat),parseFloat(xlng),LAB_LAT,LAB_LNG) : '';

      // Batch write — one setValues call instead of 7 separate setValue calls
      var lock = LockService.getScriptLock();
      lock.waitLock(5000);
      try {
        sheet.getRange(i+1, AC.XTIME, 1, 7).setValues([[
          timeStr,      // exit_time
          xlat,         // exit_lat
          xlng,         // exit_lng
          xaddr,        // exit_address
          xdist,        // exit_distance_m
          durationMins, // duration_mins (integer)
          ''            // padding (col 17 = entry_timestamp, leave untouched)
        ]]);
        SpreadsheetApp.flush();
        invalidateCache(SH_ATTENDANCE);
      } finally { lock.releaseLock(); }

      var hrs  = durationMins !== '' ? Math.floor(durationMins/60) : 0;
      var mins = durationMins !== '' ? durationMins%60 : 0;
      var durLabel = durationMins !== '' ? (hrs>0 ? hrs+'h '+mins+'m' : mins+'m') : '';

      return {
        success:        true,
        message:        '\u2713 Exit recorded at '+timeStr+(durLabel?' \u00b7 Duration: '+durLabel:''),
        exitTime:       timeStr,
        durationMins:   durationMins,
        duration:       durLabel,
        location:       xaddr || (xlat?xlat+', '+xlng:'not captured'),
        distanceMeters: xdist
      };
    }

    // Row not found — return dates we do have for debug
    var foundDates = [];
    for (var d2=1; d2<data.length; d2++) {
      if (String(data[d2][AC.UID-1]||'').trim()===String(b.userId).trim())
        foundDates.push(normDate(data[d2][AC.DATE-1],tz));
    }
    return {
      success: false,
      message: 'No attendance entry found for today ('+dateStr+').'+
        (foundDates.length ? ' Dates on record: '+foundDates.join(', ') : ' No records for this account.'),
      debug: 'userId='+b.userId+' tz='+tz
    };
  } catch(err) { return { success:false, message:'markExit: '+err }; }
}

// ============================================================
//  5. MY ATTENDANCE (student history)
// ============================================================

function getMyAttendance(b) {
  try {
    if (!b.userId) return { success:false, message:'userId required' };
    var tz = Session.getScriptTimeZone();
    var sheet = getSheet(SH_ATTENDANCE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success:true, records:[] };

    var data = sheet.getRange(2, 1, lastRow-1, 17).getValues();
    var records = [];
    data.forEach(function(row) {
      if (String(row[AC.UID-1]).trim() !== String(b.userId).trim()) return;
      records.push({
        date:          normDate(row[AC.DATE-1], tz),
        entryTime:     row[AC.ETIME-1]   || '',
        exitTime:      row[AC.XTIME-1]   || '',
        durationMins:  row[AC.DURATION-1] || '',
        entryAddress:  row[AC.EADDR-1]   || '',
        exitAddress:   row[AC.XADDR-1]   || '',
        method:        row[AC.METHOD-1]  || '',
        entryDistance: row[AC.EDIST-1]   || '',
        exitDistance:  row[AC.XDIST-1]   || ''
      });
    });
    records.reverse(); // newest first
    return { success:true, records:records };
  } catch(err) { return { success:false, message:'getMyAttendance: '+err }; }
}

// ============================================================
//  6. SESSIONS (teacher)
// ============================================================

function createSession(b) {
  try {
    if (b.role!=='teacher'&&b.role!=='admin')
      return { success:false, message:'Only teachers can create sessions' };
    if (!b.subject||!b.windowMinutes)
      return { success:false, message:'Subject and window required' };

    var tz     = Session.getScriptTimeZone();
    var now    = new Date();
    var sessId = generateId('s');
    var date   = Utilities.formatDate(now,tz,'yyyy-MM-dd');
    var start  = Utilities.formatDate(now,tz,'HH:mm:ss');
    var end    = Utilities.formatDate(new Date(now.getTime()+b.windowMinutes*60000),tz,'HH:mm:ss');

    var sheet  = getSheet(SH_SESSIONS);
    // Close existing open sessions by same teacher
    var data   = sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][SC.TID-1])===String(b.userId)&&data[i][SC.STATUS-1]==='open') {
        sheet.getRange(i+1,SC.STATUS).setValue('closed');
      }
    }
    sheet.appendRow([sessId,b.userId,b.teacherName||'',b.subject,date,start,end,'open',b.windowMinutes]);

    // Cache the active session for fast student polling
    CacheService.getScriptCache().put(
      'active_session', JSON.stringify({ session_id:sessId, teacher_id:b.userId,
        subject:b.subject, date:date, start_time:start, end_time:end, status:'open',
        window_minutes:b.windowMinutes }),
      TTL_SESSION
    );

    return { success:true, sessionId:sessId, subject:b.subject, startTime:start, endTime:end };
  } catch(err) { return { success:false, message:'createSession: '+err }; }
}

function closeSession(b) {
  try {
    var sheet = getSheet(SH_SESSIONS);
    var data  = sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][SC.ID-1])===String(b.sessionId)) {
        sheet.getRange(i+1,SC.STATUS).setValue('closed');
        CacheService.getScriptCache().remove('active_session');
        return { success:true };
      }
    }
    return { success:false, message:'Session not found' };
  } catch(err) { return { success:false, message:'closeSession: '+err }; }
}

function getActiveSession(b) {
  try {
    // Check CacheService first — avoids sheet read on every 30s student poll
    var cache = CacheService.getScriptCache();
    var cached = cache.get('active_session');
    if (cached) {
      var s = JSON.parse(cached);
      var now = new Date();
      var tz  = Session.getScriptTimeZone();
      var today = Utilities.formatDate(now,tz,'yyyy-MM-dd');
      if (s.status==='open' && s.date===today) {
        var endDt = new Date(s.date+'T'+s.end_time);
        if (now <= endDt) {
          return { success:true, active:true, session:s,
                   secondsLeft:Math.max(0,Math.round((endDt-now)/1000)) };
        }
      }
      cache.remove('active_session');
    }

    // Cache miss — read Sessions sheet
    var sheet = getSheet(SH_SESSIONS);
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
    var rows  = getRows(sheet);
    var now2  = new Date();

    for (var i=0;i<rows.length;i++) {
      var s=rows[i];
      if (s.status==='open' && normDate(s.date,tz)===today) {
        var st=new Date(s.date+'T'+s.start_time), en=new Date(s.date+'T'+s.end_time);
        if (now2>=st && now2<=en) {
          var secs=Math.max(0,Math.round((en-now2)/1000));
          cache.put('active_session',JSON.stringify(s),TTL_SESSION);
          return { success:true, active:true, session:s, secondsLeft:secs };
        }
      }
    }
    return { success:true, active:false };
  } catch(err) { return { success:false, message:'getActiveSession: '+err }; }
}

function getSessions(b) {
  try {
    var tz   = Session.getScriptTimeZone();
    var rows = getRows(getSheet(SH_SESSIONS));
    if (b.userId) rows = rows.filter(function(r){ return String(r.teacher_id)===String(b.userId); });

    // Count present students per session — narrow read (user_id + date cols only)
    var attSheet  = getSheet(SH_ATTENDANCE);
    var attLast   = attSheet.getLastRow();
    var attCounts = {};
    if (attLast >= 2) {
      var uids  = attSheet.getRange(2,AC.UID, attLast-1,1).getValues();
      var dates = attSheet.getRange(2,AC.DATE,attLast-1,1).getValues();
      uids.forEach(function(u,i){
        var key=normDate(dates[i][0],tz);
        attCounts[key] = (attCounts[key]||0)+1;
      });
    }

    var result = rows.map(function(s){
      return {
        sessionId:    s.session_id,
        subject:      s.subject,
        date:         normDate(s.date,tz),
        startTime:    s.start_time,
        endTime:      s.end_time,
        status:       s.status,
        presentCount: attCounts[normDate(s.date,tz)] || 0
      };
    }).reverse().slice(0,20);

    return { success:true, sessions:result };
  } catch(err) { return { success:false, message:'getSessions: '+err }; }
}

// ============================================================
//  7. DASHBOARD (cached)
// ============================================================

function getDashboard(b) {
  try {
    if (!b.sessionId) return { success:false, message:'sessionId required' };

    // Cache check
    var cacheKey = 'dashboard_'+b.sessionId;
    var cache    = CacheService.getScriptCache();
    var cached   = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    // All students (cached lookup)
    var userRows = getCachedRows(SH_USERS, TTL_LOOKUP);
    var students = userRows.filter(function(u){ return (u.role_id||'student')==='student'; });

    // Who marked — read only user_id column for this session's date
    // (We don't have session-to-date in this version — use all today's attendance)
    var tz      = Session.getScriptTimeZone();
    var attSheet = getSheet(SH_ATTENDANCE);
    var attLast  = attSheet.getLastRow();
    var presentMap = {};
    if (attLast >= 2) {
      var all = attSheet.getRange(2,1,attLast-1,17).getValues();
      all.forEach(function(row){
        presentMap[String(row[AC.UID-1]).trim()] = {
          time:    row[AC.ETIME-1]  || '',
          method:  row[AC.METHOD-1] || '',
          dist:    row[AC.EDIST-1]  || '',
          xtime:   row[AC.XTIME-1]  || '',
          durMins: row[AC.DURATION-1] || ''
        };
      });
    }

    var present=[], absent=[];
    students.forEach(function(s){
      var uid=String(s.user_id).trim();
      if (presentMap[uid]) {
        var pm=presentMap[uid];
        var durLabel='';
        if (pm.durMins!=='') {
          var h=Math.floor(pm.durMins/60),m=pm.durMins%60;
          durLabel=h>0?h+'h '+m+'m':m+'m';
        }
        present.push({
          userId:     uid,
          name:       s.full_name,
          email:      s.email,
          department: s.department_id,
          entryTime:  pm.time,
          exitTime:   pm.xtime,
          duration:   durLabel,
          method:     pm.method,
          distance:   pm.dist
        });
      } else {
        absent.push({
          userId:     uid,
          name:       s.full_name,
          email:      s.email,
          department: s.department_id
        });
      }
    });

    var result = {
      success:      true,
      total:        students.length,
      presentCount: present.length,
      absentCount:  absent.length,
      present:      present,
      absent:       absent
    };
    try { cache.put(cacheKey, JSON.stringify(result), TTL_DASHBOARD); } catch(e){}
    return result;
  } catch(err) { return { success:false, message:'getDashboard: '+err }; }
}

// ============================================================
//  8. STUDENTS ROSTER
// ============================================================

function getStudents(b) {
  try {
    var rows = getCachedRows(SH_USERS, TTL_LOOKUP)
      .filter(function(r){ return (r.role_id||'student')==='student'; })
      .map(function(r){
        return {
          userId:     r.user_id,
          name:       r.full_name,
          email:      r.email,
          department: r.department_id,
          hasBio:     !!r.biometric_credential_id,
          hasDevice:  !!r.device_id
        };
      });
    return { success:true, students:rows, total:rows.length };
  } catch(err) { return { success:false, message:'getStudents: '+err }; }
}

// ============================================================
//  9. EXPORT ATTENDANCE (CSV)
// ============================================================

function exportAttendance(b) {
  try {
    var tz      = Session.getScriptTimeZone();
    var sheet   = getSheet(SH_ATTENDANCE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success:true, csv:'', rowCount:0 };

    var data = sheet.getRange(2,1,lastRow-1,17).getValues();

    // Build user map once
    var userMap = buildMap(getCachedRows(SH_USERS, TTL_LOOKUP), 'user_id');

    var header = ['Name','Email','Department','Date','Entry Time','Exit Time',
                  'Duration (min)','Method','Entry Address','Exit Address',
                  'Entry Distance (m)','Exit Distance (m)'];
    var lines  = [header.join(',')];

    data.forEach(function(row) {
      var uid = String(row[AC.UID-1]).trim();
      // Filter by date or userId if provided
      if (b.date && normDate(row[AC.DATE-1],tz)!==b.date) return;
      if (b.userId && uid!==String(b.userId).trim()) return;

      var u = userMap[uid] || {};
      lines.push([
        '"'+(u.full_name||'')     +'"',
        '"'+(u.email||'')         +'"',
        '"'+(u.department_id||'') +'"',
        normDate(row[AC.DATE-1],tz),
        row[AC.ETIME-1]     || '',
        row[AC.XTIME-1]     || '',
        row[AC.DURATION-1]  || '',
        row[AC.METHOD-1]    || '',
        '"'+(row[AC.EADDR-1]||'')+'"',
        '"'+(row[AC.XADDR-1]||'')+'"',
        row[AC.EDIST-1]     || '',
        row[AC.XDIST-1]     || ''
      ].join(','));
    });

    return { success:true, csv:lines.join('\n'), rowCount:lines.length-1 };
  } catch(err) { return { success:false, message:'exportAttendance: '+err }; }
}

// ============================================================
//  10. BIOMETRIC & DEVICE
// ============================================================

function getBiometric(b) {
  try {
    var user = getUserByEmail(b.email);
    if (!user) return { success:false, message:'No account found for this email. Please register first.' };
    if (!user.biometric_credential_id)
      return { success:false, message:'No biometric registered. Please register your fingerprint first.' };
    return { success:true, credentialId:user.biometric_credential_id,
             userId:user.user_id, name:user.full_name };
  } catch(err) { return { success:false, message:'getBiometric: '+err }; }
}

function saveBiometric(b) {
  try {
    var sheet   = getSheet(SH_USERS);
    var lastRow = sheet.getLastRow();
    var uids    = sheet.getRange(2,UC.ID,lastRow-1,1).getValues();
    for (var i=0;i<uids.length;i++) {
      if (String(uids[i][0])===String(b.userId)) {
        sheet.getRange(i+2,UC.BIO).setValue(b.credentialId);
        invalidateCache(SH_USERS);
        return { success:true };
      }
    }
    return { success:false, message:'User not found' };
  } catch(err) { return { success:false, message:'saveBiometric: '+err }; }
}

function registerDevice(b) {
  try {
    if (!b.userId||!b.deviceId) return { success:false, message:'userId and deviceId required' };
    var user = getUserById(b.userId);
    if (!user) return { success:false, message:'User not found' };
    if (user.device_id && user.device_id!==b.deviceId)
      return { success:false, alreadyBound:true, message:'Account already bound to another device.' };

    // Write device_id
    var sheet   = getSheet(SH_USERS);
    var lastRow = sheet.getLastRow();
    var uids    = sheet.getRange(2,UC.ID,lastRow-1,1).getValues();
    for (var i=0;i<uids.length;i++) {
      if (String(uids[i][0])===String(b.userId)) {
        sheet.getRange(i+2,UC.DEVICE).setValue(b.deviceId);
        invalidateCache(SH_USERS);
        return { success:true, alreadyBound:false, message:'Device registered' };
      }
    }
    return { success:false, message:'User not found' };
  } catch(err) { return { success:false, message:'registerDevice: '+err }; }
}

function checkDevice(b) {
  try {
    if (!b.userId||!b.deviceId) return { success:false, message:'userId and deviceId required' };
    var user = getUserById(b.userId);
    if (!user) return { success:false, message:'User not found' };
    var stored = String(user.device_id||'').trim();
    if (!stored) return { success:true, status:'unbound' };
    if (stored===String(b.deviceId).trim()) return { success:true, status:'match' };
    return { success:false, status:'mismatch', message:'Account registered to a different device.' };
  } catch(err) { return { success:false, message:'checkDevice: '+err }; }
}

// ============================================================
//  11. SETUP — run once from editor to initialise all sheets
// ============================================================

function setupSheets() {
  try {
    // Create all sheets with headers
    [SH_USERS, SH_ATTENDANCE, SH_SESSIONS,
     SH_DEPARTMENTS, SH_ROLES, SH_LOCATIONS, SH_USER_INDEX
    ].forEach(function(name){ getSheet(name); });

    // Seed Roles if empty
    var rolesSheet = getSheet(SH_ROLES);
    if (rolesSheet.getLastRow() < 2) {
      rolesSheet.appendRow(['student', 'Student']);
      rolesSheet.appendRow(['teacher', 'Teacher']);
      rolesSheet.appendRow(['admin',   'Admin']);
    }

    // Seed Departments (SIT departments)
    var deptSheet = getSheet(SH_DEPARTMENTS);
    if (deptSheet.getLastRow() < 2) {
      var depts = [
        ['CSE',  'Computer Science & Engineering',    '',  ''],
        ['ISE',  'Information Science & Engineering', '',  ''],
        ['ECE',  'Electronics & Communication',       '',  ''],
        ['EEE',  'Electrical & Electronics',          '',  ''],
        ['MECH', 'Mechanical Engineering',            '',  ''],
        ['CIVIL','Civil Engineering',                 '',  ''],
        ['MBA',  'MBA',                               '',  ''],
        ['MCA',  'MCA',                               '',  '']
      ];
      depts.forEach(function(d){ deptSheet.appendRow(d); });
    }

    // Seed default location
    var locSheet = getSheet(SH_LOCATIONS);
    if (locSheet.getLastRow() < 2) {
      locSheet.appendRow(['DEFAULT', 'SIT Campus Lab', 13.32603, 77.12621, 200]);
    }

    return { success:true, message:'All sheets initialised. Sheets created: '+
      [SH_USERS,SH_ATTENDANCE,SH_SESSIONS,SH_DEPARTMENTS,
       SH_ROLES,SH_LOCATIONS,SH_USER_INDEX].join(', ') };
  } catch(err) { return { success:false, message:'setupSheets: '+err }; }
}

// ============================================================
//  12. DEBUG
// ============================================================

function debugInfo(b) {
  try {
    var sheets = ss().getSheets().map(function(s){
      return { name:s.getName(), rows:s.getLastRow()-1 };
    });
    var idxRows = getSheet(SH_USER_INDEX).getLastRow()-1;
    return {
      success:    true,
      spreadsheet: ss().getName(),
      sheets:     sheets,
      userIndex:  idxRows+' users indexed',
      geofence:   GEOFENCE_M+'m around ('+LAB_LAT+', '+LAB_LNG+')'
    };
  } catch(err) { return { success:false, message:'debug: '+err }; }
}