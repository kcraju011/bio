// ============================================================
//  BioAttend – Google Apps Script Backend  v6
//  College  : Siddaganga Institute of Technology, Tumkur
//  Deploy   : Web App → Execute as Me → Access: Anyone
//
//  SCHEMA — 8 tables exactly as specified:
//  1. Users
//  2. Departments
//  3. Roles
//  4. Attendance
//  5. LocationMonitor
//  6. AttendanceType
//  7. AttendanceLocations
//  8. UserAttendanceLocationMap
//  9. CategoryLocationMap
// ============================================================

// ── Sheet names ───────────────────────────────────────────────
var SH = {
  USERS        : 'Users',
  DEPARTMENTS  : 'Departments',
  ROLES        : 'Roles',
  ATTENDANCE   : 'Attendance',
  LOC_MONITOR  : 'LocationMonitor',
  ATT_TYPE     : 'AttendanceType',
  ATT_LOCATIONS: 'AttendanceLocations',
  USER_LOC_MAP : 'UserAttendanceLocationMap',
  CATEGORY_LOC_MAP: 'CategoryLocationMap',
  WORK_LOC_REQUESTS: 'WorkLocationRequests',
  SESSIONS     : 'Sessions',
  USER_INDEX   : 'UserIndex'
};

// ── Exact column headers per table ───────────────────────────
var HEADERS = {
  Users: [
    'user_id', 'institute_id', 'department_id', 'role_id',
    'full_name', 'dob', 'mobile', 'email',
    'password_hash', 'biometric_code', 'device_identification',
    'subcategory_id', 'tenant_id'
  ],
  Departments: [
    'department_id', 'name', 'in_charge', 'email', 'tenant_id'
  ],
  Roles: [
    'role_id', 'name'
  ],
  Attendance: [
    'attendance_id', 'user_id', 'full_name', 'type_attendance',
    'entry_time', 'exit_time', 'attendance_date', 'login_method',
    'latitude', 'longitude', 'attendance_location_id',
    'address', 'distance_from_centre', 'tenant_id',
    'category_id', 'subcategory_id'
  ],
  LocationMonitor: [
    'location_monitor_id', 'user_id', 'latitude', 'longitude',
    'distance_from_centre', 'timestamp', 'tenant_id'
  ],
  AttendanceType: [
    'attendance_type_id', 'type'
  ],
  AttendanceLocations: [
    'attendance_location_id', 'name', 'address', 'latitude', 'longitude',
    'geofence_radius', 'tenant_id'
  ],
  UserAttendanceLocationMap: [
    'user_attendance_location_map_id', 'user_id',
    'attendance_location_id', 'allowed_distance', 'tenant_id'
  ],
  CategoryLocationMap: [
    'category_location_map_id', 'category_id', 'subcategory_id',
    'attendance_location_id', 'allowed_distance', 'tenant_id', 'status'
  ],
  WorkLocationRequests: [
    'request_id', 'user_id', 'full_name', 'request_date',
    'location_name', 'address', 'latitude', 'longitude',
    'allowed_distance', 'reason', 'status',
    'approved_by', 'approved_at', 'rejected_by', 'rejected_at',
    'tenant_id', 'created_at', 'updated_at'
  ],
  Sessions: [
    'session_id', 'teacher_id', 'teacher_name', 'subject',
    'date', 'start_time', 'end_time', 'status', 'window_minutes', 'tenant_id'
  ],
  UserIndex: [
    'user_id', 'email', 'row_number', 'tenant_id'
  ]
};

// ── GPS defaults ──────────────────────────────────────────────
var DEFAULT_LAT = 13.32603;
var DEFAULT_LNG = 77.12621;
var DEFAULT_RADIUS = 200;
var DEFAULT_GUID = '1';
var REQUEST_CONTEXT = { guid: DEFAULT_GUID };
var AUTO_SESSION_WINDOW_MINUTES = 10;
var AUTO_SESSION_TIMES = ['09:00', '10:30', '14:00', '15:30'];
var AUTO_SESSION_SUBJECT = 'Automatic Attendance Session';

// ── Cache TTLs ────────────────────────────────────────────────
var TTL_LOOKUP  = 600;
var TTL_DASH    = 60;
var TTL_SESSION = 30;

// ============================================================
//  TRANSPORT
// ============================================================

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOut(obj, callback) {
  return ContentService.createTextOutput(String(callback || '') + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function iframeOut(obj, callback) {
  var payload = JSON.stringify(obj).replace(/</g, '\\u003c');
  var cb = String(callback || '').replace(/'/g, '\\u0027');
  var html = '<!doctype html><html><body><script>' +
    'parent.postMessage({__ba_iframe_cb:\'' + cb + '\',__ba_iframe_data:' + payload + '},\"*\");' +
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html);
}

function respondOut(obj, callback) {
  return callback ? jsonpOut(obj, callback) : jsonOut(obj);
}

function doGet(e) {
  try {
    var p = e && e.parameter ? e.parameter : {};
    var callback = p.callback || '';
    var transport = String(p.transport || '').toLowerCase();
    setRequestGuid(p.guid);
    var result = !p.data
      ? { status: 'BioAttend v6 running', time: new Date().toString() }
      : route(JSON.parse(decodeURIComponent(p.data)));
    if (transport === 'iframe') return iframeOut(result, callback);
    return respondOut(result, callback);
  } catch(err) { return jsonOut({ success: false, message: 'doGet: ' + err }); }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var callback = (e.parameter && e.parameter.callback) || body.callback || '';
    var transport = String((e.parameter && e.parameter.transport) || body.transport || '').toLowerCase();
    setRequestGuid(body.guid || (e.parameter && e.parameter.guid));
    var result = route(body);
    if (transport === 'iframe') return iframeOut(result, callback);
    return respondOut(result, callback);
  }
  catch(err) { return jsonOut({ success: false, message: 'doPost: ' + err }); }
}

// ============================================================
//  ROUTER
// ============================================================

function route(b) {
  setRequestGuid(b && b.guid);
  switch (b.action) {
    // ── Auth & Registration
    case 'registerUser':          return registerUser(b);
    case 'register':              return registerUser(b);
    case 'loginUser':             return signInUser(b);
    case 'signIn':                return signInUser(b);
    case 'saveBiometric':         return saveBiometric(b);
    case 'getBiometric':          return getBiometric(b);
    case 'registerDevice':        return registerDevice(b);
    case 'checkDevice':           return checkDevice(b);

    // ── Attendance
    case 'markAttendance':        return markAttendance(b);
    case 'markEntry':             return markEntry(b);
    case 'markExit':              return markExit(b);
    case 'exitAttendance':        return markExit(b);
    case 'trackStudentLocation':  return trackStudentLocation(b);
    case 'getMyAttendance':       return getMyAttendance(b);
    case 'exportAttendance':      return exportAttendance(b);

    // ── Sessions (teacher)
    case 'createSession':         return createSession(b);
    case 'closeSession':          return closeSession(b);
    case 'getActiveSession':      return getActiveSession(b);
    case 'getSessions':           return getSessions(b);
    case 'getTeacherNotifications': return getTeacherNotifications(b);

    // ── Dashboard
    case 'getDashboard':          return getDashboard(b);
    case 'getUsers':              return getUsers(b);
    case 'getStudents':           return getStudents(b);

    // ── Lookup tables (read)
    case 'getDepartments':        return getDepartments();
    case 'getRoles':              return getRoles();
    case 'getAttendanceTypes':    return getLookup(SH.ATT_TYPE);
    case 'getLocations':          return getLocations();
    case 'getAttendanceLocations': return getLocations();
    case 'getUserLocMap':         return getUserLocMap(b);
    case 'getCategoryLocationMap': return getCategoryLocationMap(b);
    case 'getWorkLocationRequests': return getWorkLocationRequests(b);

    // ── Admin write actions (called from Admin tab UI)
    case 'addDepartment':         return addDepartment(b);
    case 'addAttendanceLocation': return addAttendanceLocation(b);
    case 'addUserLocMap':         return addUserLocMap(b);
    case 'addCategoryLocationMap': return addCategoryLocationMap(b);
    case 'createWorkLocationRequest': return createWorkLocationRequest(b);
    case 'approveWorkLocationRequest': return approveWorkLocationRequest(b);
    case 'rejectWorkLocationRequest': return rejectWorkLocationRequest(b);

    // ── Admin (run from editor only)
    case 'setupSheets':           return setupSheets();
    case 'debug':                 return debugInfo();
    default: return { success: false, message: 'Unknown action: ' + b.action };
  }
}

// ============================================================
//  SHEET UTILITIES
// ============================================================

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Get or create a sheet — uses exact headers from HEADERS map
function getSheet(name) {
  var sheet = ss().getSheetByName(name);
  if (!sheet) {
    sheet = ss().insertSheet(name);
    var h = HEADERS[name];
    if (h) {
      sheet.appendRow(h);
      sheet.getRange(1, 1, 1, h.length)
        .setFontWeight('bold')
        .setBackground('#0f172a')
        .setFontColor('#f1f5f9');
      sheet.setFrozenRows(1);
    }
  }
  ensureSheetHeaders(sheet, name);
  return sheet;
}

function ensureSheetHeaders(sheet, name) {
  var desired = HEADERS[name];
  if (!desired || !desired.length) return;
  var lastCol = Math.max(sheet.getLastColumn(), desired.length);
  var current = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || '').trim(); })
    : [];
  var matches = current.length >= desired.length;
  if (matches) {
    for (var i = 0; i < desired.length; i++) {
      if (current[i] !== desired[i]) { matches = false; break; }
    }
  }
  if (!matches) {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(desired);
    } else {
      sheet.getRange(1, 1, 1, desired.length).setValues([desired]);
    }
    sheet.getRange(1, 1, 1, desired.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#f1f5f9');
    sheet.setFrozenRows(1);
  }
}

// Read rows as array of objects keyed by header name
// Uses getLastRow() — never reads empty rows
function getRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var data = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  var hdrs = data[0];
  return data.slice(1).map(function(row) {
    var o = {};
    hdrs.forEach(function(h, i) { o[h] = row[i]; });
    return o;
  }).filter(function(row) {
    var tenant = String(row.tenant_id || '').trim();
    return !tenant || tenant === currentGuid();
  });
}

// CacheService wrapper — avoids sheet reads for lookup tables
function getCached(sheetName, ttl) {
  var cache = CacheService.getScriptCache();
  var hit   = cache.get(tenantScopedKey('rows_' + sheetName));
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  var rows  = getRows(getSheet(sheetName));
  try { cache.put(tenantScopedKey('rows_' + sheetName), JSON.stringify(rows), ttl || TTL_LOOKUP); } catch(e) {}
  return rows;
}

function invalidate(sheetName) {
  try { CacheService.getScriptCache().remove(tenantScopedKey('rows_' + sheetName)); } catch(e) {}
}

// Build O(1) lookup map from array of objects
function buildMap(rows, key) {
  var m = {};
  rows.forEach(function(r) { m[r[key]] = r; });
  return m;
}

function normalizeGuid(value) {
  var raw = String(value || '').trim();
  return raw || DEFAULT_GUID;
}

function setRequestGuid(guid) {
  REQUEST_CONTEXT.guid = normalizeGuid(guid);
  return REQUEST_CONTEXT.guid;
}

function currentGuid() {
  return normalizeGuid(REQUEST_CONTEXT.guid);
}

function getTenantGeofence(guid) {
  var tenantGuid = normalizeGuid(guid || currentGuid());
  return {
    latitude: DEFAULT_LAT,
    longitude: DEFAULT_LNG,
    radius: DEFAULT_RADIUS,
    label: tenantGuid
  };
}

function tenantScopedKey(key) {
  return currentGuid() + '__' + String(key || '');
}

function tenantIdOf(value) {
  return normalizeGuid(value || currentGuid());
}

function rowMatchesTenant(rowTenant, tenantId) {
  var rowValue = String(rowTenant || '').trim();
  if (!rowValue) return true;
  return rowValue === tenantIdOf(tenantId);
}

function toCleanText(value, maxLen) {
  var s = String(value || '').replace(/\s+/g, ' ').trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

function appendAuditLog(eventType, actorUserId, actorRole, targetUserId, details) {
  try {
    var sheet = getSheet('AuditLog');
    sheet.appendRow([
      genId('audit'),
      String(eventType || ''),
      String(actorUserId || ''),
      String(actorRole || ''),
      String(targetUserId || ''),
      JSON.stringify(details || {}),
      '',
      new Date().toISOString(),
      currentGuid()
    ]);
  } catch (e) {}
}

// ── UserIndex: simulated index for O(1) user lookup ──────────
// Maintains a 3-col sheet: user_id | email | row_number_in_Users
// Written on register — avoids full Users scan on every login.

function getUserByEmail(email) {
  var lower = String(email).toLowerCase();
  var idx   = getSheet(SH.USER_INDEX);
  var data  = idx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === lower) {
      if (data[i][3] && String(data[i][3]).trim() !== currentGuid()) continue;
      return fetchUserRow(data[i][2]);
    }
  }
  return null;
}

function getUserById(userId) {
  var idx  = getSheet(SH.USER_INDEX);
  var data = idx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      if (data[i][3] && String(data[i][3]).trim() !== currentGuid()) continue;
      return fetchUserRow(data[i][2]);
    }
  }
  return null;
}

function fetchUserRow(rowNum) {
  var sheet = getSheet(SH.USERS);
  var ncols = HEADERS.Users.length;
  var row   = sheet.getRange(rowNum, 1, 1, ncols).getValues()[0];
  var obj   = {};
  HEADERS.Users.forEach(function(h, i) { obj[h] = row[i]; });
  return obj;
}

function addToUserIndex(userId, email) {
  var rowNum = getSheet(SH.USERS).getLastRow();
  getSheet(SH.USER_INDEX).appendRow([userId, String(email).toLowerCase(), rowNum, currentGuid()]);
}

// ── Helpers ───────────────────────────────────────────────────

function genId(prefix) {
  return (prefix || 'id') + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
}

function genNumericId() {
  var ts = Utilities.formatDate(new Date(), tz(), 'yyyyMMddHHmmssSSS');
  var suffix = Math.floor(Math.random() * 100).toString().padStart(2, '0');
  return ts + suffix;
}

function hashPw(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function haversine(lat1, lng1, lat2, lng2) {
  var R = 6371000, d2r = Math.PI / 180;
  var a = Math.sin((lat2-lat1)*d2r/2) * Math.sin((lat2-lat1)*d2r/2) +
          Math.cos(lat1*d2r) * Math.cos(lat2*d2r) *
          Math.sin((lng2-lng1)*d2r/2) * Math.sin((lng2-lng1)*d2r/2);
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
}

function normDate(raw, tz) {
  if (!raw && raw !== 0) return '';
  if (raw instanceof Date) return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
  var s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (s.length > 10 && s.indexOf('T') >= 0) return s.slice(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return s.slice(0, 10);
}

function tz() { return Session.getScriptTimeZone(); }

function pad2(n) {
  return String(n).padStart(2, '0');
}

function normalizeClockTime(raw) {
  var m = String(raw || '').trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return '';
  var h = parseInt(m[1], 10);
  var min = parseInt(m[2], 10);
  if (h < 0 || h > 23 || min < 0 || min > 59) return '';
  return pad2(h) + ':' + pad2(min);
}

function buildAutoSession(dateStr, startHm) {
  var norm = normalizeClockTime(startHm);
  if (!norm) return null;
  var startDt = new Date(dateStr + 'T' + norm + ':00');
  var endDt = new Date(startDt.getTime() + AUTO_SESSION_WINDOW_MINUTES * 60000);
  return {
    session_id: 'auto_' + dateStr + '_' + norm.replace(':', ''),
    teacher_id: 'ALL',
    teacher_name: 'Automatic',
    subject: AUTO_SESSION_SUBJECT,
    date: dateStr,
    start_time: norm + ':00',
    end_time: Utilities.formatDate(endDt, tz(), 'HH:mm:ss'),
    status: 'open',
    window_minutes: AUTO_SESSION_WINDOW_MINUTES,
    is_auto: true
  };
}

function getAutoSessionsForDate(dateStr) {
  return AUTO_SESSION_TIMES.map(function(startHm) {
    return buildAutoSession(dateStr, startHm);
  }).filter(function(s) { return !!s; });
}

function getCurrentAutoSession(now) {
  var t = tz();
  var dateStr = Utilities.formatDate(now || new Date(), t, 'yyyy-MM-dd');
  var current = now || new Date();
  var sessions = getAutoSessionsForDate(dateStr);
  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    var st = new Date(s.date + 'T' + s.start_time);
    var en = new Date(s.date + 'T' + s.end_time);
    if (current >= st && current <= en) return s;
  }
  return null;
}

function getRoleDirectory() {
  var rows = getCached(SH.ROLES, TTL_LOOKUP);
  var byId = {};
  var byName = {};
  rows.forEach(function(r) {
    var id = String(r.role_id || '').trim();
    var name = String(r.name || '').trim().toLowerCase();
    if (id) byId[id] = name;
    if (name) byName[name] = id;
  });
  return { byId: byId, byName: byName };
}

function normalizeRoleValue(roleValue) {
  var raw = String(roleValue || '').trim().toLowerCase();
  if (!raw) return '';
  var roles = getRoleDirectory();
  return roles.byId[raw] || raw;
}

function resolveRoleIdInput(roleValue) {
  var raw = String(roleValue || '').trim();
  if (!raw) return '';
  var roles = getRoleDirectory();
  if (roles.byId[raw]) return raw;
  var roleId = findRoleIdByName(raw);
  if (roleId) return roleId;
  return raw;
}

function normalizeAttendanceType(typeValue) {
  var raw = String(typeValue || '').trim().toLowerCase();
  if (!raw) return '';
  if (raw === 'entry' || raw === 'checkin' || raw === 'check-in') return 'check-in';
  if (raw === 'exit' || raw === 'checkout' || raw === 'check-out') return 'check-out';
  return raw;
}

function isCheckInAttendance(typeValue) {
  return normalizeAttendanceType(typeValue) === 'check-in';
}

function isCheckOutAttendance(typeValue) {
  return normalizeAttendanceType(typeValue) === 'check-out';
}

function findRoleIdByName(roleName) {
  var raw = String(roleName || '').trim().toLowerCase();
  if (!raw) return '';
  return getRoleDirectory().byName[raw] || '';
}

function ensureExactRows(sheetName, rows) {
  var sheet = getSheet(sheetName);
  var hdrs = HEADERS[sheetName];
  sheet.clearContents();
  sheet.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  if (rows && rows.length) {
    sheet.getRange(2, 1, rows.length, hdrs.length).setValues(rows);
  }
  invalidate(sheetName);
}

function pushTeacherNotification(notification) {
  try {
    if (!notification || !notification.sessionId) return;
    var cache = CacheService.getScriptCache();
    var key = 'teacher_notif_' + notification.sessionId;
    var rows = [];
    var cached = cache.get(key);
    if (cached) {
      try { rows = JSON.parse(cached) || []; } catch(e) {}
    }
    rows.unshift(notification);
    rows = rows.slice(0, 25);
    cache.put(key, JSON.stringify(rows), 21600);
  } catch(e) {}
}

function getOpenSessionForToday() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('active_session');
  var now = new Date();
  var today = Utilities.formatDate(now, tz(), 'yyyy-MM-dd');

  if (cached) {
    try {
      var active = JSON.parse(cached);
      if (active && active.status === 'open' && active.date === today) return active;
    } catch(e) {}
  }

  var rows = getRows(getSheet(SH.SESSIONS));
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].status === 'open' && normDate(rows[i].date, tz()) === today) return rows[i];
  }
  return getCurrentAutoSession(now);
}

function parseNumber(value) {
  if (value === '' || value === null || typeof value === 'undefined') return null;
  var n = parseFloat(value);
  return isNaN(n) ? null : n;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function isValidMobile(mobile) {
  var raw = String(mobile || '').trim();
  if (!raw) return true;
  return /^[0-9+\-\s]{7,20}$/.test(raw);
}

function userCategoryOf(user) {
  return toCleanText((user && (user.department_id || user.category_id)) || '', 120);
}

function userSubcategoryOf(user) {
  return toCleanText((user && user.subcategory_id) || '', 120);
}

function locationDistanceMeters(lat1, lng1, lat2, lng2) {
  return haversine(lat1, lng1, lat2, lng2);
}

function findLocationById(locationId, tenantId) {
  var rows = getCached(SH.ATT_LOCATIONS, TTL_LOOKUP);
  var id = String(locationId || '').trim();
  for (var i = 0; i < rows.length; i++) {
    if (!rowMatchesTenant(rows[i].tenant_id, tenantId)) continue;
    if (String(rows[i].attendance_location_id) === id) return rows[i];
  }
  return null;
}

function findNearbyAttendanceLocations(lat, lng, tenantId, excludeLocationId) {
  var rows = getCached(SH.ATT_LOCATIONS, TTL_LOOKUP);
  var matches = [];
  rows.forEach(function(loc) {
    if (!rowMatchesTenant(loc.tenant_id, tenantId)) return;
    if (excludeLocationId && String(loc.attendance_location_id) === String(excludeLocationId)) return;
    var dLat = parseNumber(loc.latitude);
    var dLng = parseNumber(loc.longitude);
    if (dLat === null || dLng === null) return;
    var dist = locationDistanceMeters(lat, lng, dLat, dLng);
    if (dist <= 20) {
      matches.push({
        attendance_location_id: loc.attendance_location_id,
        name: loc.name,
        address: loc.address || '',
        latitude: dLat,
        longitude: dLng,
        geofence_radius: loc.geofence_radius || '',
        distance: dist
      });
    }
  });
  matches.sort(function(a, b) { return a.distance - b.distance; });
  return matches;
}

function findCategoryLocationMap(categoryId, subcategoryId, tenantId) {
  var rows = getCached(SH.CATEGORY_LOC_MAP, TTL_LOOKUP);
  var cat = String(categoryId || '').trim().toLowerCase();
  var sub = String(subcategoryId || '').trim().toLowerCase();
  var fallback = null;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!rowMatchesTenant(row.tenant_id, tenantId)) continue;
    if (String(row.status || 'active').toLowerCase() !== 'active') continue;
    var rowCat = String(row.category_id || '').trim().toLowerCase();
    var rowSub = String(row.subcategory_id || '').trim().toLowerCase();
    if (rowCat === cat && rowSub === sub) return row;
    if (rowCat === cat && !rowSub && !fallback) fallback = row;
  }
  return fallback;
}

function resolveMappedAttendanceLocation(user, tenantId) {
  var cat = userCategoryOf(user);
  var sub = userSubcategoryOf(user);
  var mapping = findCategoryLocationMap(cat, sub, tenantId);
  if (!mapping) return null;
  var location = findLocationById(mapping.attendance_location_id, tenantId);
  if (!location) return null;
  var allowedDistance = parseInt(mapping.allowed_distance || location.geofence_radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
  return {
    categoryId: cat,
    subcategoryId: sub,
    attendanceLocationId: location.attendance_location_id,
    allowedDistance: allowedDistance,
    anchorLat: parseNumber(location.latitude) || DEFAULT_LAT,
    anchorLng: parseNumber(location.longitude) || DEFAULT_LNG,
    location: location,
    mapping: mapping
  };
}

function getWorkLocationRequestForDate(userId, requestDate, tenantId, status) {
  var rows = getRows(getSheet(SH.WORK_LOC_REQUESTS));
  var uid = String(userId || '').trim();
  var day = String(requestDate || '').trim();
  var targetStatus = String(status || '').trim().toLowerCase();
  var tenant = tenantIdOf(tenantId);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (String(row.user_id || '').trim() !== uid) continue;
    if (normDate(row.request_date, tz()) !== day) continue;
    if (!rowMatchesTenant(row.tenant_id, tenant)) continue;
    if (targetStatus && String(row.status || '').toLowerCase() !== targetStatus) continue;
    return row;
  }
  return null;
}

function resolveApprovedWorkLocation(userId, requestDate, tenantId) {
  var day = normDate(requestDate || new Date(), tz());
  var request = getWorkLocationRequestForDate(userId, day, tenantId, 'approved');
  if (!request) return null;
  var lat = parseNumber(request.latitude);
  var lng = parseNumber(request.longitude);
  var radius = parseInt(request.allowed_distance || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
  return {
    categoryId: userCategoryOf(getUserById(userId)),
    subcategoryId: userSubcategoryOf(getUserById(userId)),
    attendanceLocationId: String(request.request_id || ''),
    allowedDistance: radius,
    anchorLat: lat === null ? DEFAULT_LAT : lat,
    anchorLng: lng === null ? DEFAULT_LNG : lng,
    location: {
      attendance_location_id: String(request.request_id || ''),
      name: String(request.location_name || 'Temporary Work Location'),
      address: String(request.address || ''),
      latitude: lat === null ? DEFAULT_LAT : lat,
      longitude: lng === null ? DEFAULT_LNG : lng,
      geofence_radius: radius
    },
    mapping: request,
    source: 'work_request'
  };
}

function resolveEffectiveAttendanceLocation(user, tenantId, requestDate) {
  var day = normDate(requestDate || new Date(), tz());
  var approved = user ? resolveApprovedWorkLocation(user.user_id, day, tenantId) : null;
  if (approved) return approved;
  var mapped = user ? resolveMappedAttendanceLocation(user, tenantId) : null;
  if (mapped) return mapped;
  var geofence = getTenantGeofence(tenantId);
  return {
    categoryId: userCategoryOf(user),
    subcategoryId: userSubcategoryOf(user),
    attendanceLocationId: '1',
    allowedDistance: geofence.radius,
    anchorLat: geofence.latitude,
    anchorLng: geofence.longitude,
    location: null,
    mapping: null
  };
}

function resolveUserTrackingLocation(userId, requestDate) {
  var user = getUserById(userId);
  var resolved = user ? resolveEffectiveAttendanceLocation(user, currentGuid(), requestDate || new Date()) : null;
  if (resolved) return resolved;
  var geofence = getTenantGeofence(currentGuid());
  return {
    categoryId: userCategoryOf(user),
    subcategoryId: userSubcategoryOf(user),
    attendanceLocationId: '1',
    allowedDistance: geofence.radius,
    anchorLat: geofence.latitude,
    anchorLng: geofence.longitude,
    location: null,
    mapping: null
  };
}

// ============================================================
//  1. REGISTER USER
//  Writes to: Users, UserIndex
//  Attendance location is resolved from category/subcategory mappings
// ============================================================

function registerUser(b) {
  try {
    if (!b.name || !b.email || !b.password || !b.roleId || !b.departmentId || !b.subcategoryId)
      return { success: false, message: 'name, email, password and roleId are required' };

    // Fast duplicate check via UserIndex
    if (getUserByEmail(b.email))
      return { success: false, message: 'Email already registered' };

    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var userId = genNumericId();
      var instituteId = String(b.instituteId || b.organization || b.organizationName || 'Siddaganga Institute of Technology').trim();
      var departmentId = String(b.departmentId || b.department || '').trim();
      var subcategoryId = String(b.subcategoryId || b.subcategory || '').trim();
      var roleId = resolveRoleIdInput(b.roleId || b.role || b.roleKey);
      if (!roleId) return { success: false, message: 'roleId is required' };

      // ── Write to Users table (exact columns) ──
      getSheet(SH.USERS).appendRow([
        userId,                      // user_id
        instituteId,                 // institute_id
        departmentId,                // department_id
        roleId,                      // role_id
        b.name,                      // full_name
        b.dob     || '',             // dob
        b.mobile  || '',             // mobile
        String(b.email).toLowerCase(), // email
        hashPw(b.password),          // password_hash
        '',                          // biometric_code  — filled later
        b.deviceId || '',
        subcategoryId,
        currentGuid()
      ]);

      // ── Update UserIndex ──
      addToUserIndex(userId, b.email);

      // ── Auto-create UserAttendanceLocationMap entry ──
      // Maps the new user to the default/selected attendance location
      invalidate(SH.USERS);
      appendAuditLog('register_user', userId, roleId, userId, { categoryId: departmentId, subcategoryId: subcategoryId });
      return { success: true, userId: userId, message: 'Account created successfully' };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success: false, message: 'registerUser: ' + err }; }
}

// ============================================================
//  2. SIGN IN
// ============================================================

function signInUser(b) {
  try {
    var user = getUserByEmail(b.email);
    if (!user) return { success: false, message: 'No account found for this email' };
    if (user.password_hash !== hashPw(b.password || ''))
      return { success: false, message: 'Incorrect password' };
    var roleName = normalizeRoleValue(user.role_id);
    return {
      success: true,
      userId:  user.user_id,
      name:    user.full_name,
      roleId:  roleName || String(user.role_id || ''),
      roleKey: roleName || String(user.role_id || ''),
      roleDbId: String(user.role_id || ''),
      deptId:  user.department_id,
      subcategoryId: user.subcategory_id || '',
      guid: currentGuid()
    };
  } catch(err) { return { success: false, message: 'signIn: ' + err }; }
}

// ============================================================
//  3. MARK ENTRY
//  Writes to: Attendance, LocationMonitor
//  Reads from: Users, UserAttendanceLocationMap, AttendanceLocations
// ============================================================

function markEntry(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };

    var now    = new Date();
    var t      = tz();
    var user   = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var dateStrNew = Utilities.formatDate(now, t, 'yyyy-MM-dd');

    var lat = parseNumber(b.latitude);
    var lng = parseNumber(b.longitude);
    if (lat === null || lng === null) {
      return { success: false, message: 'GPS location is required to mark attendance' };
    }

    var resolved = resolveEffectiveAttendanceLocation(user, currentGuid(), dateStrNew);
    var locId = resolved.attendanceLocationId || '1';
    var allowedDist = parseInt(resolved.allowedDistance || b.allowedDistance || b.allowed_distance || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
    var anchorLat = parseNumber(resolved.anchorLat) || DEFAULT_LAT;
    var anchorLng = parseNumber(resolved.anchorLng) || DEFAULT_LNG;
    var dist = haversine(lat, lng, anchorLat, anchorLng);
    if (dist > allowedDist) {
      return {
        success: false,
        code: 'TOO_FAR',
        distance: dist,
        allowed: allowedDist,
        message: 'You are outside allowed location. Must be within ' + allowedDist + 'm of campus.'
      };
    }

    // ── A. Resolve allowed distance for this user ──
    var attSheetNew = getSheet(SH.ATTENDANCE);
    var existingRowsNew = attSheetNew.getLastRow() >= 2
      ? attSheetNew.getRange(2, 1, attSheetNew.getLastRow() - 1, attSheetNew.getLastColumn()).getValues()
      : [];
    for (var er = 0; er < existingRowsNew.length; er++) {
      if (String(existingRowsNew[er][1]).trim() === String(b.userId).trim() &&
          normDate(existingRowsNew[er][6], t) === dateStrNew &&
          isCheckInAttendance(existingRowsNew[er][3])) {
        return { success: false, message: 'Attendance already marked for today.' };
      }
    }

    var lockNew = LockService.getScriptLock();
    lockNew.waitLock(6000);
    try {
      var timeStrNew = Utilities.formatDate(now, t, 'HH:mm:ss');
      var attIdNew = genId('att');
      attSheetNew.appendRow([
        attIdNew, b.userId, user.full_name, 'check-in', timeStrNew, '', dateStrNew,
        b.loginMethod || 'biometric', lat, lng, locId, b.address || '', dist, currentGuid(),
        userCategoryOf(user), userSubcategoryOf(user)
      ]);
      getSheet(SH.LOC_MONITOR).appendRow([genId('lm'), b.userId, lat, lng, dist, now.toISOString(), currentGuid()]);
      SpreadsheetApp.flush();
      invalidate(SH.ATTENDANCE);
      appendAuditLog('mark_entry', b.userId, normalizeRoleValue(user.role_id), b.userId, {
        time: timeStrNew, date: dateStrNew, locationId: locId, distance: dist,
        categoryId: userCategoryOf(user), subcategoryId: userSubcategoryOf(user)
      });
      return {
        success: true, attendanceId: attIdNew, message: '\u2713 Attendance marked at ' + timeStrNew,
        name: user.full_name, date: dateStrNew, time: timeStrNew,
        location: b.address || (lat ? lat + ', ' + lng : 'not captured'),
        distanceFromCentre: dist
      };
    } finally { lockNew.releaseLock(); }

    var userLocRows = getRows(getSheet(SH.USER_LOC_MAP));
    var userLocMap  = null;
    for (var x = 0; x < userLocRows.length; x++) {
      if (String(userLocRows[x].user_id) === String(b.userId)) {
        userLocMap = userLocRows[x]; break;
      }
    }
    var requestedLocId = String(b.attendanceLocationId || b.attendance_location_id || '').trim();
    var allowedDist = parseInt(b.allowedDistance || b.allowed_distance || (userLocMap ? userLocMap.allowed_distance : DEFAULT_RADIUS), 10) || DEFAULT_RADIUS;
    var locId       = requestedLocId || (userLocMap ? userLocMap.attendance_location_id : '1');

    // ── B. Resolve anchor coordinates for this location ──
    var anchorLat = DEFAULT_LAT, anchorLng = DEFAULT_LNG;
    var locRows = getCached(SH.ATT_LOCATIONS, TTL_LOOKUP);
    for (var lx = 0; lx < locRows.length; lx++) {
      if (String(locRows[lx].attendance_location_id) === String(locId)) {
        anchorLat = parseFloat(locRows[lx].latitude  || DEFAULT_LAT);
        anchorLng = parseFloat(locRows[lx].longitude || DEFAULT_LNG);
        break;
      }
    }

    // ── C. Geofence check ──
    var lat  = b.latitude  ? parseFloat(b.latitude)  : null;
    var lng  = b.longitude ? parseFloat(b.longitude) : null;
    var dist = 0;
    if (lat !== null && !isNaN(lat)) {
      dist = haversine(lat, lng, anchorLat, anchorLng);
      if (dist > allowedDist) {
        return {
          success:  false,
          code:     'TOO_FAR',
          distance: dist,
          allowed:  allowedDist,
          message:  'You are outside allowed location. Must be within ' + allowedDist + 'm of campus.'
        };
      }
    }

    // ── D. Prevent duplicate entry today ──
    var attSheet = getSheet(SH.ATTENDANCE);
    var lastRow  = attSheet.getLastRow();
    var dateStr  = Utilities.formatDate(now, t, 'yyyy-MM-dd');
    if (lastRow >= 2) {
      // Narrow read: user_id (col2) + attendance_date (col7) + type (col4)
      var uids  = attSheet.getRange(2, 2, lastRow - 1, 1).getValues();
      var dates = attSheet.getRange(2, 7, lastRow - 1, 1).getValues();
      var types = attSheet.getRange(2, 4, lastRow - 1, 1).getValues();
      for (var j = 0; j < uids.length; j++) {
        if (String(uids[j][0]).trim() === String(b.userId).trim() &&
            normDate(dates[j][0], t) === dateStr &&
            isCheckInAttendance(types[j][0])) {
          return { success: false, message: 'Attendance already marked for today.' };
        }
      }
    }

    // ── E. Write to Attendance table ──
    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var timeStr = Utilities.formatDate(now, t, 'HH:mm:ss');
      var attId   = genId('att');

      // Attendance row — exact schema columns
      attSheet.appendRow([
        attId,                        // attendance_id
        b.userId,                     // user_id
        user.full_name,               // full_name
        'check-in',                   // type_attendance
        timeStr,                      // entry_time  (attendance_time col 1)
        '',                           // exit_time   (attendance_time col 2)
        dateStr,                      // attendance_date
        b.loginMethod || 'biometric', // login_method
        lat !== null ? lat : '',      // latitude
        lng !== null ? lng : '',      // longitude
        locId,                        // attendance_location_id
        b.address || '',              // address
        dist || ''                    // distance_from_centre
      ]);

      // ── F. Write to LocationMonitor table ──
      if (lat !== null) {
        getSheet(SH.LOC_MONITOR).appendRow([
          genId('lm'),        // location_monitor_id
          b.userId,           // user_id
          lat,                // latitude
          lng,                // longitude
          dist,               // distance_from_centre
          now.toISOString()   // timestamp
        ]);
      }

      SpreadsheetApp.flush();
      invalidate(SH.ATTENDANCE);

      var activeSession = getOpenSessionForToday();
      if (activeSession) {
        pushTeacherNotification({
          id: genId('ntf'),
          sessionId: String(activeSession.session_id || ''),
          teacherId: String(activeSession.teacher_id || ''),
          userId: String(b.userId || ''),
          studentName: String(user.full_name || ''),
          message: String(user.full_name || 'Student') + ' marked attendance at ' + timeStr,
          time: timeStr,
          date: dateStr,
          createdAt: now.toISOString()
        });
      }

      return {
        success:          true,
        attendanceId:     attId,
        message:          '\u2713 Attendance marked at ' + timeStr,
        name:             user.full_name,
        date:             dateStr,
        time:             timeStr,
        location:         b.address || (lat ? lat + ', ' + lng : 'not captured'),
        distanceFromCentre: dist
      };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success: false, message: 'markEntry: ' + err }; }
}

// ============================================================
//  4. MARK EXIT
//  Updates existing Attendance row: sets exit_time, type → 'exit'
//  Also writes a new LocationMonitor row for exit coordinates
// ============================================================

function markExit(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };

    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var nowNew = new Date();
    var tNew = tz();
    var dateStrNew = Utilities.formatDate(nowNew, tNew, 'yyyy-MM-dd');
    var timeStrNew = Utilities.formatDate(nowNew, tNew, 'HH:mm:ss');
    var latNew = parseNumber(b.latitude);
    var lngNew = parseNumber(b.longitude);
    var resolvedNew = resolveEffectiveAttendanceLocation(user, currentGuid(), dateStrNew);
    var xdistNew = (latNew !== null && lngNew !== null)
      ? haversine(latNew, lngNew, resolvedNew.anchorLat, resolvedNew.anchorLng)
      : '';
    var sheetNew = getSheet(SH.ATTENDANCE);
    var lastRowNew = sheetNew.getLastRow();
    if (lastRowNew < 2) return { success: false, message: 'No attendance records found' };
    var dataNew = sheetNew.getRange(2, 1, lastRowNew - 1, sheetNew.getLastColumn()).getValues();
    for (var iNew = 0; iNew < dataNew.length; iNew++) {
      var rowNew = dataNew[iNew];
      if (String(rowNew[1]).trim() !== String(b.userId).trim()) continue;
      if (normDate(rowNew[6], tNew) !== dateStrNew) continue;
      if (!isCheckInAttendance(rowNew[3])) continue;
      if (String(rowNew[5]).trim()) continue;
      var entryTimeStrNew = String(rowNew[4] || '');
      var durationMinsNew = '';
      if (entryTimeStrNew) {
        var entryDtNew = new Date(dateStrNew + 'T' + entryTimeStrNew);
        if (!isNaN(entryDtNew.getTime())) durationMinsNew = Math.max(0, Math.round((nowNew - entryDtNew) / 60000));
      }
      var lockNew = LockService.getScriptLock();
      lockNew.waitLock(6000);
      try {
        var sheetRowNew = iNew + 2;
        sheetNew.getRange(sheetRowNew, 4, 1, 1).setValue('check-out');
        sheetNew.getRange(sheetRowNew, 6, 1, 1).setValue(timeStrNew);
        if (latNew !== null && lngNew !== null) {
          sheetNew.getRange(sheetRowNew, 9, 1, 5).setValues([[latNew, lngNew, resolvedNew.attendanceLocationId || '', b.address || '', xdistNew]]);
          getSheet(SH.LOC_MONITOR).appendRow([genId('lm'), b.userId, latNew, lngNew, xdistNew, nowNew.toISOString(), currentGuid()]);
        }
        SpreadsheetApp.flush();
        invalidate(SH.ATTENDANCE);
        appendAuditLog('mark_exit', b.userId, normalizeRoleValue(user.role_id), b.userId, {
          time: timeStrNew,
          date: dateStrNew,
          distance: xdistNew,
          categoryId: userCategoryOf(user),
          subcategoryId: userSubcategoryOf(user)
        });
      } finally { lockNew.releaseLock(); }
      var hrsNew = durationMinsNew !== '' ? Math.floor(durationMinsNew / 60) : 0;
      var minsNew = durationMinsNew !== '' ? durationMinsNew % 60 : 0;
      var durLabelNew = durationMinsNew !== '' ? (hrsNew > 0 ? hrsNew + 'h ' + minsNew + 'm' : minsNew + 'm') : '';
      return {
        success: true,
        message: '\u2713 Exit recorded at ' + timeStrNew + (durLabelNew ? ' \u00b7 ' + durLabelNew : ''),
        exitTime: timeStrNew,
        duration: durLabelNew,
        durationMins: durationMinsNew,
        location: b.address || (latNew !== null ? latNew + ', ' + lngNew : 'not captured'),
        distance: xdistNew
      };
    }
    return { success: false, message: 'No entry record found for today (' + dateStrNew + '). Mark attendance first.' };

    var now     = new Date();
    var t       = tz();
    var dateStr = Utilities.formatDate(now, t, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, t, 'HH:mm:ss');

    var sheet   = getSheet(SH.ATTENDANCE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No attendance records found' };

    // Read all cols needed
    // Cols: 1=att_id, 2=user_id, 4=type, 5=entry_time, 6=exit_time, 7=date, 9=lat, 10=lng
    var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (String(row[1]).trim() !== String(b.userId).trim()) continue; // col 2 = user_id
      if (normDate(row[6], t) !== dateStr) continue;                   // col 7 = attendance_date
      if (!isCheckInAttendance(row[3])) continue;                      // col 4 = type = check-in
      if (String(row[5]).trim()) continue;                             // col 6 = exit_time already filled

      // Found today's entry row — compute duration
      var entryTimeStr = String(row[4] || '');
      var durationMins = '';
      if (entryTimeStr) {
        var entryDt = new Date(dateStr + 'T' + entryTimeStr);
        if (!isNaN(entryDt.getTime())) {
          durationMins = Math.max(0, Math.round((now - entryDt) / 60000));
        }
      }

      var xlat = b.latitude  ? parseFloat(b.latitude)  : '';
      var xlng = b.longitude ? parseFloat(b.longitude) : '';
      var track = resolveUserTrackingLocation(b.userId, dateStrNew);
      var xdist = (xlat !== '' && xlng !== '')
        ? haversine(xlat, xlng, track.anchorLat, track.anchorLng) : '';

      var lock = LockService.getScriptLock();
      lock.waitLock(6000);
      try {
        var sheetRow = i + 2; // +2 for header + 0-index

        // Batch write exit columns — one setValues call
        // Col 4=type, 6=exit_time, 9=lat, 10=lng, 11=loc_id, 12=address, 13=distance
        sheet.getRange(sheetRow, 4, 1, 1).setValue('check-out');
        sheet.getRange(sheetRow, 6, 1, 1).setValue(timeStr);    // exit_time (col 6 = attendance_time #2)

        if (xlat !== '') {
          sheet.getRange(sheetRow, 9, 1, 5).setValues([[xlat, xlng, '', b.address || '', xdist]]);
        }

        // Write to LocationMonitor for exit coordinates
        if (xlat !== '') {
          getSheet(SH.LOC_MONITOR).appendRow([
            genId('lm'), b.userId, xlat, xlng, xdist, now.toISOString()
          ]);
        }

        SpreadsheetApp.flush();
        invalidate(SH.ATTENDANCE);
      } finally { lock.releaseLock(); }

      var hrs  = durationMins !== '' ? Math.floor(durationMins / 60) : 0;
      var mins = durationMins !== '' ? durationMins % 60 : 0;
      var durLabel = durationMins !== '' ? (hrs > 0 ? hrs + 'h ' + mins + 'm' : mins + 'm') : '';

      return {
        success:      true,
        message:      '\u2713 Exit recorded at ' + timeStr + (durLabel ? ' \u00b7 ' + durLabel : ''),
        exitTime:     timeStr,
        duration:     durLabel,
        durationMins: durationMins,
        location:     b.address || (xlat ? xlat + ', ' + xlng : 'not captured'),
        distance:     xdist
      };
    }

    return {
      success: false,
      message: 'No entry record found for today (' + dateStr + '). Mark attendance first.'
    };
  } catch(err) { return { success: false, message: 'markExit: ' + err }; }
}

function trackStudentLocation(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };
    if (b.latitude === '' || b.latitude === null || typeof b.latitude === 'undefined')
      return { success: false, message: 'latitude required' };
    if (b.longitude === '' || b.longitude === null || typeof b.longitude === 'undefined')
      return { success: false, message: 'longitude required' };

    var lat = parseFloat(b.latitude);
    var lng = parseFloat(b.longitude);
    if (isNaN(lat) || isNaN(lng)) return { success: false, message: 'invalid coordinates' };

    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var latNew = parseNumber(b.latitude);
    var lngNew = parseNumber(b.longitude);
    var nowNew = new Date();
    var tNew = tz();
    var dateStrNew = Utilities.formatDate(nowNew, tNew, 'yyyy-MM-dd');
    var resolved = resolveEffectiveAttendanceLocation(user, currentGuid(), dateStrNew);
    var distNew = (latNew !== null && lngNew !== null)
      ? haversine(latNew, lngNew, resolved.anchorLat, resolved.anchorLng)
      : 0;
    getSheet(SH.LOC_MONITOR).appendRow([
      genId('lm'),
      b.userId,
      latNew,
      lngNew,
      distNew,
      nowNew.toISOString(),
      currentGuid()
    ]);

    if (distNew > (resolved.allowedDistance || DEFAULT_RADIUS)) {
      var exitResult = markExit({ userId: b.userId, latitude: latNew, longitude: lngNew, address: b.address || '' });
      return { success: true, distance: distNew, exitMarked: !!(exitResult && exitResult.success) };
    }

    return {
      success: true,
      message: 'Location tracked',
      distanceFromCentre: distNew,
      attendanceLocationId: resolved.attendanceLocationId
    };

    var now = new Date();
    var track = resolveUserTrackingLocation(b.userId);
    var dist = haversine(lat, lng, track.anchorLat, track.anchorLng);

    getSheet(SH.LOC_MONITOR).appendRow([
      genId('lm'),
      b.userId,
      lat,
      lng,
      dist,
      now.toISOString()
    ]);

    return {
      success: true,
      message: 'Location tracked',
      distanceFromCentre: dist,
      attendanceLocationId: track.attendanceLocationId
    };
  } catch(err) { return { success: false, message: 'trackStudentLocation: ' + err }; }
}

// ============================================================
//  5. GET MY ATTENDANCE
// ============================================================

function getMyAttendance(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };
    var t        = tz();
    var sheet    = getSheet(SH.ATTENDANCE);
    var lastRow  = sheet.getLastRow();
    if (lastRow < 2) return { success: true, records: [] };

    var data     = sheet.getRange(2, 1, lastRow - 1, attendanceColumnCount()).getValues();
    var entryMap = {}; // date → entry row
    var exitMap  = {}; // date → exit row

    data.forEach(function(row) {
      if (String(row[1]).trim() !== String(b.userId).trim()) return;
      var date = normDate(row[6], t);
      if (isCheckInAttendance(row[3])) entryMap[date] = row;
      if (isCheckOutAttendance(row[3]))  exitMap[date]  = row;
    });

    var records = Object.keys(entryMap).map(function(date) {
      var e = entryMap[date];
      var x = exitMap[date] || null;
      var eDur = '';
      if (x) {
        var eDt = new Date(date + 'T' + String(e[4]));
        var xDt = new Date(date + 'T' + String(x[5]));
        if (!isNaN(eDt.getTime()) && !isNaN(xDt.getTime())) {
          var dm = Math.max(0, Math.round((xDt - eDt) / 60000));
          var h  = Math.floor(dm / 60), m = dm % 60;
          eDur   = h > 0 ? h + 'h ' + m + 'm' : m + 'm';
        }
      }
      return {
        date:           date,
        entryTime:      e[4]  || '',
        exitTime:       x ? x[5] : '',
        duration:       eDur,
        loginMethod:    e[7]  || '',
        address:        e[11] || '',
        distanceFromCentre: e[12] || ''
      };
    });

    records.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(err) { return { success: false, message: 'getMyAttendance: ' + err }; }
}

// ============================================================
//  6. SESSIONS (teacher)
// ============================================================

function createSession(b) {
  try {
    var roleName = normalizeRoleValue(b.roleId);
    if (roleName !== 'teacher' && roleName !== 'admin')
      return { success: false, message: 'Only teachers can create sessions' };
    if (!b.subject || !b.windowMinutes)
      return { success: false, message: 'subject and windowMinutes required' };

    var t      = tz();
    var now    = new Date();
    var sessId = genId('s');
    var date   = Utilities.formatDate(now, t, 'yyyy-MM-dd');
    var start  = Utilities.formatDate(now, t, 'HH:mm:ss');
    var end    = Utilities.formatDate(new Date(now.getTime() + b.windowMinutes * 60000), t, 'HH:mm:ss');

    var sheet  = getSheet(SH.SESSIONS);
    // Close existing open sessions for this teacher
    var data   = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(b.userId) && data[i][7] === 'open')
        sheet.getRange(i + 1, 8).setValue('closed');
    }
    sheet.appendRow([sessId, b.userId, b.teacherName || '', b.subject, date, start, end, 'open', b.windowMinutes]);

    // Cache for fast student polling
    CacheService.getScriptCache().put('active_session', JSON.stringify({
      session_id: sessId, teacher_id: b.userId, subject: b.subject,
      date: date, start_time: start, end_time: end, status: 'open', window_minutes: b.windowMinutes
    }), TTL_SESSION);

    return { success: true, sessionId: sessId, subject: b.subject, startTime: start, endTime: end };
  } catch(err) { return { success: false, message: 'createSession: ' + err }; }
}

function closeSession(b) {
  try {
    var sheet = getSheet(SH.SESSIONS);
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(b.sessionId)) {
        sheet.getRange(i + 1, 8).setValue('closed');
        CacheService.getScriptCache().remove('active_session');
        return { success: true };
      }
    }
    return { success: false, message: 'Session not found' };
  } catch(err) { return { success: false, message: 'closeSession: ' + err }; }
}

function getActiveSession(b) {
  try {
    // CacheService first — avoids sheet read on every 30s student poll
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('active_session');
    if (cached) {
      var s   = JSON.parse(cached);
      var now = new Date(), t = tz();
      var today = Utilities.formatDate(now, t, 'yyyy-MM-dd');
      if (s.status === 'open' && s.date === today) {
        var endDt = new Date(s.date + 'T' + s.end_time);
        if (now <= endDt)
          return { success: true, active: true, session: s,
                   secondsLeft: Math.max(0, Math.round((endDt - now) / 1000)) };
      }
      cache.remove('active_session');
    }

    var autoSession = getCurrentAutoSession(new Date());
    if (autoSession) {
      var autoEnd = new Date(autoSession.date + 'T' + autoSession.end_time);
      return {
        success: true,
        active: true,
        session: autoSession,
        secondsLeft: Math.max(0, Math.round((autoEnd - new Date()) / 1000))
      };
    }

    // Cache miss — read sheet
    var t2    = tz();
    var today2 = Utilities.formatDate(new Date(), t2, 'yyyy-MM-dd');
    var rows  = getRows(getSheet(SH.SESSIONS));
    var now2  = new Date();
    for (var i = 0; i < rows.length; i++) {
      var s = rows[i];
      if (s.status === 'open' && normDate(s.date, t2) === today2) {
        var st = new Date(s.date + 'T' + s.start_time);
        var en = new Date(s.date + 'T' + s.end_time);
        if (now2 >= st && now2 <= en) {
          var secs = Math.max(0, Math.round((en - now2) / 1000));
          cache.put('active_session', JSON.stringify(s), TTL_SESSION);
          return { success: true, active: true, session: s, secondsLeft: secs };
        }
      }
    }
    return { success: true, active: false };
  } catch(err) { return { success: false, message: 'getActiveSession: ' + err }; }
}

function getSessions(b) {
  try {
    var t    = tz();
    var rows = getRows(getSheet(SH.SESSIONS));
    var today = Utilities.formatDate(new Date(), t, 'yyyy-MM-dd');
    var autoRows = getAutoSessionsForDate(today);
    rows = autoRows.concat(rows);
    if (b.userId) {
      rows = rows.filter(function(r) {
        return String(r.teacher_id) === String(b.userId) || String(r.teacher_id) === 'ALL';
      });
    }

    // Count present students per day
    var attSheet = getSheet(SH.ATTENDANCE);
    var attLast  = attSheet.getLastRow();
    var countMap = {};
    if (attLast >= 2) {
      var uids  = attSheet.getRange(2, 2, attLast - 1, 1).getValues();
      var dates = attSheet.getRange(2, 7, attLast - 1, 1).getValues();
      var types = attSheet.getRange(2, 4, attLast - 1, 1).getValues();
      uids.forEach(function(u, i) {
        if (!isCheckInAttendance(types[i][0])) return;
        var key = normDate(dates[i][0], t);
        countMap[key] = (countMap[key] || 0) + 1;
      });
    }

    return {
      success:  true,
      sessions: rows.map(function(s) {
        var status = s.status;
        if (s.is_auto) {
          var st = new Date(s.date + 'T' + s.start_time);
          var en = new Date(s.date + 'T' + s.end_time);
          var now = new Date();
          status = now >= st && now <= en ? 'open' : 'scheduled';
        }
        return {
          sessionId:    s.session_id,
          subject:      s.subject,
          date:         normDate(s.date, t),
          startTime:    s.start_time,
          endTime:      s.end_time,
          status:       status,
          presentCount: countMap[normDate(s.date, t)] || 0
        };
      }).reverse().slice(0, 20)
    };
  } catch(err) { return { success: false, message: 'getSessions: ' + err }; }
}

function getTeacherNotifications(b) {
  try {
    if (!b.sessionId) return { success: false, message: 'sessionId required' };
    var cache = CacheService.getScriptCache();
    var key = 'teacher_notif_' + b.sessionId;
    var rows = [];
    var cached = cache.get(key);
    if (cached) {
      try { rows = JSON.parse(cached) || []; } catch(e) {}
    }
    if (b.teacherId) {
      rows = rows.filter(function(r) {
        var teacherId = String(r.teacherId || '');
        return teacherId === String(b.teacherId) || teacherId === 'ALL';
      });
    }
    return { success: true, notifications: rows };
  } catch(err) { return { success: false, message: 'getTeacherNotifications: ' + err }; }
}

// ============================================================
//  7. DASHBOARD (cached 60s)
// ============================================================

function getDashboard(b) {
  try {
    if (!b.sessionId) return { success: false, message: 'sessionId required' };

    var cacheKey = 'dash_' + b.sessionId;
    var cached   = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);

    var t         = tz();
    var today     = Utilities.formatDate(new Date(), t, 'yyyy-MM-dd');
    var userRows  = getCached(SH.USERS, TTL_LOOKUP);
    var studentRoleId = findRoleIdByName('student');
    var students  = userRows.filter(function(u) {
      return normalizeRoleValue(u.role_id) === 'student' || (studentRoleId && String(u.role_id) === studentRoleId);
    });
    var userMap   = buildMap(students, 'user_id');

    // Narrow read — entry rows for today only
    var attSheet  = getSheet(SH.ATTENDANCE);
    var attLast   = attSheet.getLastRow();
    var presentMap = {};
    if (attLast >= 2) {
      var rows = attSheet.getRange(2, 1, attLast - 1, attendanceColumnCount()).getValues();
      rows.forEach(function(row) {
        if (!isCheckInAttendance(row[3])) return;
        if (normDate(row[6], t) !== today) return;
        presentMap[String(row[1]).trim()] = {
          entryTime:  row[4]  || '',
          exitTime:   row[5]  || '',
          method:     row[7]  || '',
          address:    row[11] || '',
          distance:   row[12] || ''
        };
      });
    }

    var present = [], absent = [];
    students.forEach(function(s) {
      var uid = String(s.user_id).trim();
      if (presentMap[uid]) {
        var pm = presentMap[uid];
        present.push({
          userId:     uid,
          name:       s.full_name,
          email:      s.email,
          category:   s.department_id,
          subcategory: s.subcategory_id || '',
          department: s.department_id,
          entryTime:  pm.entryTime,
          exitTime:   pm.exitTime,
          method:     pm.method,
          address:    pm.address,
          distance:   pm.distance
        });
      } else {
        absent.push({
          userId: uid,
          name: s.full_name,
          email: s.email,
          category: s.department_id,
          subcategory: s.subcategory_id || '',
          department: s.department_id
        });
      }
    });

    var result = {
      success: true, total: students.length,
      presentCount: present.length, absentCount: absent.length,
      present: present, absent: absent
    };
    try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), TTL_DASH); } catch(e) {}
    return result;
  } catch(err) { return { success: false, message: 'getDashboard: ' + err }; }
}

// ============================================================
//  8. STUDENTS ROSTER
// ============================================================

function getUsers(b) {
  try {
    var roleId = String(b && (b.roleId || b.role || b.roleKey) || '').trim();
    var roleFilter = normalizeRoleValue(roleId || '');
    var rows = getCached(SH.USERS, TTL_LOOKUP)
      .filter(function(r) {
        if (!roleFilter) return true;
        return normalizeRoleValue(r.role_id) === roleFilter || String(r.role_id) === roleId;
      })
      .map(function(r) {
        return {
          userId:     r.user_id,
          name:       r.full_name,
          email:      r.email,
          category:   r.department_id,
          subcategory: r.subcategory_id || '',
          department: r.department_id,
          mobile:     r.mobile,
          roleId:     r.role_id,
          hasBio:     !!r.biometric_code,
          hasDevice:  !!r.device_identification
        };
      });
    return { success: true, users: rows, students: rows, total: rows.length };
  } catch(err) { return { success: false, message: 'getUsers: ' + err }; }
}

function getStudents(b) {
  return getUsers({ roleId: 'student' });
}

// ============================================================
//  9. EXPORT ATTENDANCE (CSV)
// ============================================================

function exportAttendance(b) {
  try {
    var t       = tz();
    var sheet   = getSheet(SH.ATTENDANCE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, csv: '', rowCount: 0 };

    var data    = sheet.getRange(2, 1, lastRow - 1, attendanceColumnCount()).getValues();

    var header = ['attendance_id','user_id','full_name','type_attendance',
                  'entry_time','exit_time','attendance_date','login_method',
                  'latitude','longitude','attendance_location_id','address','distance_from_centre',
                  'tenant_id','category_id','subcategory_id'];
    var lines  = [header.join(',')];

    data.forEach(function(row) {
      if (b.userId && String(row[1]).trim() !== String(b.userId).trim()) return;
      if (b.date   && normDate(row[6], t) !== b.date) return;
      lines.push(row.map(function(c) { return '"' + String(c || '').replace(/"/g, '""') + '"'; }).join(','));
    });

    return { success: true, csv: lines.join('\n'), rowCount: lines.length - 1 };
  } catch(err) { return { success: false, message: 'exportAttendance: ' + err }; }
}

// ============================================================
//  10. BIOMETRIC & DEVICE
// ============================================================

function getBiometric(b) {
  try {
    var user = getUserByEmail(b.email);
    if (!user) return { success: false, message: 'No account found. Please register first.' };
    if (!user.biometric_code)
      return { success: false, message: 'No biometric registered. Please register fingerprint first.' };
    return { success: true, credentialId: user.biometric_code, userId: user.user_id, name: user.full_name };
  } catch(err) { return { success: false, message: 'getBiometric: ' + err }; }
}

function saveBiometric(b) {
  try {
    var sheet = getSheet(SH.USERS);
    var last  = sheet.getLastRow();
    // Col 1 = user_id, Col 10 = biometric_code
    var uids  = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < uids.length; i++) {
      if (String(uids[i][0]) === String(b.userId)) {
        sheet.getRange(i + 2, 10).setValue(b.credentialId); // col 10 = biometric_code
        invalidate(SH.USERS);
        return { success: true };
      }
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: 'saveBiometric: ' + err }; }
}

function registerDevice(b) {
  try {
    if (!b.userId || !b.deviceId) return { success: false, message: 'userId and deviceId required' };
    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var stored = String(user.device_identification || '').trim();
    if (stored && stored !== String(b.deviceId).trim())
      return { success: false, alreadyBound: true, message: 'Account already bound to another device.' };

    var sheet = getSheet(SH.USERS);
    var last  = sheet.getLastRow();
    var uids  = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < uids.length; i++) {
      if (String(uids[i][0]) === String(b.userId)) {
        sheet.getRange(i + 2, 11).setValue(b.deviceId); // col 11 = device_identification
        invalidate(SH.USERS);
        return { success: true, message: stored ? 'Device confirmed' : 'Device registered' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch(err) { return { success: false, message: 'registerDevice: ' + err }; }
}

function checkDevice(b) {
  try {
    if (!b.userId || !b.deviceId) return { success: false, message: 'userId and deviceId required' };
    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var stored = String(user.device_identification || '').trim();
    if (!stored) return { success: true, status: 'unbound' };
    if (stored === String(b.deviceId).trim()) return { success: true, status: 'match' };
    return { success: false, status: 'mismatch', message: 'Account registered to a different device.' };
  } catch(err) { return { success: false, message: 'checkDevice: ' + err }; }
}

// ============================================================
//  11. LOOKUP TABLES (read-only, cached)
// ============================================================

function getLookup(sheetName) {
  try {
    return { success: true, data: getCached(sheetName, TTL_LOOKUP) };
  } catch(err) { return { success: false, message: 'getLookup: ' + err }; }
}

function attendanceColumnCount() {
  return HEADERS.Attendance.length;
}

function getUserLocMap(b) {
  try {
    var rows = getRows(getSheet(SH.USER_LOC_MAP));
    if (b.userId) rows = rows.filter(function(r) { return String(r.user_id) === String(b.userId); });
    return { success: true, data: rows };
  } catch(err) { return { success: false, message: 'getUserLocMap: ' + err }; }
}

function getCategoryLocationMap(b) {
  try {
    var rows = getRows(getSheet(SH.CATEGORY_LOC_MAP));
    if (b.categoryId) rows = rows.filter(function(r) { return String(r.category_id) === String(b.categoryId); });
    if (b.subcategoryId) rows = rows.filter(function(r) { return String(r.subcategory_id) === String(b.subcategoryId); });
    return { success: true, data: rows };
  } catch(err) { return { success: false, message: 'getCategoryLocationMap: ' + err }; }
}

// ============================================================
//  ADMIN WRITE FUNCTIONS (called from UI Admin tab)
// ============================================================

// Add a Department row
function addDepartment(b) {
  try {
    if (!b.departmentId || !b.name)
      return { success: false, message: 'category_id and name required' };
    // Check duplicate
    var rows = getRows(getSheet(SH.DEPARTMENTS));
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].department_id) === String(b.departmentId))
        return { success: false, message: 'category_id already exists' };
    }
    getSheet(SH.DEPARTMENTS).appendRow([b.departmentId, b.name, b.inCharge||'', b.email||'', currentGuid()]);
    invalidate(SH.DEPARTMENTS);
    return { success: true, message: 'Category added' };
  } catch(err) { return { success: false, message: 'addDepartment: ' + err }; }
}

// Add an AttendanceLocation row
function addAttendanceLocation(b) {
  try {
    if (!b.name || !b.latitude || !b.longitude)
      return { success: false, message: 'name, latitude and longitude required' };
    var lat = parseNumber(b.latitude);
    var lng = parseNumber(b.longitude);
    if (lat === null || lng === null) return { success: false, message: 'Valid latitude and longitude required' };
    var radius = parseInt(b.geofenceRadius || b.geofence_radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
    var address = toCleanText(b.address || '', 240);
    var tenantId = currentGuid();
    if (String(b.reuseLocationId || '').trim()) {
      var reuse = findLocationById(b.reuseLocationId, tenantId);
      if (!reuse) return { success: false, message: 'Existing location not found for reuse' };
      return { success: true, reused: true, locationId: reuse.attendance_location_id, message: 'Reused existing location' };
    }
    var nearby = findNearbyAttendanceLocations(lat, lng, tenantId);
    if (nearby.length && String(b.confirmDuplicate || '').toLowerCase() !== 'true') {
      return {
        success: false,
        duplicateWarning: true,
        message: 'A location already exists within 20 meters.',
        nearbyLocations: nearby
      };
    }
    var rows = getRows(getSheet(SH.ATT_LOCATIONS));
    var locId = String(b.locationId || b.attendance_location_id || genNumericId()).trim();
    if (!/^\d+$/.test(locId)) return { success: false, message: 'attendance_location_id must be numeric' };
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].attendance_location_id) === String(locId))
        return { success: false, message: 'attendance_location_id already exists' };
    }
    getSheet(SH.ATT_LOCATIONS).appendRow([locId, b.name, address, lat, lng, radius, tenantId]);
    invalidate(SH.ATT_LOCATIONS);
    return { success: true, locationId: locId, message: nearby.length ? 'Location added after confirmation' : 'Location added' };
  } catch(err) { return { success: false, message: 'addAttendanceLocation: ' + err }; }
}

// Add a CategoryLocationMap row
function addUserLocMap(b) {
  try {
    return addCategoryLocationMap(b);
  } catch(err) { return { success: false, message: 'addUserLocMap: ' + err }; }
}

function addCategoryLocationMap(b) {
  try {
    var categoryId = String(b.categoryId || b.departmentId || '').trim();
    var subcategoryId = String(b.subcategoryId || '').trim();
    var locationId = String(b.locationId || b.attendanceLocationId || '').trim();
    if (!categoryId || !subcategoryId || !locationId)
      return { success: false, message: 'categoryId, subcategoryId and locationId required' };
    var loc = findLocationById(locationId, currentGuid());
    if (!loc) return { success: false, message: 'attendance location not found' };
    var mapId = genId('clm');
    getSheet(SH.CATEGORY_LOC_MAP).appendRow([mapId, categoryId, subcategoryId, locationId, parseInt(b.allowedDistance || loc.geofence_radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS, currentGuid(), String(b.status || 'active')]);
    invalidate(SH.CATEGORY_LOC_MAP);
    return { success: true, mapId: mapId, message: 'Mapping added' };
  } catch(err) { return { success: false, message: 'addCategoryLocationMap: ' + err }; }
}

function createWorkLocationRequest(b) {
  try {
    var userId = String(b.userId || '').trim();
    var requestDateRaw = String(b.requestDate || b.date || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(requestDateRaw))
      return { success: false, message: 'requestDate required' };
    var requestDate = normDate(requestDateRaw, tz());
    var locationName = toCleanText(b.locationName || b.name || '', 120);
    var address = toCleanText(b.address || '', 240);
    var reason = toCleanText(b.reason || '', 240);
    var lat = parseNumber(b.latitude);
    var lng = parseNumber(b.longitude);
    var allowedDistance = parseInt(b.allowedDistance || b.allowed_distance || b.radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
    if (!userId) return { success: false, message: 'userId required' };
    if (!requestDate) return { success: false, message: 'requestDate required' };
    if (!locationName) return { success: false, message: 'locationName required' };
    if (lat === null || lng === null) return { success: false, message: 'valid latitude and longitude required' };
    if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return { success: false, message: 'latitude/longitude out of range' };
    var user = getUserById(userId);
    if (!user) return { success: false, message: 'User not found' };
    if (getWorkLocationRequestForDate(userId, requestDate, currentGuid())) {
      return { success: false, message: 'Only one request is allowed per user per day' };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      if (getWorkLocationRequestForDate(userId, requestDate, currentGuid())) {
        return { success: false, message: 'Only one request is allowed per user per day' };
      }
      var now = new Date();
      var reqId = genNumericId();
      getSheet(SH.WORK_LOC_REQUESTS).appendRow([
        reqId,
        userId,
        String(user.full_name || ''),
        requestDate,
        locationName,
        address,
        lat,
        lng,
        allowedDistance,
        reason,
        'pending',
        '',
        '',
        '',
        '',
        currentGuid(),
        now.toISOString(),
        now.toISOString()
      ]);
      invalidate(SH.WORK_LOC_REQUESTS);
      appendAuditLog('create_work_location_request', userId, normalizeRoleValue(user.role_id), userId, {
        requestId: reqId,
        requestDate: requestDate,
        locationName: locationName,
        latitude: lat,
        longitude: lng,
        allowedDistance: allowedDistance
      });
      return { success: true, requestId: reqId, status: 'pending', message: 'Work location request submitted' };
    } finally { lock.releaseLock(); }
  } catch (err) { return { success: false, message: 'createWorkLocationRequest: ' + err }; }
}

function getWorkLocationRequests(b) {
  try {
    var rows = getRows(getSheet(SH.WORK_LOC_REQUESTS));
    var userId = String(b && b.userId || '').trim();
    var status = String(b && (b.status || b.requestStatus) || '').trim().toLowerCase();
    var requestDate = String(b && (b.requestDate || b.date) || '').trim();
    if (userId) rows = rows.filter(function(r) { return String(r.user_id || '').trim() === userId; });
    if (status) rows = rows.filter(function(r) { return String(r.status || '').toLowerCase() === status; });
    if (requestDate) rows = rows.filter(function(r) { return normDate(r.request_date, tz()) === requestDate; });
    rows.sort(function(a, b) {
      var ad = String(a.request_date || '');
      var bd = String(b.request_date || '');
      if (ad === bd) return String(b.created_at || '').localeCompare(String(a.created_at || ''));
      return bd.localeCompare(ad);
    });
    return { success: true, data: rows };
  } catch (err) { return { success: false, message: 'getWorkLocationRequests: ' + err }; }
}

function findWorkRequestRowById(requestId) {
  var sheet = getSheet(SH.WORK_LOC_REQUESTS);
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(requestId).trim()) {
      return { rowNum: i + 2, row: data[i] };
    }
  }
  return null;
}

function updateWorkRequestStatus(b, newStatus) {
  try {
    var requestId = String(b.requestId || '').trim();
    if (!requestId) return { success: false, message: 'requestId required' };
    var found = findWorkRequestRowById(requestId);
    if (!found) return { success: false, message: 'Request not found' };
    var currentStatus = String(found.row[10] || '').toLowerCase();
    if (currentStatus === newStatus) return { success: true, requestId: requestId, status: newStatus, message: 'Request already ' + newStatus };

    var now = new Date().toISOString();
    var adminName = toCleanText(b.approvedBy || b.adminEmail || b.approver || '', 160);
    var sheet = getSheet(SH.WORK_LOC_REQUESTS);
    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var fresh = sheet.getRange(found.rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
      var freshStatus = String(fresh[10] || '').toLowerCase();
      if (freshStatus === newStatus) return { success: true, requestId: requestId, status: newStatus, message: 'Request already ' + newStatus };
      sheet.getRange(found.rowNum, 11).setValue(newStatus);
      sheet.getRange(found.rowNum, 12, 1, 4).setValues([[
        newStatus === 'approved' ? adminName : '',
        newStatus === 'approved' ? now : '',
        newStatus === 'rejected' ? adminName : '',
        newStatus === 'rejected' ? now : ''
      ]]);
      sheet.getRange(found.rowNum, 18).setValue(now);
      invalidate(SH.WORK_LOC_REQUESTS);
      return { success: true, requestId: requestId, status: newStatus, message: 'Request ' + newStatus };
    } finally { lock.releaseLock(); }
  } catch (err) { return { success: false, message: 'updateWorkRequestStatus: ' + err }; }
}

function approveWorkLocationRequest(b) { return updateWorkRequestStatus(b, 'approved'); }
function rejectWorkLocationRequest(b) { return updateWorkRequestStatus(b, 'rejected'); }

function getDepartments() { return getLookup(SH.DEPARTMENTS); }
function getRoles() { return getLookup(SH.ROLES); }
function getLocations() { return getLookup(SH.ATT_LOCATIONS); }

function markAttendance(b) {
  var mode = normalizeAttendanceType(b.type_attendance || b.attendanceType || b.type || 'check-in');
  if (mode === 'check-out') return markExit(b);
  return markEntry(b);
}

// ============================================================
//  12. SETUP
// ============================================================

function setupSheets() {
  try {
    var created = [];

    // Create all sheets
    Object.keys(SH).forEach(function(key) { getSheet(SH[key]); created.push(SH[key]); });

    ensureExactRows('Roles', [
      [1, 'admin'],
      [2, 'teacher'],
      [3, 'student'],
      [4, 'employee'],
      [5, 'manager']
    ]);

    ensureExactRows('Departments', [
      ['CSE', 'Computer Science & Engineering', '', '', currentGuid()],
      ['ISE', 'Information Science & Engineering', '', '', currentGuid()],
      ['ECE', 'Electronics & Communication Engineering', '', '', currentGuid()],
      ['EEE', 'Electrical & Electronics Engineering', '', '', currentGuid()],
      ['MECH', 'Mechanical Engineering', '', '', currentGuid()],
      ['CIVIL', 'Civil Engineering', '', '', currentGuid()],
      ['CHEM', 'Chemical Engineering', '', '', currentGuid()],
      ['BT', 'Biotechnology', '', '', currentGuid()],
      ['MBA', 'MBA', '', '', currentGuid()],
      ['MCA', 'MCA', '', '', currentGuid()]
    ]);

    ensureExactRows('AttendanceType', [
      [1, 'check-in'],
      [2, 'check-out']
    ]);

    // ── Seed AttendanceType ──
    var typeSheet = getSheet(SH.ATT_TYPE);
    if (typeSheet.getLastRow() < 2) {
      typeSheet.appendRow(['ATT_TYPE_001', 'check-in']);
      typeSheet.appendRow(['ATT_TYPE_002', 'check-out']);
    }

    // ── Seed AttendanceLocations ──
    var locSheet = getSheet(SH.ATT_LOCATIONS);
    if (locSheet.getLastRow() < 2) {
      var tenantId = currentGuid();
      locSheet.appendRow(['1', 'SIT Campus Main Block', 'Main administrative block', 13.32603, 77.12621, 200, tenantId]);
      locSheet.appendRow(['2', 'SIT Campus CS Lab Block', 'Computer science lab block', 13.32620, 77.12650, 200, tenantId]);
      locSheet.appendRow(['3', 'SIT Campus Seminar Hall', 'Seminar hall', 13.32580, 77.12600, 200, tenantId]);
    }

    getSheet(SH.WORK_LOC_REQUESTS);

    return {
      success: true,
      message: 'All sheets created and seeded.',
      sheets:  created
    };
  } catch(err) { return { success: false, message: 'setupSheets: ' + err }; }
}

// ============================================================
//  13. DEBUG
// ============================================================

function debugInfo() {
  try {
    var sheets = ss().getSheets().map(function(s) {
      return { name: s.getName(), dataRows: Math.max(0, s.getLastRow() - 1) };
    });
    return {
      success:     true,
      spreadsheet: ss().getName(),
      sheets:      sheets,
      geofence:    DEFAULT_RADIUS + 'm around (' + DEFAULT_LAT + ', ' + DEFAULT_LNG + ')'
    };
  } catch(err) { return { success: false, message: 'debug: ' + err }; }
}
