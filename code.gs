// ============================================================
//  BioAttend – Google Apps Script Backend  v7
//  College  : Siddaganga Institute of Technology, Tumkur
//  Deploy   : Web App → Execute as Me → Access: Anyone
//
//  ALL IDs are now auto-incrementing integers (1, 2, 3, …)
//  across every table — Users, Departments, Roles, Attendance,
//  LocationMonitor, AttendanceType, AttendanceLocations,
//  UserAttendanceLocationMap, Sessions, AuthSessions, AuditLog,
//  AttendanceWindows, UserIndex.
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
  SESSIONS     : 'Sessions',
  USER_INDEX   : 'UserIndex',
  AUTH_SESSIONS: 'AuthSessions',
  AUDIT_LOG    : 'AuditLog',
  ATT_WINDOWS  : 'AttendanceWindows',
  RATE_LIMIT   : 'RateLimit',
  WEBAUTHN     : 'WebAuthnChallenges',
  WEBAUTHN_CRED: 'WebAuthnCredentials'
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
    'address', 'distance_from_centre', 'status', 'tenant_id',
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
  Sessions: [
    'session_id', 'teacher_id', 'teacher_name', 'subject',
    'date', 'start_time', 'end_time', 'status', 'window_minutes', 'tenant_id'
  ],
  UserIndex: [
    'user_id', 'email', 'row_number', 'tenant_id'
  ],
  AuthSessions: [
    'session_id', 'user_id', 'token', 'refresh_token', 'device_id',
    'created_at', 'expires_at', 'refresh_expires_at', 'status',
    'guid', 'role_name', 'last_seen_at', 'tenant_id'
  ],
  AuditLog: [
    'audit_id', 'event_type', 'actor_user_id', 'actor_role', 'target_user_id',
    'details', 'ip_hint', 'created_at', 'tenant_id'
  ],
  AttendanceWindows: [
    'window_id', 'name', 'start_time', 'duration_minutes', 'status', 'location_id'
  ],
  RateLimit: [
    'key', 'count', 'last_request_time', 'window_start'
  ],
  WebAuthnChallenges: [
    'challenge_id', 'user_id', 'action', 'challenge', 'created_at',
    'expires_at', 'status', 'metadata'
  ],
  WebAuthnCredentials: [
    'credential_id', 'user_id', 'public_key', 'created_at', 'status', 'guid'
  ]
};

var DEFAULT_GUID = '1';
var REQUEST_CONTEXT = { guid: DEFAULT_GUID };
var TENANT_DIRECTORY = {
  '1': {
    guid: '1',
    orgType: 'college',
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/1g43D4SH7DnYAh05trU5t5xhfyVxP43unO4knX8WdA9k/edit?gid=0#gid=0',
    institution: {
      name: 'BioAttend Main',
      city: 'Tumakuru',
      logoUrl: '',
      website: '',
      address: 'Main tenant workspace'
    },
    application: {
      id: 101,
      name: 'Attendance monitoring',
      description: 'Multi-tenant biometric attendance monitoring'
    },
    apiUrl: 'https://script.google.com/macros/s/AKfycbwhFJ7oyLoed11sTYGikHyExxYs20J842q244K0MJ0VfwL5KgMDTb7E3uMN2sWhj0njYg/exec'
  },
  '2': {
    guid: '2',
    orgType: 'college',
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/1c5ZGKxr-ZNxbao6L6D87N8JNSvPNf2Il9Xnl8ha04f8/edit?gid=0#gid=0',
    institution: {
      name: 'SIT',
      city: 'Tumakuru',
      logoUrl: '',
      website: '',
      address: 'Siddaganga Institute of Technology'
    },
    application: {
      id: 102,
      name: 'Attendance monitoring',
      description: 'Biometric attendance for students, teachers and employees'
    },
    apiUrl: 'https://script.google.com/macros/s/AKfycbzR-z38NrPZZm--4OeStiRvAgMb6SpwCjtb_GW0Rl9-/dev'
  },
  '3': {
    guid: '3',
    orgType: 'college',
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/1yGS6eyC6NTqwu6dllTYALe_Sc_0hn4_furgU-Wq6bH4/edit?gid=0#gid=0',
    institution: {
      name: 'SSIT',
      city: 'Tumakuru',
      logoUrl: '',
      website: '',
      address: 'Sri Siddhartha Institute of Technology'
    },
    application: {
      id: 103,
      name: 'Attendance monitoring',
      description: 'Biometric attendance for multi-tenant institutions'
    },
    apiUrl: 'https://script.google.com/macros/s/AKfycbwlTZpSSvpBZkJsOK5GelxsX2GOMH5E6M2HdSk43N4/dev'
  }
};

// ── GPS defaults ──────────────────────────────────────────────
var DEFAULT_LAT = 13.32609;
var DEFAULT_LNG = 77.12623;
var DEFAULT_RADIUS = 2000;
var TENANT_GEOFENCE = {
  '2': { latitude: 13.32609, longitude: 77.12623, radius: 2000, label: 'SIT' },
  '3': { latitude: 13.32048, longitude: 77.09173, radius: 2000, label: 'SSIT' }
};
var AUTO_SESSION_WINDOW_MINUTES = 10;
var AUTO_SESSION_SUBJECT = 'Automatic Attendance Session';
var ACCESS_TOKEN_TTL_MINUTES = 15;
var REFRESH_TOKEN_TTL_DAYS = 7;
var AUTH_SESSION_TTL_MINUTES = 1440;
var RATE_LIMIT_MAX_PER_MINUTE = 5;
var PASSWORD_ITERATIONS = 12000;

// ── Cache TTLs ────────────────────────────────────────────────
var TTL_LOOKUP  = 600;
var TTL_DASH    = 60;
var TTL_SESSION = 30;

// ============================================================
//  NUMERIC AUTO-INCREMENT ID GENERATOR
//  Reads the last row of any sheet and returns max(id_col) + 1
//  Works for every table — guarantees 1, 2, 3, … integers
// ============================================================

/**
 * Returns the next integer ID for a given sheet.
 * @param {string} sheetName  - sheet to inspect
 * @param {number} idCol      - 1-based column index of the id column (default 1)
 */
function nextId(sheetName, idCol) {
  idCol = idCol || 1;
  var sheet = getSheet(sheetName);
  var last  = sheet.getLastRow();
  if (last < 2) return 1;
  var vals = sheet.getRange(2, idCol, last - 1, 1).getValues();
  var max  = 0;
  vals.forEach(function(r) {
    var n = parseInt(r[0], 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

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

function respondOut(obj, callback) {
  return callback ? jsonpOut(obj, callback) : jsonOut(obj);
}

function doGet(e) {
  try {
    var p = e && e.parameter ? e.parameter : {};
    if (p.guid) setRequestGuid(p.guid);
    ensureCoreTenantData();
    var callback = p.callback || '';
    if (p.getApplicationFromGuid)
      return respondOut(getOrgByGUID(p.getApplicationFromGuid), callback);
    if (p.getOrgByGUID || p.action === 'getOrgByGUID')
      return respondOut(getOrgByGUID(p.guid || p.getOrgByGUID), callback);
    if (p.action === 'getRolesDepartments')
      return respondOut(getRolesDepartments(p.guid), callback);
    if (!p.data) return respondOut({ status: 'BioAttend v7 running', time: new Date().toString() }, callback);
    return respondOut(route(JSON.parse(decodeURIComponent(p.data))), callback);
  } catch(err) { return jsonOut({ success: false, message: 'doGet: ' + err }); }
}

function doPost(e) {
  try {
    ensureCoreTenantData();
    return jsonOut(route(JSON.parse(e.postData.contents)));
  }
  catch(err) { return jsonOut({ success: false, message: 'doPost: ' + err }); }
}

// ============================================================
//  ROUTER
// ============================================================

function route(b) {
  b = b || {};
  setRequestGuid(b.guid);
  ensureCoreTenantData();
  var action = String(b.action || '').trim();
  var rateLimitedActions = {
    register: true,
    registerUser: true,
    signIn: true,
    loginUser: true,
    loginByBiometric: true,
    refreshToken: true,
    beginWebAuthnRegistration: true,
    finishWebAuthnRegistration: true,
    beginWebAuthnLogin: true,
    finishWebAuthnLogin: true
  };
  if (rateLimitedActions[action]) {
    var rl = checkRateLimit(makeRateLimitKey(action, b));
    if (!rl.ok) return rl.error;
    incrementRateLimit(makeRateLimitKey(action, b));
  }
  var publicActions = {
    register: true,
    registerUser: true,
    signIn: true,
    loginUser: true,
    loginByBiometric: true,
    refreshToken: true,
    beginWebAuthnRegistration: true,
    finishWebAuthnRegistration: true,
    beginWebAuthnLogin: true,
    finishWebAuthnLogin: true,
    getOrgByGUID: true,
    getRolesDepartments: true,
    getRoles: true,
    getDepartments: true,
    getAttendanceTypes: true,
    getLocations: true,
    getAttendanceLocations: true,
    getBiometric: true,
    saveBiometric: true,
    checkBiometric: true,
    debug: false
  };
  if (!publicActions[action]) {
    var authCheck = validateToken(b);
    if (!authCheck.ok) return authCheck.error;
    var authz = authorize(authCheck.user, action, b);
    if (!authz.ok) return authz.error;
    if (authCheck.session && authCheck.session.guid) setRequestGuid(authCheck.session.guid);
    b.authSession = authCheck.session;
    b.authUser = authCheck.user;
  }
  switch (b.action) {
    case 'register':              return registerUser(b);
    case 'registerUser':          return registerUser(b);
    case 'signIn':                return signInUser(b);
    case 'loginUser':             return signInUser(b);
    case 'loginByBiometric':      return loginByBiometric(b);
    case 'refreshToken':          return refreshAuthSession(b);
    case 'saveBiometric':         return saveBiometric(b);
    case 'getBiometric':          return getBiometric(b);
    case 'checkBiometric':        return checkBiometric(b);
    case 'registerDevice':        return registerDevice(b);
    case 'checkDevice':           return checkDevice(b);
    case 'logout':                return logoutUser(b);
    case 'logoutUser':            return logoutUser(b);
    case 'beginWebAuthnRegistration': return beginWebAuthnRegistration(b);
    case 'finishWebAuthnRegistration': return finishWebAuthnRegistration(b);
    case 'beginWebAuthnLogin':        return beginWebAuthnLogin(b);
    case 'finishWebAuthnLogin':       return finishWebAuthnLogin(b);

    case 'markEntry':             return markEntry(b);
    case 'markAttendance':        return markEntry(b);
    case 'markExit':              return markExit(b);
    case 'exitAttendance':        return markExit(b);
    case 'trackLocation':         return trackLocation(b);
    case 'trackStudentLocation':  return trackLocation(b);
    case 'getMyAttendance':       return getMyAttendance(b);
    case 'exportAttendance':      return exportAttendance(b);

    case 'createSession':         return createSession(b);
    case 'closeSession':          return closeSession(b);
    case 'getActiveSession':      return getActiveSession(b);
    case 'getSessions':           return getSessions(b);
    case 'getTeacherNotifications': return getTeacherNotifications(b);

    case 'getDashboard':          return getDashboard(b);
    case 'getStudents':           return getStudents(b);
    case 'getLiveAttendance':     return getLiveAttendance(b);
    case 'getLiveUpdates':        return getLiveUpdates(b);
    case 'getUserLocations':      return getUserLocations(b);
    case 'forceExitUser':         return forceExitUser(b);
    case 'getDailyStats':         return getDailyStats(b);
    case 'getWeeklyStats':        return getWeeklyStats(b);
    case 'getUserAttendanceSummary': return getUserAttendanceSummary(b);
    case 'getAttendanceInsights': return getAttendanceInsights(b);

    case 'getOrgByGUID':          return getOrgByGUID(b.guid);
    case 'getRolesDepartments':   return getRolesDepartments(b.guid);
    case 'getDepartments':        return getLookup(SH.DEPARTMENTS);
    case 'getRoles':              return getLookup(SH.ROLES);
    case 'getAttendanceTypes':    return getLookup(SH.ATT_TYPE);
    case 'getLocations':          return getLookup(SH.ATT_LOCATIONS);
    case 'getAttendanceLocations': return getLookup(SH.ATT_LOCATIONS);
    case 'getUserLocMap':         return getUserLocMap(b);
    case 'getCategoryLocationMap': return getCategoryLocationMap(b);

    case 'addDepartment':         return addDepartment(b);
    case 'addAttendanceLocation': return addAttendanceLocation(b);
    case 'addUserLocMap':         return addCategoryLocationMap(b);
    case 'addCategoryLocationMap': return addCategoryLocationMap(b);

    case 'setupSheets':           return setupSheets();
    case 'resetUsersSheet':       return resetUsersSheet();
    case 'debug':                 return debugInfo();
    default: return { success: false, message: 'Unknown action: ' + b.action };
  }
}

// ============================================================
//  SHEET UTILITIES
// ============================================================

function ss() {
  var cfg = getTenantConfig(currentGuid());
  if (cfg && cfg.spreadsheetUrl) {
    var spreadsheetId = extractSpreadsheetId(cfg.spreadsheetUrl);
    if (spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

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
  try {
    var desired = HEADERS[name] || [];
    if (!desired.length || !sheet) return;
    var lastCol = Math.max(sheet.getLastColumn(), desired.length);
    if (sheet.getMaxColumns() < desired.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), desired.length - sheet.getMaxColumns());
    }
    var existing = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var merged = [];
    var seen = {};
    existing.forEach(function(h) {
      var key = String(h || '').trim();
      if (key && !seen[key]) {
        seen[key] = true;
        merged.push(key);
      }
    });
    desired.forEach(function(h) {
      var key = String(h || '').trim();
      if (key && !seen[key]) {
        seen[key] = true;
        merged.push(key);
      }
    });
    if (!merged.length) merged = desired.slice();
    if (merged.length > lastCol) {
      sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
      sheet.setFrozenRows(1);
    } else {
      var current = existing.map(function(h) { return String(h || '').trim(); });
      var changed = current.length !== merged.length;
      if (!changed) {
        for (var i = 0; i < merged.length; i++) {
          if (current[i] !== merged[i]) { changed = true; break; }
        }
      }
      if (changed) {
        sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
        sheet.setFrozenRows(1);
      }
    }
  } catch (e) {}
}

function getRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var data = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  var hdrs = data[0];
  var tenantId = currentGuid();
  return data.slice(1).map(function(row) {
    var o = {};
    hdrs.forEach(function(h, i) { o[h] = row[i]; });
    if (o.tenant_id && String(o.tenant_id).trim() && String(o.tenant_id).trim() !== tenantId) return null;
    return o;
  }).filter(function(row) { return !!row; });
}

function getCached(sheetName, ttl) {
  var cache = CacheService.getScriptCache();
  var hit   = cache.get(tenantScopedKey('rows_' + sheetName));
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  var rows  = getRows(getSheet(sheetName));
  try { cache.put(tenantScopedKey('rows_' + sheetName), JSON.stringify(rows), ttl || TTL_LOOKUP); } catch(e) {}
  return rows;
}

function getCachedData(key, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get(tenantScopedKey(String(key || '')));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function setCachedData(key, value, ttl) {
  try {
    CacheService.getScriptCache().put(
      tenantScopedKey(String(key || '')),
      JSON.stringify(value),
      Math.max(1, parseInt(ttl, 10) || TTL_LOOKUP)
    );
    return true;
  } catch (e) {
    return false;
  }
}

function invalidate(sheetName) {
  try { CacheService.getScriptCache().remove(tenantScopedKey('rows_' + sheetName)); } catch(e) {}
}

function invalidateCachedData(key) {
  try { CacheService.getScriptCache().remove(tenantScopedKey(String(key || ''))); } catch (e) {}
}

function buildMap(rows, key) {
  var m = {};
  rows.forEach(function(r) { m[r[key]] = r; });
  return m;
}

function normalizeGuid(guid) {
  var raw = String(guid || '').trim();
  return raw || DEFAULT_GUID;
}

function setRequestGuid(guid) {
  REQUEST_CONTEXT.guid = normalizeGuid(guid);
  return REQUEST_CONTEXT.guid;
}

function currentGuid() {
  return normalizeGuid(REQUEST_CONTEXT.guid);
}

function tenantIdOf(value) {
  return normalizeGuid(value || currentGuid());
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

function rowMatchesTenant(rowTenant, tenantId) {
  var rowValue = String(rowTenant || '').trim();
  if (!rowValue) return true;
  return rowValue === tenantIdOf(tenantId);
}

function findNearbyAttendanceLocations(lat, lng, tenantId, excludeLocationId) {
  var rows = getCached(SH.ATT_LOCATIONS, TTL_LOOKUP);
  var matches = [];
  rows.forEach(function(loc) {
    if (!rowMatchesTenant(loc.tenant_id, tenantId)) return;
    if (excludeLocationId && String(loc.attendance_location_id) === String(excludeLocationId)) return;
    var dLat = parseFloat(loc.latitude);
    var dLng = parseFloat(loc.longitude);
    if (isNaN(dLat) || isNaN(dLng)) return;
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

function findLocationById(locationId, tenantId) {
  var rows = getCached(SH.ATT_LOCATIONS, TTL_LOOKUP);
  var id = String(locationId || '').trim();
  for (var i = 0; i < rows.length; i++) {
    if (!rowMatchesTenant(rows[i].tenant_id, tenantId)) continue;
    if (String(rows[i].attendance_location_id) === id) return rows[i];
  }
  return null;
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
  if (!mapping && cat) mapping = findCategoryLocationMap(cat, '', tenantId);
  if (!mapping) return null;
  var location = findLocationById(mapping.attendance_location_id, tenantId);
  if (!location) return null;
  var allowedDistance = parseInt(mapping.allowed_distance || location.geofence_radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
  return {
    categoryId: cat,
    subcategoryId: sub,
    attendanceLocationId: location.attendance_location_id,
    allowedDistance: allowedDistance,
    anchorLat: parseFloat(location.latitude),
    anchorLng: parseFloat(location.longitude),
    location: location,
    mapping: mapping
  };
}

function tenantScopedKey(key) {
  return currentGuid() + '__' + key;
}

function getLiveStateInfo() {
  var state = getCachedData('live_state', 300) || {};
  return {
    version: String(state.version || '0'),
    updatedAt: String(state.updatedAt || '')
  };
}

function touchLiveState(reason) {
  try {
    var state = getLiveStateInfo();
    var nextVersion = String((parseInt(state.version, 10) || 0) + 1);
    setCachedData('live_state', {
      version: nextVersion,
      updatedAt: new Date().toISOString(),
      reason: String(reason || '')
    }, 300);
    return nextVersion;
  } catch (e) {
    return '0';
  }
}

function extractSpreadsheetId(url) {
  var match = String(url || '').match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : '';
}

function getTenantConfig(guid) {
  return TENANT_DIRECTORY[normalizeGuid(guid)] || null;
}

function getTenantGeofence(guid) {
  var tenantGuid = normalizeGuid(guid);
  var configured = TENANT_GEOFENCE[tenantGuid] || {};
  return {
    latitude: parseFloat(configured.latitude || DEFAULT_LAT),
    longitude: parseFloat(configured.longitude || DEFAULT_LNG),
    radius: Math.max(parseInt(configured.radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS, DEFAULT_RADIUS),
    label: String(configured.label || '')
  };
}

function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function getOrgByGUID(guid) {
  try {
    var tenant = getTenantConfig(guid);
    if (!tenant) return { success: false, message: 'Tenant not found for GUID: ' + normalizeGuid(guid) };

    setRequestGuid(tenant.guid);
    var roles = getCached(SH.ROLES, TTL_LOOKUP).map(function(r) {
      return String(r.name || r.role_id || '').trim();
    }).filter(function(name) { return !!name; });

    return {
      success: true,
      guid: tenant.guid,
      name: tenant.institution.name,
      logo: tenant.institution.logoUrl || '',
      gsheet_url: tenant.spreadsheetUrl,
      orgType: tenant.orgType || '',
      roles: roles,
      institution: tenant.institution,
      application: tenant.application,
      apiUrl: tenant.apiUrl || getWebAppUrl()
    };
  } catch (err) {
    return { success: false, message: 'getOrgByGUID: ' + err };
  }
}

function getRolesDepartments(guid) {
  try {
    setRequestGuid(guid);
    return {
      success: true,
      roles: getCached(SH.ROLES, TTL_LOOKUP),
      departments: getCached(SH.DEPARTMENTS, TTL_LOOKUP),
      attendanceLocations: getCached(SH.ATT_LOCATIONS, TTL_LOOKUP)
    };
  } catch (err) {
    return { success: false, message: 'getRolesDepartments: ' + err };
  }
}

// ── UserIndex ─────────────────────────────────────────────────

function getUserByEmail(email) {
  var lower = String(email).toLowerCase();
  var idx   = getSheet(SH.USER_INDEX);
  var data  = idx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === lower) {
      var user = fetchUserRow(data[i][2]);
      if (!user) continue;
      if (user.tenant_id && String(user.tenant_id).trim() && String(user.tenant_id).trim() !== currentGuid()) continue;
      return user;
    }
  }
  return null;
}

function getUserById(userId) {
  var idx  = getSheet(SH.USER_INDEX);
  var data = idx.getDataRange().getValues();
  var idStr = String(userId);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === idStr) {
      var user = fetchUserRow(data[i][2]);
      if (!user) continue;
      if (user.tenant_id && String(user.tenant_id).trim() && String(user.tenant_id).trim() !== currentGuid()) continue;
      return user;
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

function updateUserPasswordHash(userId, passwordHash) {
  var sheet = getSheet(SH.USERS);
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var uids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < uids.length; i++) {
    if (String(uids[i][0]) === String(userId)) {
      sheet.getRange(i + 2, 9).setValue(passwordHash);
      invalidate(SH.USERS);
      return true;
    }
  }
  return false;
}

// ── Helpers ───────────────────────────────────────────────────

function sha256Hex(value) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return raw.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function legacyHashPw(pw) {
  return sha256Hex(pw);
}

function hashPw(pw, salt) {
  var s = salt || ('salt_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5));
  var out = String(pw || '') + '|' + s;
  for (var i = 0; i < PASSWORD_ITERATIONS; i++) {
    out = sha256Hex(out + '|' + i);
  }
  return 'v2$' + PASSWORD_ITERATIONS + '$' + s + '$' + out;
}

function verifyPw(pw, storedHash) {
  var raw = String(storedHash || '');
  if (raw.indexOf('v2$') === 0) {
    var parts = raw.split('$');
    if (parts.length !== 4) return false;
    var iterations = parseInt(parts[1], 10);
    var salt = parts[2];
    var out = String(pw || '') + '|' + salt;
    for (var i = 0; i < iterations; i++) {
      out = sha256Hex(out + '|' + i);
    }
    return out === parts[3];
  }
  return legacyHashPw(pw) === raw;
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

function pad2(n) { return String(n).padStart(2, '0'); }

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
  var duration = AUTO_SESSION_WINDOW_MINUTES;
  var windowName = 'Attendance Window';
  if (typeof startHm === 'object' && startHm) {
    duration = parseInt(startHm.duration_minutes || AUTO_SESSION_WINDOW_MINUTES, 10);
    windowName = String(startHm.name || windowName);
    norm = normalizeClockTime(startHm.start_time || norm);
    startDt = new Date(dateStr + 'T' + norm + ':00');
  }
  var endDt = new Date(startDt.getTime() + duration * 60000);
  return {
    session_id: 'auto_' + dateStr + '_' + norm.replace(':', ''),
    teacher_id: 'ALL',
    teacher_name: 'Automatic',
    subject: windowName || AUTO_SESSION_SUBJECT,
    date: dateStr,
    start_time: norm + ':00',
    end_time: Utilities.formatDate(endDt, tz(), 'HH:mm:ss'),
    status: 'open',
    window_minutes: duration,
    is_auto: true
  };
}

function getAutoSessionsForDate(dateStr) {
  return getWindowsConfig().map(function(windowRow) {
    return buildAutoSession(dateStr, windowRow);
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
  ensureDefaultLookupRows();
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

function ensureDefaultLookupRows() {
  var roles = getCached(SH.ROLES, TTL_LOOKUP);
  if (!roles || !roles.length) {
    ensureExactRows('Roles', [
      [1, 'admin'],
      [2, 'teacher'],
      [3, 'student'],
      [4, 'employee']
    ]);
  }
}

function normalizeRoleValue(roleValue) {
  var raw = String(roleValue || '').trim().toLowerCase();
  if (!raw) return '';
  var roles = getRoleDirectory();
  // If it's already a name (student/teacher/admin) return as-is
  if (roles.byName[raw]) return raw;
  // If it's a numeric id, resolve to name
  return roles.byId[raw] || raw;
}

function findRoleIdByName(roleName) {
  var raw = String(roleName || '').trim().toLowerCase();
  if (!raw) return '';
  return getRoleDirectory().byName[raw] || '';
}

function resolveRoleId(roleValue) {
  var raw = String(roleValue || '').trim();
  if (!raw) return '';
  var roles = getRoleDirectory();
  if (roles.byId[raw]) return raw;
  return roles.byName[raw.toLowerCase()] || '';
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

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function toCleanText(value, maxLen) {
  var s = String(value || '').replace(/\s+/g, ' ').trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

function parseNumber(value) {
  if (value === '' || value === null || typeof value === 'undefined') return null;
  var n = parseFloat(value);
  return isNaN(n) ? null : n;
}

function isValidMobile(mobile) {
  var raw = String(mobile || '').trim();
  if (!raw) return true;
  return /^[0-9+\-\s]{7,20}$/.test(raw);
}

function isValidCoordinate(lat, lng) {
  return lat !== null && lng !== null &&
    lat >= -90 && lat <= 90 &&
    lng >= -180 && lng <= 180;
}

function hasDepartment(departmentId) {
  var raw = String(departmentId || '').trim();
  if (!raw) return false;
  var rows = getCached(SH.DEPARTMENTS, TTL_LOOKUP);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].department_id) === raw) return true;
  }
  return false;
}

function hasRole(roleId) {
  return !!resolveRoleId(roleId);
}

function getWindowsConfig() {
  var rows = getCached(SH.ATT_WINDOWS, TTL_LOOKUP).filter(function(r) {
    return String(r.status || '').toLowerCase() === 'active';
  });
  if (!rows.length) {
    return [
      { name: 'Morning 1',   start_time: '09:00', duration_minutes: AUTO_SESSION_WINDOW_MINUTES },
      { name: 'Morning 2',   start_time: '10:30', duration_minutes: AUTO_SESSION_WINDOW_MINUTES },
      { name: 'Afternoon 1', start_time: '14:00', duration_minutes: AUTO_SESSION_WINDOW_MINUTES },
      { name: 'Afternoon 2', start_time: '15:30', duration_minutes: AUTO_SESSION_WINDOW_MINUTES }
    ];
  }
  return rows;
}

// ── Auth sessions ─────────────────────────────────────────────
// auth_token stays as a random string (it's a token, not a pk used in joins)

function authSessionHeaderIndex(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === name) return i;
  }
  return -1;
}

function rowToObject(row, headers) {
  var obj = {};
  headers.forEach(function(h, i) {
    obj[String(h).trim()] = row[i];
  });
  return obj;
}

function getAuthSecret() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('AUTH_HMAC_SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('AUTH_HMAC_SECRET', secret);
  }
  return secret;
}

function base64UrlEncode(value) {
  var bytes = value;
  if (typeof value === 'string') bytes = Utilities.newBlob(value, 'application/json').getBytes();
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function base64UrlDecodeToString(value) {
  var padded = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  while (padded.length % 4) padded += '=';
  return Utilities.newBlob(Utilities.base64Decode(padded)).getDataAsString();
}

function base64UrlDecodeToBytes(value) {
  var padded = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  while (padded.length % 4) padded += '=';
  return Utilities.base64Decode(padded);
}

function safeJsonParse(raw) {
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function getRequestIpHint() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim();
  } catch (e) {
    return '';
  }
}

function makeRateLimitKey(action, body) {
  var who = String((body && (body.email || body.userId || body.user_id || body.deviceId || body.guid)) || '').trim();
  if (!who) who = getRequestIpHint() || 'anon';
  return currentGuid() + ':' + String(action || 'unknown') + ':' + who.toLowerCase();
}

function getRateLimitSheet() {
  return getSheet(SH.RATE_LIMIT);
}

function checkRateLimit(key) {
  try {
    var cache = CacheService.getScriptCache();
    var now = new Date();
    var windowStart = new Date(Math.floor(now.getTime() / 60000) * 60000);
    var cacheKey = tenantScopedKey('rl_' + key + '_' + windowStart.getTime());
    var current = parseInt(cache.get(cacheKey) || '0', 10);
    if (current >= RATE_LIMIT_MAX_PER_MINUTE) {
      appendAuditLog('rate_limit_block', '', '', '', { key: key, count: current, windowStart: windowStart.toISOString() });
      return { ok: false, error: { success: false, message: 'Too many requests. Please try again later.' } };
    }
    cache.put(cacheKey, String(current + 1), 120);
    return { ok: true };
  } catch (e) {
    return { ok: true };
  }
}

function incrementRateLimit(key) {
  try {
    var sheet = getRateLimitSheet();
    var now = new Date();
    var rows = getRows(sheet);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].key) === String(key)) {
        sheet.getRange(i + 2, 2).setValue(parseInt(rows[i].count || 0, 10) + 1);
        sheet.getRange(i + 2, 3).setValue(now.toISOString());
        return true;
      }
    }
    sheet.appendRow([key, 1, now.toISOString(), now.toISOString()]);
    return true;
  } catch (e) {
    return false;
  }
}

function tokenToBasePayload(payload) {
  return {
    user_id: String(payload.user_id || ''),
    guid: String(payload.guid || currentGuid()),
    role: String(payload.role || ''),
    exp: parseInt(payload.exp, 10) || 0,
    typ: String(payload.typ || 'access')
  };
}

function signTokenPayload(payload) {
  var header = { alg: 'HS256', typ: 'JWT' };
  var encodedHeader = base64UrlEncode(JSON.stringify(header));
  var encodedPayload = base64UrlEncode(JSON.stringify(payload));
  var signingInput = encodedHeader + '.' + encodedPayload;
  var secretBytes = Utilities.newBlob(getAuthSecret()).getBytes();
  var signatureBytes = Utilities.computeHmacSha256Signature(signingInput, secretBytes);
  var encodedSignature = base64UrlEncode(signatureBytes);
  return signingInput + '.' + encodedSignature;
}

function verifyToken(token) {
  try {
    var parts = String(token || '').split('.');
    if (parts.length !== 3) return { ok: false, code: 'INVALID_TOKEN' };
    var signingInput = parts[0] + '.' + parts[1];
    var expected = base64UrlEncode(Utilities.computeHmacSha256Signature(signingInput, Utilities.newBlob(getAuthSecret()).getBytes()));
    if (expected !== parts[2]) return { ok: false, code: 'SIGNATURE_MISMATCH' };
    var payload = safeJsonParse(base64UrlDecodeToString(parts[1]));
    if (!payload) return { ok: false, code: 'INVALID_PAYLOAD' };
    var nowSeconds = Math.floor(Date.now() / 1000);
    if (!payload.exp || nowSeconds >= parseInt(payload.exp, 10)) return { ok: false, code: 'TOKEN_EXPIRED', payload: payload };
    return { ok: true, payload: tokenToBasePayload(payload) };
  } catch (e) {
    return { ok: false, code: 'INVALID_TOKEN' };
  }
}

function generateToken(user, ttlMinutes, tokenType, guid) {
  var exp = Math.floor((Date.now() + (ttlMinutes * 60000)) / 1000);
  var payload = tokenToBasePayload({
    user_id: user.user_id,
    guid: normalizeGuid(guid || currentGuid()),
    role: normalizeRoleValue(user.role_name || user.roleId || user.role_id || ''),
    exp: exp,
    typ: tokenType || 'access'
  });
  return signTokenPayload(payload);
}

function generateRefreshToken(user, guid) {
  return generateToken(user, REFRESH_TOKEN_TTL_DAYS * 24 * 60, 'refresh', guid);
}

function generateAccessToken(user, guid) {
  return generateToken(user, ACCESS_TOKEN_TTL_MINUTES, 'access', guid);
}

function ensureAuthSessionSchema() {
  var sheet = getSheet(SH.AUTH_SESSIONS);
  var desired = HEADERS.AuthSessions.slice();
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), desired.length);
  var currentHeaders = lastRow >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || '').trim(); })
    : [];
  var same = desired.length === currentHeaders.length;
  if (same) {
    for (var i = 0; i < desired.length; i++) {
      if (currentHeaders[i] !== desired[i]) { same = false; break; }
    }
  }
  if (same) return sheet;

  var existing = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues()
    : [];
  var migrated = existing.map(function(row) {
    var obj = rowToObject(row, currentHeaders.length ? currentHeaders : desired);
    if (!obj.session_id && obj.token) obj.session_id = String(obj.token);
    if (!obj.session_id && row[0]) obj.session_id = row[0];
    if (!obj.token && row[0]) obj.token = row[0];
    if (!obj.refresh_token) obj.refresh_token = '';
    if (!obj.device_id) obj.device_id = '';
    if (!obj.created_at && row[3]) obj.created_at = row[3];
    if (!obj.expires_at && row[4]) obj.expires_at = row[4];
    if (!obj.refresh_expires_at) obj.refresh_expires_at = obj.expires_at || row[4] || '';
    if (!obj.status && row[5]) obj.status = row[5];
    if (!obj.guid) obj.guid = currentGuid();
    if (!obj.role_name && row[2]) obj.role_name = row[2];
    if (!obj.last_seen_at) obj.last_seen_at = obj.created_at || '';
    return desired.map(function(h) {
      return typeof obj[h] === 'undefined' ? '' : obj[h];
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, desired.length).setValues([desired]);
  if (migrated.length) {
    sheet.getRange(2, 1, migrated.length, desired.length).setValues(migrated);
  }
  sheet.setFrozenRows(1);
  invalidate(SH.AUTH_SESSIONS);
  return sheet;
}

function getAuthSessionSheet() {
  return ensureAuthSessionSchema();
}

function createAuthSession(userId, roleName, deviceId, guid) {
  var sessionId = 'sess_' + Utilities.getUuid().replace(/-/g, '');
  var createdAt = new Date();
  var accessUser = {
    user_id: userId,
    role_name: roleName,
    roleId: roleName,
    role_id: roleName
  };
  var token = generateAccessToken(accessUser, guid);
  var refreshToken = generateRefreshToken(accessUser, guid);
  var expiresAt = new Date(Date.now() + ACCESS_TOKEN_TTL_MINUTES * 60000);
  var refreshExpiresAt = new Date(Date.now() + REFRESH_TOKEN_TTL_DAYS * 24 * 60 * 60000);
  getAuthSessionSheet().appendRow([
    sessionId,
    String(userId || ''),
    token,
    refreshToken,
    String(deviceId || ''),
    createdAt.toISOString(),
    expiresAt.toISOString(),
    refreshExpiresAt.toISOString(),
    'ACTIVE',
    normalizeGuid(guid || currentGuid()),
    String(roleName || ''),
    createdAt.toISOString(),
    normalizeGuid(guid || currentGuid())
  ]);
  return {
    authToken: token,
    refreshToken: refreshToken,
    sessionId: sessionId,
    expiresAt: expiresAt.toISOString(),
    refreshExpiresAt: refreshExpiresAt.toISOString()
  };
}

function findAuthSessionRow(token) {
  var sheet = getAuthSessionSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var tokenIdx = authSessionHeaderIndex(headers, 'token');
  if (tokenIdx < 0) tokenIdx = authSessionHeaderIndex(headers, 'auth_token');
  if (tokenIdx < 0) return null;
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][tokenIdx]) !== String(token)) continue;
    return { sheet: sheet, headers: headers, row: rows[i], rowIndex: i + 2 };
  }
  return null;
}

function findRefreshSessionRow(refreshToken) {
  var sheet = getAuthSessionSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var refreshIdx = authSessionHeaderIndex(headers, 'refresh_token');
  if (refreshIdx < 0) return null;
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][refreshIdx]) !== String(refreshToken)) continue;
    return { sheet: sheet, headers: headers, row: rows[i], rowIndex: i + 2 };
  }
  return null;
}

function getAuthSession(token) {
  if (!token) return null;
  var found = findAuthSessionRow(token);
  if (!found) return null;
  var session = rowToObject(found.row, found.headers);
  var status = String(session.status || '').trim().toUpperCase();
  var payloadCheck = verifyToken(token);
  var exp = new Date(session.expires_at || session.expiresAt || '');
  if (status !== 'ACTIVE' || isNaN(exp.getTime()) || exp.getTime() < Date.now() || !payloadCheck.ok) {
    try {
      var statusIdx = authSessionHeaderIndex(found.headers, 'status');
      if (statusIdx >= 0) found.sheet.getRange(found.rowIndex, statusIdx + 1).setValue('EXPIRED');
      invalidate(SH.AUTH_SESSIONS);
    } catch (e) {}
    return null;
  }
  return session;
}

function touchAuthSession(token) {
  try {
    if (!token) return;
    var cache = CacheService.getScriptCache();
    var touchKey = tenantScopedKey('auth_touch_' + token);
    var lastTouch = parseInt(cache.get(touchKey) || '0', 10);
    if (lastTouch && (Date.now() - lastTouch) < 300000) return;
    var found = findAuthSessionRow(token);
    if (!found) return;
    var lastSeenIdx = authSessionHeaderIndex(found.headers, 'last_seen_at');
    if (lastSeenIdx >= 0) {
      found.sheet.getRange(found.rowIndex, lastSeenIdx + 1).setValue(new Date().toISOString());
      cache.put(touchKey, String(Date.now()), 300);
      invalidate(SH.AUTH_SESSIONS);
    }
  } catch (e) {}
}

function expireAuthSession(token) {
  try {
    var found = findAuthSessionRow(token);
    if (!found) return false;
    var statusIdx = authSessionHeaderIndex(found.headers, 'status');
    if (statusIdx >= 0) found.sheet.getRange(found.rowIndex, statusIdx + 1).setValue('EXPIRED');
    var cache = CacheService.getScriptCache();
    cache.remove(tenantScopedKey('auth_touch_' + token));
    invalidate(SH.AUTH_SESSIONS);
    return true;
  } catch (e) {
    return false;
  }
}

function validateToken(b) {
  var token = String((b && (b.authToken || b.token)) || '').trim();
  if (!token) return { ok: false, error: { success: false, message: 'Unauthorized' } };
  var signed = verifyToken(token);
  if (!signed.ok) {
    return { ok: false, error: { success: false, code: signed.code || 'TOKEN_INVALID', message: 'Unauthorized' } };
  }
  var session = getAuthSession(token);
  if (!session) return { ok: false, error: { success: false, message: 'Unauthorized' } };
  var user = getUserById(session.user_id);
  if (!user) return { ok: false, error: { success: false, message: 'Unauthorized' } };
  var requestGuid = normalizeGuid((b && b.guid) || currentGuid());
  if (session.guid && normalizeGuid(session.guid) !== requestGuid) {
    return { ok: false, error: { success: false, message: 'Unauthorized' } };
  }
  touchAuthSession(token);
  return { ok: true, session: session, user: user, tokenPayload: signed.payload };
}

function authorize(user, action, body) {
  var roleName = normalizeRoleValue(user && (user.role_name || user.roleId || user.role_id || ''));
  if (roleName === 'admin') return { ok: true };

  var actionName = String(action || '').trim();
  var bodyUserId = String((body && (body.userId || body.user_id)) || '').trim();

  var selfOnlyActions = {
    markEntry: true,
    markAttendance: true,
    markExit: true,
    exitAttendance: true,
    trackLocation: true,
    trackStudentLocation: true,
    getMyAttendance: true,
    registerDevice: true,
    checkDevice: true
  };
  if (selfOnlyActions[actionName]) {
    if (!bodyUserId || String(user.user_id) !== bodyUserId) {
      return { ok: false, error: { success: false, message: 'Unauthorized' } };
    }
    return { ok: true };
  }

  var teacherActions = {
    createSession: true,
    closeSession: true,
    getActiveSession: true,
    getSessions: true,
    getTeacherNotifications: true,
    getDashboard: true,
    getStudents: true,
    exportAttendance: true,
    getLiveAttendance: true,
    getLiveUpdates: true,
    getUserLocations: true,
    forceExitUser: true,
    getDailyStats: true,
    getWeeklyStats: true,
    getUserAttendanceSummary: true,
    getAttendanceInsights: true
  };
  if (teacherActions[actionName]) {
    return roleName === 'teacher' || roleName === 'admin'
      ? { ok: true }
      : { ok: false, error: { success: false, message: 'Unauthorized' } };
  }

  var adminActions = {
    addDepartment: true,
    addAttendanceLocation: true,
    addUserLocMap: true,
    addCategoryLocationMap: true,
    getCategoryLocationMap: true,
    setupSheets: true,
    resetUsersSheet: true,
    debug: true
  };
  if (adminActions[actionName]) {
    return roleName === 'admin'
      ? { ok: true }
      : { ok: false, error: { success: false, message: 'Unauthorized' } };
  }

  return { ok: true };
}

function requireAuth(b, allowedRoles) {
  var result = validateToken(b);
  if (!result.ok) return { ok: false, error: result.error };
  if (allowedRoles && allowedRoles.length && allowedRoles.indexOf(normalizeRoleValue(result.session.role_name)) < 0) {
    return { ok: false, error: { success: false, message: 'Unauthorized' } };
  }
  return { ok: true, session: result.session, user: result.user };
}

function appendAuditLog(eventType, actorUserId, actorRole, targetUserId, details) {
  try {
    var auditId = nextId(SH.AUDIT_LOG);
    var metadata = details || {};
    if (typeof metadata === 'object' && metadata && !metadata.guid) {
      metadata.guid = currentGuid();
    }
    getSheet(SH.AUDIT_LOG).appendRow([
      auditId,
      String(eventType || ''),
      String(actorUserId || ''),
      String(actorRole || ''),
      String(targetUserId || ''),
      JSON.stringify(metadata || {}),
      '',
      new Date().toISOString(),
      currentGuid()
    ]);
  } catch(e) {}
}

function logoutUser(b) {
  try {
    var auth = validateToken(b);
    if (!auth.ok) return auth.error;
    expireAuthSession(b.authToken || b.token);
    appendAuditLog('logout', auth.session.user_id, normalizeRoleValue(auth.session.role_name), auth.session.user_id, {
      deviceId: String(auth.session.device_id || ''),
      sessionId: String(auth.session.session_id || '')
    });
    return { success: true, message: 'Logged out successfully' };
  } catch (err) {
    return { success: false, message: 'logout: ' + err };
  }
}

function webauthnChallengeSheet() {
  return getSheet(SH.WEBAUTHN);
}

function webauthnCredentialSheet() {
  return getSheet(SH.WEBAUTHN_CRED);
}

function createWebAuthnChallenge(userId, action, metadata) {
  var challengeId = 'wch_' + Utilities.getUuid().replace(/-/g, '');
  var challengeBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, challengeId + '|' + Date.now() + '|' + Math.random(), Utilities.Charset.UTF_8);
  var challenge = base64UrlEncode(challengeBytes);
  var createdAt = new Date();
  var expiresAt = new Date(Date.now() + 5 * 60000);
  webauthnChallengeSheet().appendRow([
    challengeId,
    String(userId || ''),
    String(action || ''),
    challenge,
    createdAt.toISOString(),
    expiresAt.toISOString(),
    'ACTIVE',
    JSON.stringify(metadata || {})
  ]);
  return {
    challengeId: challengeId,
    challenge: challenge,
    expiresAt: expiresAt.toISOString()
  };
}

function verifyWebAuthnClientData(clientDataJSON, expectedChallenge, expectedType) {
  try {
    if (!clientDataJSON) return { ok: false };
    var raw = base64UrlDecodeToString(clientDataJSON);
    var parsed = safeJsonParse(raw);
    if (!parsed) return { ok: false };
    if (String(parsed.type || '') !== String(expectedType || '')) return { ok: false };
    if (String(parsed.challenge || '') !== String(expectedChallenge || '')) return { ok: false };
    return { ok: true, parsed: parsed };
  } catch (e) {
    return { ok: false };
  }
}

function findWebAuthnChallenge(challengeId) {
  var sheet = webauthnChallengeSheet();
  var rows = getRows(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i].challenge_id) === String(challengeId)) {
      return { row: rows[i], rowIndex: i + 2, sheet: sheet };
    }
  }
  return null;
}

function verifyWebAuthnChallenge(challengeId, expectedAction, userId, providedChallenge, credentialId) {
  var found = findWebAuthnChallenge(challengeId);
  if (!found) return { ok: false, message: 'Unauthorized' };
  var row = found.row;
  var exp = new Date(row.expires_at || '');
  if (String(row.status || '').toUpperCase() !== 'ACTIVE' || isNaN(exp.getTime()) || exp.getTime() < Date.now()) {
    return { ok: false, message: 'Unauthorized' };
  }
  if (String(row.action || '') !== String(expectedAction || '')) return { ok: false, message: 'Unauthorized' };
  if (String(row.user_id || '') && String(row.user_id || '') !== String(userId || '')) return { ok: false, message: 'Unauthorized' };
  if (providedChallenge && String(row.challenge || '') !== String(providedChallenge || '')) return { ok: false, message: 'Unauthorized' };
  if (credentialId && String(row.metadata || '').indexOf(String(credentialId)) < 0) {
    // metadata is informational; credential binding is validated elsewhere
  }
  try {
    found.sheet.getRange(found.rowIndex, 7).setValue('USED');
  } catch (e) {}
  return { ok: true, row: row };
}

function beginWebAuthnRegistration(b) {
  try {
    var email = String(b.email || '').trim().toLowerCase();
    var name = String(b.name || '').trim();
    if (!email || !name) return { success: false, message: 'email and name required' };
    var challenge = createWebAuthnChallenge(email, 'register', { name: name, guid: currentGuid() });
    appendAuditLog('webauthn_register_begin', '', '', '', { email: email, challengeId: challenge.challengeId });
    return {
      success: true,
      challengeId: challenge.challengeId,
      challenge: challenge.challenge,
      expiresAt: challenge.expiresAt
    };
  } catch (err) {
    return { success: false, message: 'beginWebAuthnRegistration: ' + err };
  }
}

function finishWebAuthnRegistration(b) {
  try {
    var email = String(b.email || '').trim().toLowerCase();
    var name = String(b.name || '').trim();
    var challengeId = String(b.challengeId || '').trim();
    var credentialId = String(b.credentialId || '').trim();
    var publicKey = String(b.publicKey || '').trim();
    var clientChallenge = String(b.challenge || '').trim();
    if (!email || !name || !challengeId || !credentialId) {
      return { success: false, message: 'Missing registration data' };
    }
    var challengeCheck = verifyWebAuthnChallenge(challengeId, 'register', email, clientChallenge, credentialId);
    if (!challengeCheck.ok) return { success: false, message: 'Unauthorized' };
    var clientCheck = verifyWebAuthnClientData(b.clientDataJSON, clientChallenge, 'webauthn.create');
    if (!clientCheck.ok) return { success: false, message: 'Unauthorized' };

    var user = getUserByEmail(email);
    if (!user) return { success: false, message: 'User not found' };
    var sheet = webauthnCredentialSheet();
    sheet.appendRow([
      credentialId,
      String(user.user_id || ''),
      publicKey,
      new Date().toISOString(),
      'ACTIVE',
      currentGuid()
    ]);
    var usersSheet = getSheet(SH.USERS);
    var last = usersSheet.getLastRow();
    for (var i = 2; i <= last; i++) {
      if (String(usersSheet.getRange(i, 8).getValue()).toLowerCase() === email) {
        usersSheet.getRange(i, 10).setValue(credentialId);
        break;
      }
    }
    invalidate(SH.USERS);
    appendAuditLog('webauthn_register_complete', user.user_id, normalizeRoleValue(user.role_id), user.user_id, {
      credentialId: credentialId
    });
    return { success: true, credentialId: credentialId, userId: user.user_id };
  } catch (err) {
    return { success: false, message: 'finishWebAuthnRegistration: ' + err };
  }
}

function beginWebAuthnLogin(b) {
  try {
    var email = String(b.email || '').trim().toLowerCase();
    if (!email) return { success: false, message: 'email required' };
    var user = getUserByEmail(email);
    if (!user) return { success: false, message: 'No account found for this email' };
    var credRows = getRows(webauthnCredentialSheet()).filter(function(r) {
      return String(r.user_id) === String(user.user_id) && String(r.status || '').toUpperCase() === 'ACTIVE';
    });
    if (!credRows.length && !user.biometric_code) {
      return { success: false, message: 'No biometric registered for this account' };
    }
    var challenge = createWebAuthnChallenge(user.user_id, 'login', {
      email: email,
      credentialId: String(user.biometric_code || (credRows[0] && credRows[0].credential_id) || ''),
      guid: currentGuid()
    });
    appendAuditLog('webauthn_login_begin', user.user_id, normalizeRoleValue(user.role_id), user.user_id, {
      challengeId: challenge.challengeId
    });
    return {
      success: true,
      challengeId: challenge.challengeId,
      challenge: challenge.challenge,
      credentialId: String(user.biometric_code || (credRows[0] && credRows[0].credential_id) || ''),
      userId: user.user_id,
      name: user.full_name,
      roleId: normalizeRoleValue(user.role_id),
      roleKey: normalizeRoleValue(user.role_id)
    };
  } catch (err) {
    return { success: false, message: 'beginWebAuthnLogin: ' + err };
  }
}

function finishWebAuthnLogin(b) {
  try {
    var email = String(b.email || '').trim().toLowerCase();
    var challengeId = String(b.challengeId || '').trim();
    var providedChallenge = String(b.challenge || '').trim();
    var credentialId = String(b.credentialId || '').trim();
    var publicKey = String(b.publicKey || '').trim();
    if (!email || !challengeId || !credentialId) {
      return { success: false, message: 'Missing login data' };
    }
    var user = getUserByEmail(email);
    if (!user) return { success: false, message: 'No account found for this email' };
    var challengeCheck = verifyWebAuthnChallenge(challengeId, 'login', user.user_id, providedChallenge, credentialId);
    if (!challengeCheck.ok) return { success: false, message: 'Unauthorized' };
    var clientCheck = verifyWebAuthnClientData(b.clientDataJSON, providedChallenge, 'webauthn.get');
    if (!clientCheck.ok) return { success: false, message: 'Unauthorized' };
    var credRows = getRows(webauthnCredentialSheet());
    var matched = false;
    for (var i = 0; i < credRows.length; i++) {
      if (String(credRows[i].credential_id) === credentialId &&
          String(credRows[i].user_id) === String(user.user_id) &&
          String(credRows[i].status || '').toUpperCase() === 'ACTIVE') {
        matched = true;
        if (publicKey && !String(credRows[i].public_key || '')) {
          try {
            webauthnCredentialSheet().getRange(i + 2, 3).setValue(publicKey);
          } catch (e) {}
        }
        break;
      }
    }
    if (!matched && String(user.biometric_code || '') !== credentialId) {
      return { success: false, message: 'Unauthorized' };
    }

    var deviceId = String(b.deviceId || b.device_id || '').trim();
    var roleName = normalizeRoleValue(user.role_id);
    var auth = createAuthSession(user.user_id, roleName || String(user.role_id || ''), deviceId, currentGuid());
    appendAuditLog('webauthn_login_complete', user.user_id, roleName, user.user_id, {
      credentialId: credentialId,
      deviceId: deviceId
    });
    return {
      success: true,
      userId: user.user_id,
      name: user.full_name,
      roleId: roleName,
      roleKey: roleName,
      roleDbId: String(user.role_id || ''),
      deptId: user.department_id,
      subcategoryId: user.subcategory_id || '',
      authToken: auth.authToken,
      refreshToken: auth.refreshToken,
      authExpiresAt: auth.expiresAt,
      refreshExpiresAt: auth.refreshExpiresAt,
      sessionId: auth.sessionId,
      deviceId: deviceId || String(user.device_identification || ''),
      guid: currentGuid()
    };
  } catch (err) {
    return { success: false, message: 'finishWebAuthnLogin: ' + err };
  }
}

function refreshAuthSession(b) {
  try {
    var refreshToken = String((b && (b.refreshToken || b.refresh_token)) || '').trim();
    if (!refreshToken) return { success: false, message: 'Unauthorized' };
    var found = findRefreshSessionRow(refreshToken);
    if (!found) return { success: false, message: 'Unauthorized' };
    var session = rowToObject(found.row, found.headers);
    var status = String(session.status || '').trim().toUpperCase();
    var refreshExp = new Date(session.refresh_expires_at || session.expires_at || '');
    if (status !== 'ACTIVE' || isNaN(refreshExp.getTime()) || refreshExp.getTime() < Date.now()) {
      var sIdx = authSessionHeaderIndex(found.headers, 'status');
      if (sIdx >= 0) found.sheet.getRange(found.rowIndex, sIdx + 1).setValue('EXPIRED');
      invalidate(SH.AUTH_SESSIONS);
      return { success: false, message: 'Unauthorized' };
    }

    var user = getUserById(session.user_id);
    if (!user) return { success: false, message: 'Unauthorized' };
    var roleName = normalizeRoleValue(user.role_id);
    var newAccess = generateAccessToken({ user_id: user.user_id, role_name: roleName }, session.guid);
    var newRefresh = generateRefreshToken({ user_id: user.user_id, role_name: roleName }, session.guid);
    var expiresAt = new Date(Date.now() + ACCESS_TOKEN_TTL_MINUTES * 60000).toISOString();
    var refreshExpiresAt = new Date(Date.now() + REFRESH_TOKEN_TTL_DAYS * 24 * 60 * 60000).toISOString();

    var tokenIdx = authSessionHeaderIndex(found.headers, 'token');
    var refreshIdx = authSessionHeaderIndex(found.headers, 'refresh_token');
    var expiresIdx = authSessionHeaderIndex(found.headers, 'expires_at');
    var refreshExpIdx = authSessionHeaderIndex(found.headers, 'refresh_expires_at');
    var lastSeenIdx = authSessionHeaderIndex(found.headers, 'last_seen_at');
    if (tokenIdx >= 0) found.sheet.getRange(found.rowIndex, tokenIdx + 1).setValue(newAccess);
    if (refreshIdx >= 0) found.sheet.getRange(found.rowIndex, refreshIdx + 1).setValue(newRefresh);
    if (expiresIdx >= 0) found.sheet.getRange(found.rowIndex, expiresIdx + 1).setValue(expiresAt);
    if (refreshExpIdx >= 0) found.sheet.getRange(found.rowIndex, refreshExpIdx + 1).setValue(refreshExpiresAt);
    if (lastSeenIdx >= 0) found.sheet.getRange(found.rowIndex, lastSeenIdx + 1).setValue(new Date().toISOString());
    invalidate(SH.AUTH_SESSIONS);
    appendAuditLog('refresh_token', user.user_id, roleName, user.user_id, {
      sessionId: String(session.session_id || ''),
      deviceId: String(session.device_id || '')
    });
    return {
      success: true,
      authToken: newAccess,
      refreshToken: newRefresh,
      authExpiresAt: expiresAt,
      refreshExpiresAt: refreshExpiresAt,
      userId: user.user_id,
      name: user.full_name,
      roleId: roleName,
      roleKey: roleName,
      roleDbId: String(user.role_id || ''),
      deptId: user.department_id,
      subcategoryId: user.subcategory_id || '',
      sessionId: String(session.session_id || ''),
      deviceId: String(session.device_id || ''),
      guid: String(session.guid || currentGuid())
    };
  } catch (err) {
    return { success: false, message: 'refreshToken: ' + err };
  }
}

function pushTeacherNotification(notification) {
  try {
    if (!notification || !notification.sessionId) return;
    var cache = CacheService.getScriptCache();
    var key = tenantScopedKey('teacher_notif_' + notification.sessionId);
    var rows = [];
    var cached = cache.get(key);
    if (cached) { try { rows = JSON.parse(cached) || []; } catch(e) {} }
    rows.unshift(notification);
    rows = rows.slice(0, 25);
    cache.put(key, JSON.stringify(rows), 21600);
  } catch(e) {}
}

function getOpenSessionForToday() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(tenantScopedKey('active_session'));
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

function resolveUserTrackingLocation(userId) {
  var geofence = getTenantGeofence(currentGuid());
  var user = getUserById(userId);
  if (!user) {
    return {
      attendanceLocationId: 1,
      allowedDistance: geofence.radius,
      anchorLat: geofence.latitude,
      anchorLng: geofence.longitude,
      location: null,
      mapping: null
    };
  }
  var mapping = resolveMappedAttendanceLocation(user, currentGuid());
  if (!mapping) {
    return {
      attendanceLocationId: 1,
      allowedDistance: geofence.radius,
      anchorLat: geofence.latitude,
      anchorLng: geofence.longitude,
      location: null,
      mapping: null
    };
  }
  return mapping;
}

function getAllowedDistance(userId) {
  return resolveUserTrackingLocation(userId).allowedDistance;
}

function getAttendanceCenter(userId) {
  var track = resolveUserTrackingLocation(userId);
  return { lat: track.anchorLat, lng: track.anchorLng, attendanceLocationId: track.attendanceLocationId };
}

function saveLocation(userId, lat, lng, distance) {
  var now = new Date();
  var lmId = nextId(SH.LOC_MONITOR);   // â† integer
  getSheet(SH.LOC_MONITOR).appendRow([lmId, userId, lat, lng, distance, now.toISOString(), currentGuid()]);
  touchLiveState('location_write');
  return lmId;
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  return haversine(lat1, lon1, lat2, lon2);
}

function getAttendanceWindowRows() {
  return getRows(getSheet(SH.ATT_WINDOWS));
}

function parseWindowEndTime(windowRow) {
  var explicitEnd = String(windowRow.end_time || windowRow.endTime || '').trim();
  if (explicitEnd) return explicitEnd;
  var start = String(windowRow.start_time || '').trim();
  var durationMins = parseInt(windowRow.duration_minutes || windowRow.durationMinutes || 0, 10);
  if (!start || isNaN(durationMins)) return '';
  var parts = start.split(':');
  if (parts.length !== 2) return '';
  var h = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) return '';
  var dt = new Date();
  dt.setHours(h, m + durationMins, 0, 0);
  return Utilities.formatDate(dt, tz(), 'HH:mm:ss');
}

function validateSession(userId, locationId) {
  try {
    var now = new Date();
    var today = Utilities.formatDate(now, tz(), 'yyyy-MM-dd');
    var windows = getAttendanceWindowRows().filter(function(w) {
      return String(w.status || '').toLowerCase() === 'active';
    });
    if (!windows.length) {
      return { success: false, code: 'NO_WINDOW', message: 'No attendance window configured' };
    }

    var locId = String(locationId || '').trim();
    var scopedWindows = windows.filter(function(w) {
      var windowLoc = String(w.location_id || '').trim();
      return !windowLoc || !locId || windowLoc === locId;
    });
    if (!scopedWindows.length) scopedWindows = windows;
    var matched = null;
    for (var i = 0; i < scopedWindows.length; i++) {
      var w = scopedWindows[i];
      var startTime = String(w.start_time || '').trim();
      var endTime = parseWindowEndTime(w);
      if (!startTime || !endTime) continue;
      var startDt = new Date(today + 'T' + startTime);
      var endDt = new Date(today + 'T' + endTime);
      if (isNaN(startDt.getTime()) || isNaN(endDt.getTime())) continue;
      if (now >= startDt && now <= endDt) {
        matched = {
          sessionId: String(w.session_id || w.window_id || w.name || ''),
          startTime: startTime,
          endTime: endTime,
          locationId: windowLoc || locId || ''
        };
        break;
      }
    }

    if (!matched) {
      return { success: false, code: 'SESSION_INVALID', message: 'Attendance is outside the allowed session window' };
    }
    return { success: true, window: matched };
  } catch (err) {
    return { success: false, code: 'SESSION_ERROR', message: 'validateSession: ' + err };
  }
}

function attendanceReadColumnCount(sheet) {
  var count = sheet.getLastColumn();
  if (count < 13) count = 13;
  if (count > 17) count = 17;
  return count;
}

// ============================================================
//  1. REGISTER USER
//  user_id → nextId(Users)          — integer
//  UserAttendanceLocationMap id → nextId(UserAttendanceLocationMap) — integer
// ============================================================

function registerUser(b) {
  try {
    if (!b.name || !b.email || !b.password || !b.roleId || !b.departmentId || !b.subcategoryId)
      return { success: false, message: 'name, email, password, roleId, category and subcategory are required' };
    if (!isValidEmail(b.email))
      return { success: false, message: 'Valid email required' };
    var resolvedRoleId = resolveRoleId(b.roleId);
    if (!resolvedRoleId)
      return { success: false, message: 'Invalid role selected' };
    if (String(b.password || '').length < 8)
      return { success: false, message: 'Password must be at least 8 characters' };
    if (!isValidMobile(b.mobile))
      return { success: false, message: 'Invalid mobile number' };

    if (getUserByEmail(b.email))
      return { success: false, message: 'Email already registered' };

    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      // ── Numeric user_id ──
      var userId = nextId(SH.USERS);
      var roleId = String(resolvedRoleId).trim();
      var categoryId = toCleanText(b.departmentId || '', 120);
      var subcategoryId = toCleanText(b.subcategoryId || '', 120);

      // Default location: first row of AttendanceLocations (id = 1)
      var defaultLoc = getCached(SH.ATT_LOCATIONS, TTL_LOOKUP)[0];

      var tenant = getTenantConfig(currentGuid());
      var instituteName = tenant && tenant.institution ? tenant.institution.name : 'BioAttend Organization';

      getSheet(SH.USERS).appendRow([
        userId,
        toCleanText(b.instituteId || instituteName, 120),
        categoryId,
        parseInt(roleId, 10),
        toCleanText(b.name, 120),
        b.dob     || '',
        toCleanText(b.mobile, 20),
        String(b.email).toLowerCase(),
        hashPw(b.password),
        toCleanText(b.biometricCode || b.credentialId || '', 120),
        toCleanText(b.deviceId, 120),
        subcategoryId,
        currentGuid()
      ]);

      addToUserIndex(userId, b.email);

      // ── Numeric map id ──
      var locId = b.attendanceLocationId
        ? parseInt(b.attendanceLocationId, 10)
        : (defaultLoc ? parseInt(defaultLoc.attendance_location_id, 10) : 1);
      var mapId = nextId(SH.USER_LOC_MAP);
      getSheet(SH.USER_LOC_MAP).appendRow([
        mapId,                         // integer
        userId,
        locId,
        b.allowedDistance || DEFAULT_RADIUS
      ]);

      invalidate(SH.USERS);
      appendAuditLog('register_user', userId, normalizeRoleValue(roleId), userId, {
        email: String(b.email).toLowerCase(),
        categoryId: categoryId,
        subcategoryId: subcategoryId
      });
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
    if (!verifyPw(b.password || '', user.password_hash))
      return { success: false, message: 'Incorrect password' };
    if (String(user.password_hash || '').indexOf('v2$') !== 0) {
      updateUserPasswordHash(user.user_id, hashPw(b.password || ''));
      user = getUserByEmail(b.email) || user;
    }
    var roleName = normalizeRoleValue(user.role_id);
    var deviceId = String(b.deviceId || b.device_id || '').trim();
    var storedDevice = String(user.device_identification || '').trim();
    if (storedDevice && deviceId && storedDevice !== deviceId) {
      appendAuditLog('new_device_login', user.user_id, roleName, user.user_id, {
        email: String(user.email || ''),
        storedDeviceId: storedDevice,
        deviceId: deviceId
      });
    }
    var auth = createAuthSession(user.user_id, roleName || String(user.role_id || ''), deviceId, currentGuid());
    appendAuditLog('sign_in', user.user_id, roleName, user.user_id, {
      email: String(user.email || ''),
      deviceId: deviceId || storedDevice || ''
    });
    return {
      success:   true,
      userId:    user.user_id,
      name:      user.full_name,
      roleId:    roleName || String(user.role_id || ''),
      roleKey:   roleName || String(user.role_id || ''),
      roleDbId:  String(user.role_id || ''),
      deptId:    user.department_id,
      subcategoryId: user.subcategory_id || '',
      authToken: auth.authToken,
      refreshToken: auth.refreshToken,
      authExpiresAt: auth.expiresAt,
      refreshExpiresAt: auth.refreshExpiresAt,
      sessionId: auth.sessionId,
      deviceId: deviceId || storedDevice || '',
      guid: currentGuid()
    };
  } catch(err) { return { success: false, message: 'signIn: ' + err }; }
}

function loginByBiometric(b) {
  try {
    return { success: false, message: 'Deprecated. Use beginWebAuthnLogin/finishWebAuthnLogin.' };
  } catch (err) {
    return { success: false, message: 'loginByBiometric: ' + err };
  }
}

// ============================================================
//  3. MARK ENTRY
//  attendance_id → nextId(Attendance)       — integer
//  location_monitor_id → nextId(LocationMonitor) — integer
// ============================================================

function markEntry(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };
    var latCheck = (b.latitude === '' || typeof b.latitude === 'undefined') ? null : parseNumber(b.latitude);
    var lngCheck = (b.longitude === '' || typeof b.longitude === 'undefined') ? null : parseNumber(b.longitude);
    if ((latCheck !== null || lngCheck !== null) && !isValidCoordinate(latCheck, lngCheck))
      return { success: false, message: 'Invalid coordinates' };
    if (latCheck === null || lngCheck === null)
      return { success: false, code: 'LOCATION_REQUIRED', message: 'GPS location is required to mark attendance' };

    var now  = new Date();
    var t    = tz();
    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };

    var deviceCheck  = String(b.deviceId || '').trim();
    var boundDevice  = String(user.device_identification || '').trim();
    if (boundDevice && deviceCheck && boundDevice !== deviceCheck)
      return { success: false, message: 'Attendance blocked: device mismatch' };

    // Resolve mapped attendance location for the user's category/subcategory
    var userLoc = resolveUserTrackingLocation(b.userId);
    var geofence = getTenantGeofence(currentGuid());
    var allowedDist = userLoc.allowedDistance || geofence.radius;
    var locId       = userLoc.attendanceLocationId || 1;
    var sessionCheck = validateSession(b.userId, locId);
    if (!sessionCheck.success) return sessionCheck;

    var anchorLat = isNaN(parseFloat(userLoc.anchorLat)) ? geofence.latitude : parseFloat(userLoc.anchorLat);
    var anchorLng = isNaN(parseFloat(userLoc.anchorLng)) ? geofence.longitude : parseFloat(userLoc.anchorLng);

    var lat = latCheck, lng = lngCheck, dist = 0;
    if (lat !== null && !isNaN(lat)) {
      dist = haversine(lat, lng, anchorLat, anchorLng);
      if (dist > allowedDist) {
        return { success: false, code: 'TOO_FAR', distance: dist, allowed: allowedDist,
                 message: 'You are ' + dist + 'm away. Must be within ' + allowedDist + 'm of campus.' };
      }
    }

    // Duplicate check
    var attSheet = getSheet(SH.ATTENDANCE);
    var lastRow  = attSheet.getLastRow();
    var dateStr  = Utilities.formatDate(now, t, 'yyyy-MM-dd');
    if (lastRow >= 2) {
      var cols = attendanceReadColumnCount(attSheet);
      var rows = attSheet.getRange(2, 1, lastRow - 1, cols).getValues();
      for (var j = 0; j < rows.length; j++) {
        var row = rows[j];
        var rowStatus = String(row[13] || '').trim().toUpperCase();
        if (String(row[1]).trim() === String(b.userId).trim() &&
            normDate(row[6], t) === dateStr &&
            (String(row[3]).trim() === 'entry' || rowStatus === 'IN')) {
          return { success: false, message: 'Attendance already marked for today.' };
        }
      }
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var timeStr = Utilities.formatDate(now, t, 'HH:mm:ss');
      var attId   = nextId(SH.ATTENDANCE);   // ← integer

      attSheet.appendRow([
        attId,
        b.userId,
        user.full_name,
        'entry',
        timeStr,
        '',
        dateStr,
        b.loginMethod || 'biometric',
        lat !== null ? lat : '',
        lng !== null ? lng : '',
        locId,
        b.address || '',
        dist || '',
        'IN',
        currentGuid(),
        userCategoryOf(user),
        userSubcategoryOf(user)
      ]);

      if (lat !== null) {
        var lmId = nextId(SH.LOC_MONITOR);   // ← integer
        getSheet(SH.LOC_MONITOR).appendRow([lmId, b.userId, lat, lng, dist, now.toISOString()]);
      }

      SpreadsheetApp.flush();
      invalidate(SH.ATTENDANCE);
      touchLiveState('mark_entry');

      var activeSession = getOpenSessionForToday();
      if (activeSession) {
        pushTeacherNotification({
          id: 'ntf_' + Date.now(),
          sessionId: String(activeSession.session_id || ''),
          teacherId: String(activeSession.teacher_id || ''),
          userId: String(b.userId || ''),
          studentName: String(user.full_name || ''),
          message: String(user.full_name || 'Student') + ' marked attendance at ' + timeStr,
          time: timeStr, date: dateStr, createdAt: now.toISOString()
        });
      }
      appendAuditLog('mark_entry', b.userId, normalizeRoleValue(user.role_id), b.userId,
        { time: timeStr, date: dateStr, locationId: locId, distance: dist, categoryId: userCategoryOf(user), subcategoryId: userSubcategoryOf(user) });

      return {
        success: true, attendanceId: attId,
        message: '\u2713 Attendance marked at ' + timeStr,
        name: user.full_name, date: dateStr, time: timeStr,
        location: b.address || (lat ? lat + ', ' + lng : 'not captured'),
        latitude: lat !== null ? lat : '',
        longitude: lng !== null ? lng : '',
        distanceFromCentre: dist,
        sessionId: sessionCheck.window.sessionId || ''
      };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success: false, message: 'markEntry: ' + err }; }
}

// ============================================================
//  4. MARK EXIT
// ============================================================

function markExit(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };
    var latCheck = (b.latitude === '' || typeof b.latitude === 'undefined') ? null : parseNumber(b.latitude);
    var lngCheck = (b.longitude === '' || typeof b.longitude === 'undefined') ? null : parseNumber(b.longitude);
    if ((latCheck !== null || lngCheck !== null) && !isValidCoordinate(latCheck, lngCheck))
      return { success: false, message: 'Invalid coordinates' };

    var now     = new Date();
    var t       = tz();
    var dateStr = Utilities.formatDate(now, t, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, t, 'HH:mm:ss');
    var sheet   = getSheet(SH.ATTENDANCE);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No attendance records found' };

    var data = sheet.getRange(2, 1, lastRow - 1, attendanceReadColumnCount(sheet)).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (String(row[1]).trim() !== String(b.userId).trim()) continue;
      if (normDate(row[6], t) !== dateStr) continue;
      if (String(row[3]).trim() !== 'entry') continue;
      if (String(row[5]).trim() || String(row[13] || '').trim().toUpperCase() === 'OUT') continue;

      var entryTimeStr = String(row[4] || '');
      var durationMins = '';
      if (entryTimeStr) {
        var entryDt = new Date(dateStr + 'T' + entryTimeStr);
        if (!isNaN(entryDt.getTime())) {
          durationMins = Math.max(0, Math.round((now - entryDt) / 60000));
        }
      }

      var xlat  = latCheck !== null ? latCheck : '';
      var xlng  = lngCheck !== null ? lngCheck : '';
      var mapped = resolveUserTrackingLocation(b.userId);
      var xdist = (xlat !== '' && xlng !== '') ? haversine(xlat, xlng, mapped.anchorLat, mapped.anchorLng) : '';

      var lock = LockService.getScriptLock();
      lock.waitLock(6000);
      try {
        var sheetRow = i + 2;
        sheet.getRange(sheetRow, 4, 1, 1).setValue('exit');
        sheet.getRange(sheetRow, 6, 1, 1).setValue(timeStr);
        sheet.getRange(sheetRow, 14, 1, 1).setValue('OUT');
        if (xlat !== '') {
          sheet.getRange(sheetRow, 9, 1, 5).setValues([[xlat, xlng, '', b.address || '', xdist]]);
        }
        if (xlat !== '') {
          var lmId = nextId(SH.LOC_MONITOR);  // ← integer
          getSheet(SH.LOC_MONITOR).appendRow([lmId, b.userId, xlat, xlng, xdist, now.toISOString()]);
        }
        SpreadsheetApp.flush();
        invalidate(SH.ATTENDANCE);
        touchLiveState('mark_exit');
      } finally { lock.releaseLock(); }

      appendAuditLog('mark_exit', b.userId, '', b.userId, { time: timeStr, date: dateStr, distance: xdist, categoryId: userCategoryOf(user), subcategoryId: userSubcategoryOf(user) });
      var hrs = durationMins !== '' ? Math.floor(durationMins / 60) : 0;
      var mins = durationMins !== '' ? durationMins % 60 : 0;
      var durLabel = durationMins !== '' ? (hrs > 0 ? hrs + 'h ' + mins + 'm' : mins + 'm') : '';
      return {
        success: true,
        message: '\u2713 Exit recorded at ' + timeStr + (durLabel ? ' \u00b7 ' + durLabel : ''),
        exitTime: timeStr, duration: durLabel, durationMins: durationMins,
        location: b.address || (xlat ? xlat + ', ' + xlng : 'not captured'), distance: xdist
      };
    }
    return { success: false, message: 'No entry record found for today (' + dateStr + '). Mark attendance first.' };
  } catch(err) { return { success: false, message: 'markExit: ' + err }; }
}

// ============================================================
//  5. TRACK STUDENT LOCATION
// ============================================================

function trackLocation(b) {
  try {
    var userId = String(b.user_id || b.userId || '').trim();
    if (!userId) return { success: false, message: 'userId required' };
    var lat = parseFloat(b.latitude), lng = parseFloat(b.longitude);
    if (isNaN(lat) || isNaN(lng)) return { success: false, message: 'invalid coordinates' };
    if (!isValidCoordinate(lat, lng)) return { success: false, message: 'invalid coordinates' };
    var user = getUserById(userId);
    if (!user) return { success: false, message: 'User not found' };

    var track = resolveUserTrackingLocation(userId);
    var distance = calculateDistance(lat, lng, track.anchorLat, track.anchorLng);
    var cache = CacheService.getScriptCache();
    var now = new Date();
    var lastTrackTs = parseInt(cache.get('track_last_ts_' + userId) || '0', 10);
    var lastTrackLat = parseFloat(cache.get('track_last_lat_' + userId) || '');
    var lastTrackLng = parseFloat(cache.get('track_last_lng_' + userId) || '');
    if (lastTrackTs && !isNaN(lastTrackLat) && !isNaN(lastTrackLng)) {
      var elapsedMs = now.getTime() - lastTrackTs;
      var movement = calculateDistance(lat, lng, lastTrackLat, lastTrackLng);
      if (elapsedMs > 0 && elapsedMs < 10000 && movement > 500) {
        appendAuditLog('gps_spoofing_suspected', userId, normalizeRoleValue(user.role_id), userId, {
          movement: movement,
          elapsedMs: elapsedMs,
          latitude: lat,
          longitude: lng
        });
      }
    }
    cache.put('track_last_ts_' + userId, String(now.getTime()), 300);
    cache.put('track_last_lat_' + userId, String(lat), 300);
    cache.put('track_last_lng_' + userId, String(lng), 300);

    var lastPersistTs = parseInt(cache.get('track_persist_ts_' + userId) || '0', 10);
    var lastPersistLat = parseFloat(cache.get('track_persist_lat_' + userId) || '');
    var lastPersistLng = parseFloat(cache.get('track_persist_lng_' + userId) || '');
    var shouldPersist = !lastPersistTs || (now.getTime() - lastPersistTs) >= 120000;
    if (!shouldPersist && !isNaN(lastPersistLat) && !isNaN(lastPersistLng)) {
      shouldPersist = calculateDistance(lat, lng, lastPersistLat, lastPersistLng) > 10;
    }
    if (shouldPersist) {
      saveLocation(userId, lat, lng, distance);
      cache.put('track_persist_ts_' + userId, String(now.getTime()), 300);
      cache.put('track_persist_lat_' + userId, String(lat), 300);
      cache.put('track_persist_lng_' + userId, String(lng), 300);
    }

    var cacheKey = 'exit_counter_' + userId;
    var limit = track.allowedDistance + 20;
    var count = parseInt(cache.get(cacheKey) || '0', 10);
    if (distance > limit) {
      count++;
      cache.put(cacheKey, String(count), 300);
    } else {
      count = 0;
      cache.put(cacheKey, '0', 300);
    }

    var exitMarked = false;
    if (count >= 3) {
      var exitResult = markExit({
        userId: userId,
        latitude: lat,
        longitude: lng,
        address: b.address || ''
      });
      cache.remove(cacheKey);
      cache.remove('track_persist_ts_' + userId);
      cache.remove('track_persist_lat_' + userId);
      cache.remove('track_persist_lng_' + userId);
      exitMarked = !!(exitResult && exitResult.success);
      if (!exitMarked && exitResult && /No entry record found/i.test(String(exitResult.message || ''))) {
        exitMarked = true;
      }
      if (exitMarked) touchLiveState('auto_exit');
    }

    return { success: true, distance: distance, exitMarked: exitMarked };
  } catch(err) { return { success: false, message: 'trackLocation: ' + err }; }
}

function trackStudentLocation(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };
    var lat = parseFloat(b.latitude), lng = parseFloat(b.longitude);
    if (isNaN(lat) || isNaN(lng)) return { success: false, message: 'invalid coordinates' };
    if (!isValidCoordinate(lat, lng)) return { success: false, message: 'invalid coordinates' };
    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var now = new Date();
    var track = resolveUserTrackingLocation(b.userId);
    var dist = haversine(lat, lng, track.anchorLat, track.anchorLng);
    var lmId = nextId(SH.LOC_MONITOR);   // ← integer
    getSheet(SH.LOC_MONITOR).appendRow([lmId, b.userId, lat, lng, dist, now.toISOString(), currentGuid()]);
    touchLiveState('track_student_location');
    return { success: true, message: 'Location tracked', distanceFromCentre: dist,
             attendanceLocationId: track.attendanceLocationId };
  } catch(err) { return { success: false, message: 'trackStudentLocation: ' + err }; }
}

// ============================================================
//  6. GET MY ATTENDANCE
// ============================================================

function getMyAttendance(b) {
  try {
    if (!b.userId) return { success: false, message: 'userId required' };
    var t = tz(), sheet = getSheet(SH.ATTENDANCE), lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, records: [] };
    var data = sheet.getRange(2, 1, lastRow - 1, attendanceReadColumnCount(sheet)).getValues();
    var entryMap = {}, exitMap = {};
    data.forEach(function(row) {
      if (String(row[1]).trim() !== String(b.userId).trim()) return;
      var date = normDate(row[6], t);
      var status = String(row[13] || '').trim().toUpperCase();
      if (String(row[3]).trim() === 'entry' || status === 'IN') entryMap[date] = row;
      if (String(row[3]).trim() === 'exit' || status === 'OUT')  exitMap[date]  = row;
    });
    var records = Object.keys(entryMap).map(function(date) {
      var e = entryMap[date], x = exitMap[date] || null;
      var eDur = '';
      if (x) {
        var eDt = new Date(date + 'T' + String(e[4]));
        var xDt = new Date(date + 'T' + String(x[5]));
        if (!isNaN(eDt.getTime()) && !isNaN(xDt.getTime())) {
          var dm = Math.max(0, Math.round((xDt - eDt) / 60000));
          var h = Math.floor(dm/60), m = dm%60;
          eDur = h > 0 ? h+'h '+m+'m' : m+'m';
        }
      }
      return { date: date, entryTime: e[4]||'', exitTime: x ? x[5]:'',
               duration: eDur, loginMethod: e[7]||'', address: e[11]||'',
               distanceFromCentre: e[12]||'', status: (x ? 'OUT' : 'IN') };
    });
    records.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(err) { return { success: false, message: 'getMyAttendance: ' + err }; }
}

// ============================================================
//  7. SESSIONS
//  session_id → nextId(Sessions) — integer
// ============================================================

function createSession(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    if (!b.subject || !b.windowMinutes)
      return { success: false, message: 'subject and windowMinutes required' };
    var t = tz(), now = new Date();
    var sessId = nextId(SH.SESSIONS);   // ← integer
    var date   = Utilities.formatDate(now, t, 'yyyy-MM-dd');
    var start  = Utilities.formatDate(now, t, 'HH:mm:ss');
    var end    = Utilities.formatDate(new Date(now.getTime() + b.windowMinutes*60000), t, 'HH:mm:ss');
    var sheet  = getSheet(SH.SESSIONS);
    var data   = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(auth.session.user_id) && data[i][7] === 'open')
        sheet.getRange(i+1, 8).setValue('closed');
    }
    sheet.appendRow([sessId, auth.session.user_id, b.teacherName||'', b.subject, date, start, end, 'open', b.windowMinutes, currentGuid()]);
        CacheService.getScriptCache().put(tenantScopedKey('active_session'), JSON.stringify({
      session_id: sessId, teacher_id: auth.session.user_id, subject: b.subject,
      date: date, start_time: start, end_time: end, status: 'open', window_minutes: b.windowMinutes
    }), TTL_SESSION);
    appendAuditLog('create_session', auth.session.user_id, auth.session.role_name, '',
      { sessionId: sessId, subject: b.subject });
    touchLiveState('create_session');
    return { success: true, sessionId: sessId, subject: b.subject, startTime: start, endTime: end };
  } catch(err) { return { success: false, message: 'createSession: ' + err }; }
}

function closeSession(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var sheet = getSheet(SH.SESSIONS), data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(b.sessionId)) {
        if (String(auth.session.role_name) !== 'admin' && String(data[i][1]) !== String(auth.session.user_id))
          return { success: false, message: 'Access denied' };
        sheet.getRange(i+1, 8).setValue('closed');
        CacheService.getScriptCache().remove(tenantScopedKey('active_session'));
        appendAuditLog('close_session', auth.session.user_id, auth.session.role_name, '', { sessionId: b.sessionId });
        touchLiveState('close_session');
        return { success: true };
      }
    }
    return { success: false, message: 'Session not found' };
  } catch(err) { return { success: false, message: 'closeSession: ' + err }; }
}

function getActiveSession(b) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(tenantScopedKey('active_session'));
    if (cached) {
      var s = JSON.parse(cached), now = new Date(), t = tz();
      var today = Utilities.formatDate(now, t, 'yyyy-MM-dd');
      if (s.status === 'open' && s.date === today) {
        var endDt = new Date(s.date + 'T' + s.end_time);
        if (now <= endDt)
          return { success: true, active: true, session: s, secondsLeft: Math.max(0, Math.round((endDt-now)/1000)) };
      }
      cache.remove(tenantScopedKey('active_session'));
    }
    var autoSession = getCurrentAutoSession(new Date());
    if (autoSession) {
      var autoEnd = new Date(autoSession.date + 'T' + autoSession.end_time);
      return { success: true, active: true, session: autoSession,
               secondsLeft: Math.max(0, Math.round((autoEnd-new Date())/1000)) };
    }
    var t2 = tz(), today2 = Utilities.formatDate(new Date(), t2, 'yyyy-MM-dd');
    var rows = getRows(getSheet(SH.SESSIONS)), now2 = new Date();
    for (var i = 0; i < rows.length; i++) {
      var s = rows[i];
      if (s.status === 'open' && normDate(s.date, t2) === today2) {
        var st = new Date(s.date + 'T' + s.start_time), en = new Date(s.date + 'T' + s.end_time);
        if (now2 >= st && now2 <= en) {
          var secs = Math.max(0, Math.round((en-now2)/1000));
          cache.put(tenantScopedKey('active_session'), JSON.stringify(s), TTL_SESSION);
          return { success: true, active: true, session: s, secondsLeft: secs };
        }
      }
    }
    return { success: true, active: false };
  } catch(err) { return { success: false, message: 'getActiveSession: ' + err }; }
}

function getSessions(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var t = tz(), rows = getRows(getSheet(SH.SESSIONS));
    var today = Utilities.formatDate(new Date(), t, 'yyyy-MM-dd');
    var autoRows = getAutoSessionsForDate(today);
    rows = autoRows.concat(rows);
    if (b.userId) {
      rows = rows.filter(function(r) {
        return String(r.teacher_id) === String(auth.session.user_id) || String(r.teacher_id) === 'ALL';
      });
    }
    var attSheet = getSheet(SH.ATTENDANCE), attLast = attSheet.getLastRow(), countMap = {};
    if (attLast >= 2) {
      var uids  = attSheet.getRange(2, 2, attLast-1, 1).getValues();
      var dates = attSheet.getRange(2, 7, attLast-1, 1).getValues();
      var types = attSheet.getRange(2, 4, attLast-1, 1).getValues();
      uids.forEach(function(u, i) {
        if (String(types[i][0]).trim() !== 'entry') return;
        var key = normDate(dates[i][0], t);
        countMap[key] = (countMap[key] || 0) + 1;
      });
    }
    return {
      success: true,
      sessions: rows.map(function(s) {
        var status = s.status;
        if (s.is_auto) {
          var st = new Date(s.date+'T'+s.start_time), en = new Date(s.date+'T'+s.end_time), now = new Date();
          status = now >= st && now <= en ? 'open' : 'scheduled';
        }
        return { sessionId: s.session_id, subject: s.subject, date: normDate(s.date, t),
                 startTime: s.start_time, endTime: s.end_time, status: status,
                 presentCount: countMap[normDate(s.date, t)] || 0 };
      }).reverse().slice(0, 20)
    };
  } catch(err) { return { success: false, message: 'getSessions: ' + err }; }
}

function getTeacherNotifications(b) {
  try {
    if (!b.sessionId) return { success: false, message: 'sessionId required' };
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var cache = CacheService.getScriptCache();
    var rows = [];
    var cached = cache.get(tenantScopedKey('teacher_notif_' + b.sessionId));
    if (cached) { try { rows = JSON.parse(cached) || []; } catch(e) {} }
    if (b.teacherId) {
      rows = rows.filter(function(r) {
        var tid = String(r.teacherId || '');
        return tid === String(auth.session.user_id) || tid === 'ALL';
      });
    }
    return { success: true, notifications: rows };
  } catch(err) { return { success: false, message: 'getTeacherNotifications: ' + err }; }
}

// ============================================================
//  8. DASHBOARD
// ============================================================

function getDashboard(b) {
  try {
    if (!b.sessionId) return { success: false, message: 'sessionId required' };
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var cacheKey = tenantScopedKey('dash_' + b.sessionId);
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
    var t = tz(), today = Utilities.formatDate(new Date(), t, 'yyyy-MM-dd');
    var userRows = getCached(SH.USERS, TTL_LOOKUP);
    var studentRoleId = findRoleIdByName('student');
    var students = userRows.filter(function(u) {
      return normalizeRoleValue(u.role_id) === 'student' ||
             (studentRoleId && String(u.role_id) === studentRoleId);
    });
    var userMap = buildMap(students, 'user_id');
    var attSheet = getSheet(SH.ATTENDANCE), attLast = attSheet.getLastRow(), presentMap = {};
    if (attLast >= 2) {
      var rows = attSheet.getRange(2, 1, attLast-1, attendanceReadColumnCount(attSheet)).getValues();
      rows.forEach(function(row) {
        var rowStatus = String(row[13] || '').trim().toUpperCase();
        if (String(row[3]).trim() !== 'entry' && rowStatus !== 'IN') return;
        if (normDate(row[6], t) !== today) return;
        presentMap[String(row[1]).trim()] = { entryTime: row[4]||'', exitTime: row[5]||'',
          method: row[7]||'', address: row[11]||'', distance: row[12]||'', status: rowStatus || 'IN' };
      });
    }
    var present = [], absent = [];
    students.forEach(function(s) {
      var uid = String(s.user_id).trim();
      if (presentMap[uid]) {
        var pm = presentMap[uid];
        present.push({ userId: uid, name: s.full_name, email: s.email, department: s.department_id,
          entryTime: pm.entryTime, exitTime: pm.exitTime, method: pm.method,
          address: pm.address, distance: pm.distance });
      } else {
        absent.push({ userId: uid, name: s.full_name, email: s.email, department: s.department_id });
      }
    });
    var result = { success: true, total: students.length, presentCount: present.length,
                   absentCount: absent.length, present: present, absent: absent };
    try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), TTL_DASH); } catch(e) {}
    return result;
  } catch(err) { return { success: false, message: 'getDashboard: ' + err }; }
}

function resolveLiveSessionContext(b) {
  var sessionId = String((b && (b.sessionId || b.session_id)) || '').trim();
  if (!sessionId) {
    var active = getActiveSession({});
    if (active && active.active && active.session) {
      sessionId = String(active.session.session_id || '');
    }
  }
  if (!sessionId) {
    var autoSession = getCurrentAutoSession(new Date());
    if (autoSession) sessionId = String(autoSession.session_id || '');
  }
  var sessionRow = null;
  if (sessionId) {
    var rows = getRows(getSheet(SH.SESSIONS));
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i].session_id) === sessionId) { sessionRow = rows[i]; break; }
    }
    if (!sessionRow) {
      var auto = getCurrentAutoSession(new Date());
      if (auto && String(auto.session_id) === sessionId) sessionRow = auto;
    }
  }
  if (!sessionRow) {
    var fallback = getCurrentAutoSession(new Date());
    if (fallback) {
      sessionRow = fallback;
      sessionId = String(fallback.session_id || '');
    }
  }
  return { sessionId: sessionId, session: sessionRow };
}

function latestAuthSessionsByUser() {
  var rows = getRows(getAuthSessionSheet());
  var map = {};
  rows.forEach(function(r) {
    var uid = String(r.user_id || '').trim();
    if (!uid) return;
    var status = String(r.status || '').trim().toUpperCase();
    var exp = new Date(r.expires_at || r.refresh_expires_at || '');
    if (status !== 'ACTIVE' || isNaN(exp.getTime()) || exp.getTime() < Date.now()) return;
    var existing = map[uid];
    if (!existing) {
      map[uid] = r;
      return;
    }
    var current = new Date(existing.last_seen_at || existing.created_at || 0).getTime();
    var incoming = new Date(r.last_seen_at || r.created_at || 0).getTime();
    if (incoming >= current) map[uid] = r;
  });
  return map;
}

function latestAttendanceByUser(sessionId) {
  var sheet = getSheet(SH.ATTENDANCE);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var cols = attendanceReadColumnCount(sheet);
  var data = sheet.getRange(2, 1, lastRow - 1, cols).getValues();
  var map = {};
  var t = tz();
  var sessionCtx = resolveLiveSessionContext({ sessionId: sessionId });
  var sessionRow = sessionCtx.session || null;
  data.forEach(function(row) {
    var uid = String(row[1] || '').trim();
    if (!uid) return;
    if (sessionRow && !attendanceMatchesSession(row, sessionRow, t)) return;
    var existing = map[uid];
    var currentTime = String(row[5] || row[4] || '');
    var currentSort = (normDate(row[6], t) || '') + 'T' + currentTime;
    if (!existing) {
      map[uid] = row;
      map[uid]._sort = currentSort;
      return;
    }
    if (currentSort >= (existing._sort || '')) {
      row._sort = currentSort;
      map[uid] = row;
    }
  });
  return map;
}

function attendanceMatchesSession(row, sessionRow, tzName) {
  try {
    if (!sessionRow) return true;
    var rowDate = normDate(row[6], tzName || tz());
    var sessionDate = normDate(sessionRow.date || sessionRow.start_time || '', tzName || tz());
    if (sessionDate && rowDate && sessionDate !== rowDate) return false;
    var entryTime = String(row[4] || '').trim();
    if (!entryTime) return true;
    var startTime = String(sessionRow.start_time || '').trim();
    var endTime = String(sessionRow.end_time || '').trim();
    if (!startTime || !endTime) return true;
    var entryDt = new Date((sessionDate || rowDate) + 'T' + entryTime);
    var startDt = new Date((sessionDate || rowDate) + 'T' + startTime);
    var endDt = new Date((sessionDate || rowDate) + 'T' + endTime);
    if (isNaN(entryDt.getTime()) || isNaN(startDt.getTime()) || isNaN(endDt.getTime())) return true;
    return entryDt >= startDt && entryDt <= endDt;
  } catch (e) {
    return true;
  }
}

function latestLocationsByUser() {
  var sheet = getSheet(SH.LOC_MONITOR);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var map = {};
  data.forEach(function(row) {
    var uid = String(row[1] || '').trim();
    if (!uid) return;
    var existing = map[uid];
    var currentTs = new Date(row[5] || 0).getTime();
    if (!existing || currentTs >= new Date(existing.timestamp || 0).getTime()) {
      map[uid] = {
        locationMonitorId: row[0] || '',
        userId: uid,
        latitude: row[2] || '',
        longitude: row[3] || '',
        distanceFromCentre: row[4] || '',
        timestamp: row[5] || ''
      };
    }
  });
  return map;
}

function liveCacheKey(action, sessionId, extra) {
  return tenantScopedKey('live_' + action + '_' + String(sessionId || 'all') + '_' + String(extra || ''));
}

function getLiveAttendance(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var liveCtx = resolveLiveSessionContext(b);
    var sessionId = liveCtx.sessionId || '';
    if (!sessionId) return { success: true, totalIn: 0, totalOut: 0, activeUsers: [], offlineUsers: [], recentlyExited: [], sessionId: '' };

    var liveState = getLiveStateInfo();
    var cacheKey = liveCacheKey('attendance', sessionId, liveState.version);
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    var t = tz();
    var now = new Date();
    var sessionRow = liveCtx.session || null;
    var sessionUsers = getCached(SH.USERS, TTL_LOOKUP).filter(function(u) {
      return normalizeRoleValue(u.role_id) === 'student' || String(findRoleIdByName('student')) === String(u.role_id);
    });
    var attendanceMap = latestAttendanceByUser(sessionId);
    var locationMap = latestLocationsByUser();
    var authMap = latestAuthSessionsByUser();
    var activeUsers = [];
    var offlineUsers = [];
    var recentlyExited = [];
    var totalIn = 0;
    var totalOut = 0;
    sessionUsers.forEach(function(u) {
      var uid = String(u.user_id || '').trim();
      var att = attendanceMap[uid] || null;
      var authRow = authMap[uid] || null;
      var lastSeenAt = authRow ? String(authRow.last_seen_at || authRow.created_at || '') : '';
      var lastSeenMs = lastSeenAt ? new Date(lastSeenAt).getTime() : 0;
      var isOnline = !!(lastSeenMs && (now.getTime() - lastSeenMs) <= 120000);
      var isIn = att && (String(att[3]).trim() === 'entry' || String(att[13] || '').toUpperCase() === 'IN') && !String(att[5] || '').trim();
      var isOut = att && (String(att[3]).trim() === 'exit' || String(att[13] || '').toUpperCase() === 'OUT' || String(att[5] || '').trim());
      var loc = locationMap[uid] || null;
      var base = {
        userId: uid,
        name: u.full_name,
        email: u.email,
        department: u.department_id,
        status: isIn ? 'IN' : (isOut ? 'OUT' : 'UNKNOWN'),
        online: isOnline,
        lastSeenAt: lastSeenAt || '',
        entryTime: att ? (att[4] || '') : '',
        exitTime: att ? (att[5] || '') : '',
        sessionId: sessionId,
        location: loc
      };
      if (isIn) {
        totalIn++;
        activeUsers.push(base);
      } else if (isOut) {
        totalOut++;
        if (att && String(att[5] || '').trim()) {
          var exitMs = new Date((normDate(att[6], t) || '') + 'T' + String(att[5] || '')).getTime();
          if (exitMs && (now.getTime() - exitMs) <= 300000) recentlyExited.push(base);
        }
        offlineUsers.push(base);
      } else {
        offlineUsers.push(base);
      }
    });

    var result = {
      success: true,
      sessionId: sessionId,
      session: liveCtx.session || null,
      totalIn: totalIn,
      totalOut: totalOut,
      activeUsers: activeUsers,
      offlineUsers: offlineUsers,
      recentlyExited: recentlyExited,
      updatedAt: now.toISOString()
    };
    try { cache.put(cacheKey, JSON.stringify(result), 5); } catch (e) {}
    return result;
  } catch (err) {
    return { success: false, message: 'getLiveAttendance: ' + err };
  }
}

function getUserLocations(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var liveCtx = resolveLiveSessionContext(b);
    var sessionId = liveCtx.sessionId || '';
    var liveState = getLiveStateInfo();
    var cacheKey = liveCacheKey('locations', sessionId, liveState.version);
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    var users = getCached(SH.USERS, TTL_LOOKUP);
    var authMap = latestAuthSessionsByUser();
    var locMap = latestLocationsByUser();
    var rows = users.map(function(u) {
      var uid = String(u.user_id || '').trim();
      var loc = locMap[uid] || null;
      var authRow = authMap[uid] || null;
      return {
        userId: uid,
        name: u.full_name,
        email: u.email,
        department: u.department_id,
        latitude: loc ? loc.latitude : '',
        longitude: loc ? loc.longitude : '',
        distanceFromCentre: loc ? loc.distanceFromCentre : '',
        timestamp: loc ? loc.timestamp : '',
        online: !!(authRow && authRow.last_seen_at && (Date.now() - new Date(authRow.last_seen_at).getTime()) <= 120000)
      };
    });
    var result = { success: true, sessionId: sessionId, locations: rows, updatedAt: new Date().toISOString() };
    try { cache.put(cacheKey, JSON.stringify(result), 5); } catch (e) {}
    return result;
  } catch (err) {
    return { success: false, message: 'getUserLocations: ' + err };
  }
}

function getLiveUpdates(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var liveCtx = resolveLiveSessionContext(b);
    var sessionId = liveCtx.sessionId || '';
    var lastSyncTime = String((b && b.lastSyncTime) || '').trim();
    var lastSyncMs = lastSyncTime ? new Date(lastSyncTime).getTime() : 0;
    var waitMs = Math.min(parseInt((b && (b.waitMs || b.waitSeconds * 1000)) || 0, 10) || 0, 5000);
    var result = {
      success: true,
      sessionId: sessionId,
      updates: [],
      liveAttendance: null,
      locations: null,
      syncedAt: new Date().toISOString()
    };

    function collectUpdates() {
      var updates = [];
      if (!lastSyncMs) return updates;
      var auditRows = getRows(getSheet(SH.AUDIT_LOG));
      auditRows.forEach(function(row) {
        var ts = new Date(row.created_at || 0).getTime();
        if (ts > lastSyncMs) {
          updates.push({
            type: 'audit',
            action: row.event_type,
            userId: row.actor_user_id,
            targetUserId: row.target_user_id,
            createdAt: row.created_at,
            metadata: safeJsonParse(String(row.details || '{}')) || {}
          });
        }
      });
      var locRows = getRows(getSheet(SH.LOC_MONITOR));
      locRows.forEach(function(row) {
        var ts2 = new Date(row.timestamp || 0).getTime();
        if (ts2 > lastSyncMs) {
          updates.push({
            type: 'location',
            userId: row.user_id,
            latitude: row.latitude,
            longitude: row.longitude,
            distanceFromCentre: row.distance_from_centre,
            timestamp: row.timestamp
          });
        }
      });
      return updates;
    }

    var updates = collectUpdates();
    if (!updates.length && waitMs > 0 && lastSyncMs) {
      var startWait = Date.now();
      var beforeVersion = getLiveStateInfo().version;
      while ((Date.now() - startWait) < waitMs) {
        Utilities.sleep(1000);
        if (getLiveStateInfo().version !== beforeVersion) {
          updates = collectUpdates();
          break;
        }
      }
    }

    result.updates = updates.slice(0, 100);
    result.liveAttendance = getLiveAttendance(b);
    result.locations = getUserLocations(b);
    return result;
  } catch (err) {
    return { success: false, message: 'getLiveUpdates: ' + err };
  }
}

function forceExitUser(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    if (!b.userId) return { success: false, message: 'userId required' };
    var res = markExit({
      userId: b.userId,
      latitude: b.latitude || '',
      longitude: b.longitude || '',
      address: b.address || ''
    });
    appendAuditLog('force_exit_user', auth.session.user_id, auth.session.role_name, b.userId, {
      sessionId: String((b && b.sessionId) || ''),
      forced: true
    });
    touchLiveState('force_exit');
    return res;
  } catch (err) {
    return { success: false, message: 'forceExitUser: ' + err };
  }
}

function getAttendanceRowsForAnalytics() {
  return getRows(getSheet(SH.ATTENDANCE));
}

function getDailyStats(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var t = tz();
    var date = String((b && b.date) || Utilities.formatDate(new Date(), t, 'yyyy-MM-dd'));
    var cacheKey = 'daily_stats_' + date;
    var cached = getCachedData(cacheKey, 10);
    if (cached) return cached;

    var users = getCached(SH.USERS, TTL_LOOKUP).filter(function(u) {
      return normalizeRoleValue(u.role_id) === 'student' || String(findRoleIdByName('student')) === String(u.role_id);
    });
    var rows = getAttendanceRowsForAnalytics();
    var byUser = {};
    var entryHours = Array(24).fill(0);
    var exitHours = Array(24).fill(0);
    var lateEntries = 0;
    var earlyExits = 0;
    var presentCount = 0;
    var totalEntries = 0;
    var sessionWindows = getAttendanceWindowRows().filter(function(w) {
      return String(w.status || '').toLowerCase() === 'active';
    });

    rows.forEach(function(row) {
      if (normDate(row[6], t) !== date) return;
      var uid = String(row[1] || '').trim();
      var status = String(row[13] || '').trim().toUpperCase();
      var type = String(row[3] || '').trim().toLowerCase();
      if (type !== 'entry' && status !== 'IN' && type !== 'exit' && status !== 'OUT') return;
      var entryTime = String(row[4] || '').trim();
      var exitTime = String(row[5] || '').trim();
      if (entryTime) {
        totalEntries++;
        presentCount++;
        var eh = parseInt(entryTime.split(':')[0], 10);
        if (!isNaN(eh)) entryHours[eh] += 1;
        byUser[uid] = byUser[uid] || {};
        byUser[uid].entry = row;
        var matchedWindow = findMatchingWindowForRow(row, sessionWindows, t);
        if (matchedWindow && isLateEntry(row, matchedWindow, t)) lateEntries++;
      }
      if (exitTime) {
        var xh = parseInt(exitTime.split(':')[0], 10);
        if (!isNaN(xh)) exitHours[xh] += 1;
        byUser[uid] = byUser[uid] || {};
        byUser[uid].exit = row;
        var matchedExitWindow = findMatchingWindowForRow(row, sessionWindows, t);
        if (matchedExitWindow && isEarlyExit(row, matchedExitWindow, t)) earlyExits++;
      }
    });

    var result = {
      success: true,
      date: date,
      totalStudents: users.length,
      totalPresent: Object.keys(byUser).length,
      totalAbsent: Math.max(0, users.length - Object.keys(byUser).length),
      avgAttendancePercent: users.length ? Math.round((Object.keys(byUser).length / users.length) * 1000) / 10 : 0,
      lateEntries: lateEntries,
      earlyExits: earlyExits,
      entryTrend: entryHours,
      exitTrend: exitHours,
      updatedAt: new Date().toISOString()
    };
    setCachedData(cacheKey, result, 10);
    return result;
  } catch (err) {
    return { success: false, message: 'getDailyStats: ' + err };
  }
}

function getWeeklyStats(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var t = tz();
    var end = new Date();
    end.setHours(23, 59, 59, 999);
    var start = new Date(end.getTime() - 6 * 24 * 60 * 60 * 1000);
    var cacheKey = 'weekly_stats_' + Utilities.formatDate(end, t, 'yyyy-MM-dd');
    var cached = getCachedData(cacheKey, 10);
    if (cached) return cached;

    var users = getCached(SH.USERS, TTL_LOOKUP).filter(function(u) {
      return normalizeRoleValue(u.role_id) === 'student' || String(findRoleIdByName('student')) === String(u.role_id);
    });
    var rows = getAttendanceRowsForAnalytics();
    var days = [];
    for (var i = 0; i < 7; i++) {
      var dt = new Date(start.getTime() + i * 24 * 60 * 60 * 1000);
      days.push(Utilities.formatDate(dt, t, 'yyyy-MM-dd'));
    }
    var byDay = {};
    days.forEach(function(d) {
      byDay[d] = { present: 0, late: 0, early: 0, entries: 0 };
    });
    rows.forEach(function(row) {
      var d = normDate(row[6], t);
      if (!byDay[d]) return;
      var entryTime = String(row[4] || '').trim();
      var exitTime = String(row[5] || '').trim();
      if (entryTime) {
        byDay[d].present += 1;
        byDay[d].entries += 1;
      }
      var windows = getAttendanceWindowRows().filter(function(w) {
        return String(w.status || '').toLowerCase() === 'active';
      });
      var matched = findMatchingWindowForRow(row, windows, t);
      if (matched && entryTime && isLateEntry(row, matched, t)) byDay[d].late += 1;
      if (matched && exitTime && isEarlyExit(row, matched, t)) byDay[d].early += 1;
    });

    var series = days.map(function(d) {
      return {
        date: d,
        present: byDay[d].present,
        late: byDay[d].late,
        early: byDay[d].early,
        attendancePercent: users.length ? Math.round((byDay[d].present / users.length) * 1000) / 10 : 0
      };
    });

    var result = {
      success: true,
      range: { start: Utilities.formatDate(start, t, 'yyyy-MM-dd'), end: Utilities.formatDate(end, t, 'yyyy-MM-dd') },
      totalStudents: users.length,
      series: series,
      updatedAt: new Date().toISOString()
    };
    setCachedData(cacheKey, result, 10);
    return result;
  } catch (err) {
    return { success: false, message: 'getWeeklyStats: ' + err };
  }
}

function getUserAttendanceSummary(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var userId = String((b && (b.userId || b.targetUserId)) || '').trim();
    if (!userId) return { success: false, message: 'userId required' };
    var user = getUserById(userId);
    if (!user) return { success: false, message: 'User not found' };
    var t = tz();
    var cacheKey = 'user_summary_' + userId;
    var cached = getCachedData(cacheKey, 10);
    if (cached) return cached;

    var rows = getAttendanceRowsForAnalytics().filter(function(row) {
      return String(row[1] || '').trim() === userId;
    });
    var totalDays = {};
    var lateCount = 0;
    var earlyCount = 0;
    var presentCount = 0;
    var records = [];
    var windows = getAttendanceWindowRows().filter(function(w) {
      return String(w.status || '').toLowerCase() === 'active';
    });

    rows.forEach(function(row) {
      var date = normDate(row[6], t);
      if (!date) return;
      totalDays[date] = true;
      var entryTime = String(row[4] || '').trim();
      var exitTime = String(row[5] || '').trim();
      var matched = findMatchingWindowForRow(row, windows, t);
      if (entryTime) {
        presentCount += 1;
        if (matched && isLateEntry(row, matched, t)) lateCount += 1;
      }
      if (exitTime && matched && isEarlyExit(row, matched, t)) earlyCount += 1;
      records.push({
        date: date,
        entryTime: entryTime,
        exitTime: exitTime,
        status: String(row[13] || '').trim().toUpperCase() || (exitTime ? 'OUT' : 'IN'),
        loginMethod: row[7] || '',
        locationId: row[10] || '',
        distance: row[12] || ''
      });
    });

    records.sort(function(a, b) { return b.date.localeCompare(a.date); });
    var result = {
      success: true,
      userId: userId,
      name: user.full_name,
      email: user.email,
      totalRecords: records.length,
      presentCount: presentCount,
      lateEntries: lateCount,
      earlyExits: earlyCount,
      attendancePercent: records.length ? Math.round((presentCount / records.length) * 1000) / 10 : 0,
      recentRecords: records.slice(0, 20),
      updatedAt: new Date().toISOString()
    };
    setCachedData(cacheKey, result, 10);
    return result;
  } catch (err) {
    return { success: false, message: 'getUserAttendanceSummary: ' + err };
  }
}

function getAttendanceInsights(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var t = tz();
    var days = parseInt((b && b.days) || 14, 10);
    if (isNaN(days) || days < 7) days = 14;
    if (days > 60) days = 60;
    var cacheKey = 'attendance_insights_' + days;
    var cached = getCachedData(cacheKey, 10);
    if (cached) return cached;

    var users = getCached(SH.USERS, TTL_LOOKUP);
    var rows = getAttendanceRowsForAnalytics();
    var start = new Date();
    start.setHours(0, 0, 0, 0);
    start = new Date(start.getTime() - (days - 1) * 24 * 60 * 60 * 1000);
    var lateEntries = [];
    var earlyExits = [];
    var absenteeMap = {};
    var windows = getAttendanceWindowRows().filter(function(w) {
      return String(w.status || '').toLowerCase() === 'active';
    });

    rows.forEach(function(row) {
      var rowDateStr = normDate(row[6], t);
      var rowDate = rowDateStr ? new Date(rowDateStr + 'T00:00:00') : null;
      if (!rowDate || rowDate < start) return;
      var uid = String(row[1] || '').trim();
      var user = getUserById(uid) || { full_name: uid, email: '' };
      var matched = findMatchingWindowForRow(row, windows, t);
      if (matched && isLateEntry(row, matched, t) && String(row[4] || '').trim()) {
        lateEntries.push({
          userId: uid,
          name: user.full_name,
          date: rowDateStr,
          entryTime: row[4] || '',
          locationId: row[10] || ''
        });
      }
      if (matched && isEarlyExit(row, matched, t) && String(row[5] || '').trim()) {
        earlyExits.push({
          userId: uid,
          name: user.full_name,
          date: rowDateStr,
          exitTime: row[5] || '',
          locationId: row[10] || ''
        });
      }
      absenteeMap[uid] = absenteeMap[uid] || {};
      absenteeMap[uid][rowDateStr] = true;
    });

    var frequentAbsentees = users.map(function(u) {
      var uid = String(u.user_id || '').trim();
      var missed = 0;
      for (var i = 0; i < days; i++) {
        var dt = new Date(start.getTime() + i * 24 * 60 * 60 * 1000);
        var d = Utilities.formatDate(dt, t, 'yyyy-MM-dd');
        if (!absenteeMap[uid] || !absenteeMap[uid][d]) missed += 1;
      }
      return { userId: uid, name: u.full_name, email: u.email, absentDays: missed };
    }).sort(function(a, b) { return b.absentDays - a.absentDays; }).slice(0, 10);

    var result = {
      success: true,
      days: days,
      lateEntries: lateEntries.slice(0, 20),
      earlyExits: earlyExits.slice(0, 20),
      frequentAbsentees: frequentAbsentees,
      updatedAt: new Date().toISOString()
    };
    setCachedData(cacheKey, result, 10);
    return result;
  } catch (err) {
    return { success: false, message: 'getAttendanceInsights: ' + err };
  }
}

function parseClockMinutes(value) {
  var parts = String(value || '').trim().split(':');
  if (parts.length < 2) return null;
  var h = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function findMatchingWindowForRow(row, windows, tzName) {
  var rowDate = normDate(row[6], tzName || tz());
  if (!rowDate) return null;
  var locId = String(row[10] || '').trim();
  if (!windows || !windows.length) return null;
  for (var i = 0; i < windows.length; i++) {
    var w = windows[i];
    var windowLoc = String(w.location_id || '').trim();
    if (windowLoc && locId && windowLoc !== locId) continue;
    var startTime = String(w.start_time || '').trim();
    var endTime = parseWindowEndTime(w);
    if (!startTime || !endTime) continue;
    return {
      startMin: parseClockMinutes(startTime),
      endMin: parseClockMinutes(endTime),
      startTime: startTime,
      endTime: endTime
    };
  }
  return null;
}

function isLateEntry(row, windowInfo) {
  if (!windowInfo) return false;
  var entryTime = String(row[4] || '').trim();
  var entryMin = parseClockMinutes(entryTime);
  if (entryMin === null || windowInfo.startMin === null) return false;
  return entryMin > (windowInfo.startMin + 5);
}

function isEarlyExit(row, windowInfo) {
  if (!windowInfo) return false;
  var exitTime = String(row[5] || '').trim();
  var exitMin = parseClockMinutes(exitTime);
  if (exitMin === null || windowInfo.endMin === null) return false;
  return exitMin < (windowInfo.endMin - 5);
}

// ============================================================
//  9. STUDENTS ROSTER
// ============================================================

function getStudents(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var studentRoleId = findRoleIdByName('student');
    var rows = getCached(SH.USERS, TTL_LOOKUP)
      .filter(function(r) {
        return normalizeRoleValue(r.role_id) === 'student' ||
               (studentRoleId && String(r.role_id) === studentRoleId);
      })
      .map(function(r) {
        return { userId: r.user_id, name: r.full_name, email: r.email,
                 department: r.department_id, mobile: r.mobile,
                 hasBio: !!r.biometric_code, hasDevice: !!r.device_identification };
      });
    return { success: true, students: rows, total: rows.length };
  } catch(err) { return { success: false, message: 'getStudents: ' + err }; }
}

// ============================================================
//  10. EXPORT ATTENDANCE
// ============================================================

function exportAttendance(b) {
  try {
    var auth = requireAuth(b, ['teacher', 'admin']);
    if (!auth.ok) return auth.error;
    var t = tz(), sheet = getSheet(SH.ATTENDANCE), lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, csv: '', rowCount: 0 };
    var data = sheet.getRange(2, 1, lastRow-1, attendanceReadColumnCount(sheet)).getValues();
    var header = ['attendance_id','user_id','full_name','type_attendance','entry_time','exit_time',
                  'attendance_date','login_method','latitude','longitude','attendance_location_id',
                  'address','distance_from_centre','status'];
    var lines = [header.join(',')];
    data.forEach(function(row) {
      if (b.userId && String(row[1]).trim() !== String(b.userId).trim()) return;
      if (b.date   && normDate(row[6], t) !== b.date) return;
      lines.push(row.map(function(c) { return '"' + String(c||'').replace(/"/g,'""') + '"'; }).join(','));
    });
    return { success: true, csv: lines.join('\n'), rowCount: lines.length - 1 };
  } catch(err) { return { success: false, message: 'exportAttendance: ' + err }; }
}

// ============================================================
//  11. BIOMETRIC & DEVICE
// ============================================================

function getBiometric(b) {
  try {
    if (!b.email)
      return { success: false, message: 'Email required' };

    var user = getUserByEmail(b.email);
    if (!user)
      return { success: false, message: 'No account found. Please register first.' };

    if (!user.biometric_code)
      return { success: false, message: 'No biometric registered. Please register first.' };

    return {
      success: true,
      credentialId: String(user.biometric_code || ''),
      userId: user.user_id,
      name: user.full_name
    };
  } catch (err) {
    return { success: false, message: 'getBiometric: ' + err };
  }
}

function saveBiometric(b) {
  try {
    if (!b.userId || !b.credentialId)
      return { success: false, message: 'userId and credentialId required' };

    var sheet = getSheet(SH.USERS);
    var last  = sheet.getLastRow();
    if (last < 2) return { success: false, message: 'User not found' };

    var uids  = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < uids.length; i++) {
      if (String(uids[i][0]) === String(b.userId)) {
        sheet.getRange(i + 2, 10).setValue(b.credentialId);
        invalidate(SH.USERS);
        return { success: true, message: 'Biometric registered successfully' };
      }
    }

    return { success: false, message: 'User not found' };
  } catch (err) {
    return { success: false, message: 'saveBiometric: ' + err };
  }
}

function checkBiometric(b) {
  try {
    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    return {
      success: true,
      hasBiometric: !!user.biometric_code
    };
  } catch (err) {
    return { success: false, message: 'checkBiometric: ' + err };
  }
}

function registerDevice(b) {
  try {
    if (!b.userId || !b.deviceId) return { success: false, message: 'userId and deviceId required' };
    var user = getUserById(b.userId);
    if (!user) return { success: false, message: 'User not found' };
    var stored = String(user.device_identification || '').trim();
    if (stored && stored !== String(b.deviceId).trim())
      return { success: false, alreadyBound: true, message: 'Account already bound to another device.' };
    var sheet = getSheet(SH.USERS), last = sheet.getLastRow();
    var uids = sheet.getRange(2, 1, last-1, 1).getValues();
    for (var i = 0; i < uids.length; i++) {
      if (String(uids[i][0]) === String(b.userId)) {
        sheet.getRange(i+2, 11).setValue(b.deviceId);
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
//  12. LOOKUP TABLES
// ============================================================

function getLookup(sheetName) {
  try {
    ensureCoreTenantData();
    return { success: true, data: getCached(sheetName, TTL_LOOKUP) };
  }
  catch(err) { return { success: false, message: 'getLookup: ' + err }; }
}

function ensureCoreTenantData() {
  var coreSheets = [SH.ROLES, SH.DEPARTMENTS, SH.ATT_TYPE, SH.ATT_LOCATIONS, SH.ATT_WINDOWS];
  var hasAnyData = coreSheets.some(function(name) {
    try {
      return getSheet(name).getLastRow() > 1;
    } catch (e) {
      return false;
    }
  });
  if (!hasAnyData) {
    setupSheets();
  }
}

function getUserLocMap(b) {
  try {
    var rows = getRows(getSheet(SH.USER_LOC_MAP));
    if (b.userId) rows = rows.filter(function(r) { return String(r.user_id) === String(b.userId); });
    return { success: true, data: rows };
  } catch(err) { return { success: false, message: 'getUserLocMap: ' + err }; }
}

// ============================================================
//  13. ADMIN WRITE — all use nextId() for integer PKs
// ============================================================

function addDepartment(b) {
  try {
    var auth = requireAuth(b, ['admin']);
    if (!auth.ok) return auth.error;
    if (!b.name) return { success: false, message: 'name required' };
    if (b.email && !isValidEmail(b.email))
      return { success: false, message: 'Valid department email required' };

    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var deptId = nextId(SH.DEPARTMENTS);   // ← integer (1, 2, 3…)
      getSheet(SH.DEPARTMENTS).appendRow([
        deptId,
        toCleanText(b.name, 120),
        toCleanText(b.inCharge, 120),
        toCleanText(b.email, 120),
        currentGuid()
      ]);
      invalidate(SH.DEPARTMENTS);
      appendAuditLog('add_department', auth.session.user_id, auth.session.role_name, '', { departmentId: deptId });
      return { success: true, departmentId: deptId, message: 'Department added' };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success: false, message: 'addDepartment: ' + err }; }
}

function addAttendanceLocation(b) {
  try {
    var auth = requireAuth(b, ['admin']);
    if (!auth.ok) return auth.error;
    if (!b.name || !b.latitude || !b.longitude)
      return { success: false, message: 'name, latitude and longitude required' };
    var lat = parseNumber(b.latitude), lng = parseNumber(b.longitude);
    var radius = parseInt(b.geofenceRadius || b.geofence_radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
    var address = toCleanText(b.address || '', 240);
    if (lat === null || lng === null)
      return { success: false, message: 'Valid latitude and longitude required' };

    var tenantId = currentGuid();
    var nearby = findNearbyAttendanceLocations(lat, lng, tenantId, b.reuseLocationId || '');
    if (String(b.reuseLocationId || '').trim()) {
      var reuseLoc = findLocationById(b.reuseLocationId, tenantId);
      if (!reuseLoc) return { success: false, message: 'Existing location not found for reuse' };
      return {
        success: true,
        reused: true,
        locationId: reuseLoc.attendance_location_id,
        message: 'Reused existing location'
      };
    }
    if (nearby.length && String(b.confirmDuplicate || '').toLowerCase() !== 'true') {
      return {
        success: false,
        duplicateWarning: true,
        message: 'A location already exists within 20 meters.',
        nearbyLocations: nearby,
        canReuse: true
      };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var locId = nextId(SH.ATT_LOCATIONS);  // integer
      getSheet(SH.ATT_LOCATIONS).appendRow([
        locId,
        toCleanText(b.name, 120),
        address,
        lat,
        lng,
        radius,
        tenantId
      ]);
      invalidate(SH.ATT_LOCATIONS);
      appendAuditLog('add_location', auth.session.user_id, auth.session.role_name, '', {
        locationId: locId,
        address: address,
        geofenceRadius: radius
      });
      return {
        success: true,
        locationId: locId,
        message: nearby.length ? 'Location added after confirmation' : 'Location added'
      };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success: false, message: 'addAttendanceLocation: ' + err }; }
}

function addCategoryLocationMap(b) {
  try {
    var auth = requireAuth(b, ['admin']);
    if (!auth.ok) return auth.error;
    var categoryId = toCleanText(b.categoryId || b.departmentId || '', 120);
    var subcategoryId = toCleanText(b.subcategoryId || '', 120);
    var locationId = String(b.locationId || b.attendanceLocationId || '').trim();
    if (!categoryId || !locationId) {
      return { success: false, message: 'categoryId and locationId required' };
    }
    var loc = findLocationById(locationId, currentGuid());
    if (!loc) return { success: false, message: 'attendance location not found' };

    var allowedDistance = parseInt(b.allowedDistance || b.allowed_distance || loc.geofence_radius || DEFAULT_RADIUS, 10) || DEFAULT_RADIUS;
    var lock = LockService.getScriptLock();
    lock.waitLock(6000);
    try {
      var mapId = nextId(SH.CATEGORY_LOC_MAP);
      getSheet(SH.CATEGORY_LOC_MAP).appendRow([
        mapId,
        categoryId,
        subcategoryId,
        parseInt(locationId, 10) || locationId,
        allowedDistance,
        currentGuid(),
        String(b.status || 'active')
      ]);
      invalidate(SH.CATEGORY_LOC_MAP);
      appendAuditLog('add_category_location_map', auth.session.user_id, auth.session.role_name, '', {
        categoryId: categoryId,
        subcategoryId: subcategoryId,
        locationId: locationId,
        allowedDistance: allowedDistance
      });
      return { success: true, mapId: mapId, message: 'Mapping added' };
    } finally { lock.releaseLock(); }
  } catch(err) { return { success: false, message: 'addCategoryLocationMap: ' + err }; }
}

function addUserLocMap(b) {
  return addCategoryLocationMap(b);
}

function getCategoryLocationMap(b) {
  try {
    var auth = requireAuth(b, ['admin']);
    if (!auth.ok) return auth.error;
    var rows = getRows(getSheet(SH.CATEGORY_LOC_MAP));
    if (b.categoryId || b.subcategoryId) {
      rows = rows.filter(function(r) {
        return (!b.categoryId || String(r.category_id) === String(b.categoryId)) &&
               (!b.subcategoryId || String(r.subcategory_id) === String(b.subcategoryId));
      });
    }
    return { success: true, data: rows };
  } catch(err) { return { success: false, message: 'getCategoryLocationMap: ' + err }; }
}

// ============================================================
//  14. SETUP — seeds all tables with numeric IDs
// ============================================================

function setupSheets() {
  try {
    var created = [];
    Object.keys(SH).forEach(function(key) { getSheet(SH[key]); created.push(SH[key]); });

    // Roles: id 1 = admin, 2 = teacher/faculty, 3 = student, 4 = employee
    ensureExactRows('Roles', [
      [1, 'admin'],
      [2, 'teacher'],
      [3, 'student'],
      [4, 'employee']
    ]);

    // Departments: id 1 = CSE
    ensureExactRows('Departments', [
      [1, 'CSE', 'Dr. Sunitha N R', 'nrsunitha@sit.ac.in']
    ]);

    // Attendance types: 1 = entry, 2 = exit
    ensureExactRows('AttendanceType', [
      [1, 'entry'],
      [2, 'exit']
    ]);

    // Attendance windows: numeric ids
    ensureExactRows('AttendanceWindows', [
      [1, 'Morning 1',   '09:00', 10, 'active', 1],
      [2, 'Morning 2',   '10:30', 10, 'active', 1],
      [3, 'Afternoon 1', '14:00', 10, 'active', 1],
      [4, 'Afternoon 2', '15:30', 10, 'active', 1]
    ]);

    // Attendance locations: 1, 2, 3
    var locSheet = getSheet(SH.ATT_LOCATIONS);
    if (locSheet.getLastRow() < 2) {
      var fence = getTenantGeofence(currentGuid());
      var label = fence.label || 'Campus';
      locSheet.appendRow([1, label + ' Campus Main Block', 'Main entry point', fence.latitude, fence.longitude, fence.radius, currentGuid()]);
      locSheet.appendRow([2, label + ' Campus CS Lab Block', 'Computer science block', fence.latitude + 0.00012, fence.longitude + 0.00012, fence.radius, currentGuid()]);
      locSheet.appendRow([3, label + ' Campus Seminar Hall', 'Seminar hall', fence.latitude - 0.00012, fence.longitude - 0.00012, fence.radius, currentGuid()]);
    }

    return { success: true, message: 'All sheets created and seeded with numeric IDs.', sheets: created };
  } catch(err) { return { success: false, message: 'setupSheets: ' + err }; }
}

// ============================================================
//  15. RESET USERS SHEET (admin utility)
// ============================================================

function resetUsersSheet() {
  try {
    ensureExactRows(SH.USERS, []);
    ensureExactRows(SH.USER_INDEX, []);
    ensureExactRows(SH.USER_LOC_MAP, []);
    ensureExactRows(SH.CATEGORY_LOC_MAP, []);
    invalidate(SH.USERS);
    invalidate(SH.USER_INDEX);
    invalidate(SH.USER_LOC_MAP);
    invalidate(SH.CATEGORY_LOC_MAP);
    return { success: true, message: 'Users, UserIndex, UserLocMap and CategoryLocationMap cleared.' };
  } catch(err) { return { success: false, message: 'resetUsersSheet: ' + err }; }
}

// ============================================================
//  16. DEBUG
// ============================================================

function debugInfo() {
  try {
    var fence = getTenantGeofence(currentGuid());
    var sheets = ss().getSheets().map(function(s) {
      return { name: s.getName(), dataRows: Math.max(0, s.getLastRow()-1) };
    });
    return { success: true, spreadsheet: ss().getName(), sheets: sheets,
             geofence: fence.radius + 'm around (' + fence.latitude + ', ' + fence.longitude + ')' };
  } catch(err) { return { success: false, message: 'debug: ' + err }; }
}



