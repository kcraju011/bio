// ============================================================
//  BioAttend Auth – Google Apps Script Backend
//  Spreadsheet: https://docs.google.com/spreadsheets/d/1rruBrVa-CPJfE8-2_ubYJWMG2hU8I6pYjzpWGNCPzP8
//  Deploy: Web App → Execute as Me → Access: Anyone
// ============================================================

var SHEET_USERS = 'Users';

var USER_HEADERS = [
  'UserID', 'FullName', 'Email', 'PasswordHash',
  'DOB', 'Mobile', 'Institution', 'Department',
  'Role', 'BiometricCredentialId', 'DeviceId', 'CreatedAt'
];

// ── JSON output ───────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var param = (e && e.parameter) ? e.parameter : {};
    if (!param.data) return jsonOut({ status: 'BioAttend Auth API v1', time: new Date().toString() });
    var body = JSON.parse(decodeURIComponent(param.data));
    return jsonOut(route(body));
  } catch (err) {
    return jsonOut({ success: false, message: 'Error: ' + err.toString() });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return jsonOut(route(body));
  } catch (err) {
    return jsonOut({ success: false, message: 'Error: ' + err.toString() });
  }
}

function route(body) {
  switch (body.action) {
    case 'register':         return registerUser(body);
    case 'signIn':           return signInUser(body);
    case 'saveBiometric':    return saveBiometric(body);
    case 'getBiometric':     return getBiometric(body);
    case 'registerDevice':   return registerDevice(body);
    case 'checkDevice':      return checkDevice(body);
    case 'getUser':          return getUser(body);
    case 'debug':            return debugInfo();
    default: return { success: false, message: 'Unknown action: ' + body.action };
  }
}

// ── Sheet helpers ─────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(USER_HEADERS);
    sheet.getRange(1, 1, 1, USER_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1a2a52')
      .setFontColor('#ffffff');
  } else {
    // Auto-add missing columns
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    USER_HEADERS.forEach(function(col) {
      if (existing.indexOf(col) === -1) {
        var c = sheet.getLastColumn() + 1;
        sheet.getRange(1, c).setValue(col)
          .setFontWeight('bold')
          .setBackground('#1a2a52')
          .setFontColor('#ffffff');
      }
    });
  }
  return sheet;
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

function generateId(prefix) {
  return (prefix || 'id') + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
}

function hashPassword(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ── 1. Register ───────────────────────────────────────────────
function registerUser(body) {
  try {
    if (!body.name || !body.email || !body.password)
      return { success: false, message: 'Name, email and password are required' };
    if (body.password.length < 6)
      return { success: false, message: 'Password must be at least 6 characters' };

    var sheet = getSheet(SHEET_USERS);
    var rows = getRows(sheet);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].Email).toLowerCase() === String(body.email).toLowerCase())
        return { success: false, message: 'Email already registered' };
    }

    var userId = generateId('u');
    var role = body.role || 'student';
    var now = new Date().toISOString();

    sheet.appendRow([
      userId,
      body.name,
      body.email,
      hashPassword(body.password),
      body.dob || '',
      body.mobile || '',
      body.institution || '',
      body.department || '',
      role,
      '', // BiometricCredentialId
      '', // DeviceId
      now  // CreatedAt
    ]);

    return {
      success: true,
      userId: userId,
      name: body.name,
      email: body.email,
      role: role,
      message: 'Account created successfully'
    };
  } catch (err) {
    return { success: false, message: 'register: ' + err.toString() };
  }
}

// ── 2. Sign In ────────────────────────────────────────────────
function signInUser(body) {
  try {
    if (!body.email || !body.password)
      return { success: false, message: 'Email and password are required' };

    var sheet = getSheet(SHEET_USERS);
    var rows = getRows(sheet);
    var hash = hashPassword(body.password);

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r.Email).toLowerCase() === String(body.email).toLowerCase() && r.PasswordHash === hash) {
        return {
          success: true,
          userId: r.UserID,
          name: r.FullName,
          email: r.Email,
          role: String(r.Role || 'student').trim(),
          department: r.Department || '',
          institution: r.Institution || '',
          hasBiometric: !!r.BiometricCredentialId,
          hasDevice: !!r.DeviceId
        };
      }
    }
    return { success: false, message: 'Invalid email or password' };
  } catch (err) {
    return { success: false, message: 'signIn: ' + err.toString() };
  }
}

// ── 3. Get User ───────────────────────────────────────────────
function getUser(body) {
  try {
    if (!body.userId && !body.email)
      return { success: false, message: 'userId or email required' };

    var rows = getRows(getSheet(SHEET_USERS));
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var match = (body.userId && r.UserID === body.userId) ||
                  (body.email && String(r.Email).toLowerCase() === String(body.email).toLowerCase());
      if (match) {
        return {
          success: true,
          userId: r.UserID,
          name: r.FullName,
          email: r.Email,
          role: r.Role || 'student',
          department: r.Department || '',
          institution: r.Institution || '',
          dob: r.DOB || '',
          mobile: r.Mobile || '',
          hasBiometric: !!r.BiometricCredentialId,
          hasDevice: !!r.DeviceId,
          createdAt: r.CreatedAt || ''
        };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (err) {
    return { success: false, message: 'getUser: ' + err.toString() };
  }
}

// ── 4. Save Biometric ─────────────────────────────────────────
function saveBiometric(body) {
  try {
    if (!body.userId || !body.credentialId)
      return { success: false, message: 'userId and credentialId required' };

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = headers.indexOf('BiometricCredentialId') + 1;
    if (col < 1) return { success: false, message: 'BiometricCredentialId column missing' };

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.userId) {
        sheet.getRange(i + 1, col).setValue(body.credentialId);
        return { success: true, message: 'Biometric saved' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (err) {
    return { success: false, message: 'saveBiometric: ' + err.toString() };
  }
}

// ── 5. Get Biometric ─────────────────────────────────────────
function getBiometric(body) {
  try {
    if (!body.email && !body.userId)
      return { success: false, message: 'email or userId required' };

    var rows = getRows(getSheet(SHEET_USERS));
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var match = (body.email && String(r.Email).toLowerCase() === String(body.email).toLowerCase()) ||
                  (body.userId && r.UserID === body.userId);
      if (match) {
        if (!r.BiometricCredentialId)
          return { success: false, message: 'No biometric registered for this account' };
        return { success: true, credentialId: r.BiometricCredentialId, userId: r.UserID, name: r.FullName };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (err) {
    return { success: false, message: 'getBiometric: ' + err.toString() };
  }
}

// ── 6. Register Device ────────────────────────────────────────
function registerDevice(body) {
  try {
    if (!body.userId || !body.deviceId)
      return { success: false, message: 'userId and deviceId required' };

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = headers.indexOf('DeviceId') + 1;
    if (col < 1) return { success: false, message: 'DeviceId column not found' };

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.userId) {
        var existing = String(data[i][col - 1] || '').trim();
        if (existing && existing !== String(body.deviceId).trim())
          return { success: false, alreadyBound: true, message: 'Account already bound to another device' };
        sheet.getRange(i + 1, col).setValue(body.deviceId);
        return { success: true, alreadyBound: false, message: existing ? 'Device confirmed' : 'Device registered' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (err) {
    return { success: false, message: 'registerDevice: ' + err.toString() };
  }
}

// ── 7. Check Device ───────────────────────────────────────────
function checkDevice(body) {
  try {
    if (!body.userId || !body.deviceId)
      return { success: false, message: 'userId and deviceId required' };

    var rows = getRows(getSheet(SHEET_USERS));
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].UserID === body.userId) {
        var stored = String(rows[i].DeviceId || '').trim();
        if (!stored) return { success: true, status: 'unbound' };
        if (stored === String(body.deviceId).trim()) return { success: true, status: 'match' };
        return { success: false, status: 'mismatch', message: 'Account registered to a different device' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (err) {
    return { success: false, message: 'checkDevice: ' + err.toString() };
  }
}

// ── 8. Debug ──────────────────────────────────────────────────
function debugInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheet(SHEET_USERS);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowCount = Math.max(0, sheet.getLastRow() - 1);
    return {
      success: true,
      api: 'BioAttend Auth v1',
      spreadsheet: ss.getName(),
      userHeaders: headers,
      userCount: rowCount
    };
  } catch (err) {
    return { success: false, message: 'debug: ' + err.toString() };
  }
}