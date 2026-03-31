// ============================================================
//  BioAttend – Google Apps Script Backend  (Code.gs)
//  Deploy as: Web App → Execute as: Me → Access: Anyone
// ============================================================

// ── Sheet names ──────────────────────────────────────────────
var SHEET_USERS      = 'Users';
var SHEET_ATTENDANCE = 'Attendance';

// ── CORS helper ──────────────────────────────────────────────
function setCors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── Entry points ─────────────────────────────────────────────
function doOptions() {
  return setCors(ContentService.createTextOutput(''));
}

function doPost(e) {
  try {
    var body    = JSON.parse(e.postData.contents);
    var action  = body.action;
    var result  = {};

    if      (action === 'register')       result = registerUser(body);
    else if (action === 'signIn')         result = signInUser(body);
    else if (action === 'markAttendance') result = markAttendance(body);
    else if (action === 'saveBiometric')  result = saveBiometric(body);
    else if (action === 'getBiometric')   result = getBiometric(body);
    else result = { success: false, message: 'Unknown action' };

    return setCors(ContentService.createTextOutput(JSON.stringify(result)));
  } catch (err) {
    var errPayload = { success: false, message: err.toString() };
    return setCors(ContentService.createTextOutput(JSON.stringify(errPayload)));
  }
}

// Also support GET for simple testing
function doGet(e) {
  return setCors(ContentService.createTextOutput(JSON.stringify({ status: 'BioAttend API running' })));
}

// ── Spreadsheet helpers ──────────────────────────────────────
function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_USERS) {
      sheet.appendRow([
        'UserID', 'FullName', 'Email', 'PasswordHash',
        'DOB', 'Mobile', 'Institution', 'Department',
        'MarkFromAnywhere', 'BiometricCredentialId', 'CreatedAt'
      ]);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    }
    if (name === SHEET_ATTENDANCE) {
      sheet.appendRow(['AttendanceID', 'UserID', 'FullName', 'Email', 'Timestamp', 'Date', 'Time', 'Method']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1a2a52').setFontColor('#ffffff');
    }
  }
  return sheet;
}

function generateId() {
  return 'u_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 6);
}

function hashPassword(pw) {
  // Simple hash using Utilities.computeDigest (SHA-256 equivalent via MD5 for Apps Script)
  // For production, consider a stronger approach or use an external hashing service.
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ── 1. Register ──────────────────────────────────────────────
function registerUser(body) {
  var sheet = getSheet(SHEET_USERS);
  var data  = sheet.getDataRange().getValues();

  // Check duplicate email (skip header row at index 0)
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === body.email) {
      return { success: false, message: 'Email already registered' };
    }
  }

  var userId = generateId();
  var now    = new Date();

  sheet.appendRow([
    userId,
    body.name        || '',
    body.email       || '',
    hashPassword(body.password || ''),
    body.dob         || '',
    body.mobile      || '',
    body.institution || '',
    body.department  || '',
    body.anywhere ? 'YES' : 'NO',
    '',          // BiometricCredentialId – filled later
    now.toISOString()
  ]);

  return { success: true, userId: userId, message: 'Account created' };
}

// ── 2. Sign In ───────────────────────────────────────────────
function signInUser(body) {
  var sheet = getSheet(SHEET_USERS);
  var data  = sheet.getDataRange().getValues();
  var hash  = hashPassword(body.password || '');

  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === body.email && data[i][3] === hash) {
      return { success: true, userId: data[i][0], name: data[i][1] };
    }
  }
  return { success: false, message: 'Invalid email or password' };
}

// ── 3. Mark Attendance ───────────────────────────────────────
function markAttendance(body) {
  var usersSheet = getSheet(SHEET_USERS);
  var attSheet   = getSheet(SHEET_ATTENDANCE);
  var userData   = usersSheet.getDataRange().getValues();

  var userName  = '';
  var userEmail = '';

  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0] === body.userId) {
      userName  = userData[i][1];
      userEmail = userData[i][2];
      break;
    }
  }

  if (!userName) return { success: false, message: 'User not found' };

  var now       = new Date();
  var attId     = 'a_' + now.getTime();
  var dateStr   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var timeStr   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  var method    = body.method || 'password';

  attSheet.appendRow([attId, body.userId, userName, userEmail, now.toISOString(), dateStr, timeStr, method]);

  return { success: true, message: 'Attendance marked at ' + timeStr };
}

// ── 4. Save Biometric Credential ID ─────────────────────────
function saveBiometric(body) {
  var sheet = getSheet(SHEET_USERS);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === body.userId) {
      sheet.getRange(i + 1, 10).setValue(body.credentialId); // col 10 = BiometricCredentialId
      return { success: true };
    }
  }
  return { success: false, message: 'User not found' };
}

// ── 5. Get Biometric Credential ID (for sign-in) ─────────────
function getBiometric(body) {
  var sheet = getSheet(SHEET_USERS);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === body.email) {
      var credId = data[i][9];
      if (!credId) return { success: false, message: 'No biometric registered for this account' };
      return { success: true, credentialId: credId, userId: data[i][0] };
    }
  }
  return { success: false, message: 'User not found' };
}
