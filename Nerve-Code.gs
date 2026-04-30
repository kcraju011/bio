var NERVE_SPREADSHEET_ID = '1g43D4SH7DnYAh05trU5t5xhfyVxP43unO4knX8WdA9k';
var TENANT_SHEET_NAME = 'Tenants';

var TENANT_HEADERS = [
  'guid',
  'alias',
  'institution_name',
  'org_type',
  'city',
  'spreadsheet_url',
  'api_url',
  'application_id',
  'application_name',
  'application_description',
  'logo_url',
  'website',
  'status',
  'created_at',
  'updated_at'
];

var DEFAULT_TENANTS = [
  {
    guid: '1',
    alias: 'SIT',
    institution_name: 'SIT',
    org_type: 'college',
    city: 'Tumakuru',
    spreadsheet_url: 'https://docs.google.com/spreadsheets/d/1c5ZGKxr-ZNxbao6L6D87N8JNSvPNf2Il9Xnl8ha04f8/edit?gid=0#gid=0',
    api_url: 'https://script.google.com/macros/s/AKfycbyfgsjU607novJJlhwZfMdlSreCGP7OLaaj6ztikWQb4VawkisqPGLwdqDkDuqYfjQlZw/exec',
    application_id: '102',
    application_name: 'Attendance monitoring',
    application_description: 'Biometric attendance for students, teachers and employees',
    logo_url: 'https://web.sit.ac.in/wp-content/uploads/2025/03/SIT-Logo-1.png',
    website: '',
    status: 'active'
  },
  {
    guid: '2',
    alias: 'SSIT',
    institution_name: 'SSIT',
    org_type: 'college',
    city: 'Tumakuru',
    spreadsheet_url: 'https://docs.google.com/spreadsheets/d/1yGS6eyC6NTqwu6dllTYALe_Sc_0hn4_furgU-Wq6bH4/edit?gid=0#gid=0',
    api_url: 'https://script.google.com/macros/s/AKfycbzQ1Abzar9Ydny0mZyw4JMMOrxs868Flpdb7vHdHhltbYHdRIVQYl-aTsUQ6uP0zei_/exec',
    application_id: '103',
    application_name: 'Attendance monitoring',
    application_description: 'Biometric attendance for students, teachers and employees',
    logo_url: 'https://ssit.edu.in/img/ssit-logo.png',
    website: '',
    status: 'active'
  }
];

var TENANT_GUID_ALIASES = {
  '1': 'sit',
  '2': 'ssit'
};

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

function nerveSs() {
  return SpreadsheetApp.openById(NERVE_SPREADSHEET_ID);
}

function nerveSheet() {
  var sheet = nerveSs().getSheetByName(TENANT_SHEET_NAME);
  if (!sheet) {
    sheet = nerveSs().insertSheet(TENANT_SHEET_NAME);
    sheet.appendRow(TENANT_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureNerveRegistry() {
  var sheet = nerveSheet();
  var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  if (rows.length) return;
  var values = DEFAULT_TENANTS.map(function(t) {
    return [
      t.guid, t.alias, t.institution_name, t.org_type, t.city, t.spreadsheet_url, t.api_url,
      t.application_id, t.application_name, t.application_description, t.logo_url, t.website,
      t.status, new Date().toISOString(), new Date().toISOString()
    ];
  });
  sheet.getRange(2, 1, values.length, TENANT_HEADERS.length).setValues(values);
}

function rowsToObjects(values) {
  if (!values || !values.length) return [];
  var headers = values[0];
  return values.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(header, idx) {
      obj[String(header || '').trim()] = row[idx];
    });
    return obj;
  });
}

function getTenants() {
  ensureNerveRegistry();
  return rowsToObjects(nerveSheet().getDataRange().getValues());
}

function normalizeGuid(value) {
  return String(value || '').trim().toLowerCase();
}

function canonicalTenantGuid(target) {
  var raw = normalizeGuid(target);
  if (raw === '1' || raw === 'sit') return '1';
  if (raw === '2' || raw === 'ssit') return '2';
  return raw;
}

function getTenantByGuid(guid) {
  var target = normalizeGuid(guid);
  var canonical = canonicalTenantGuid(target);
  var aliasTarget = TENANT_GUID_ALIASES[canonical] || TENANT_GUID_ALIASES[target] || '';
  if (!target) return null;
  var tenants = getTenants();
  for (var i = 0; i < tenants.length; i++) {
    var row = tenants[i];
    var rowGuid = normalizeGuid(row.guid);
    var alias = normalizeGuid(row.alias);
    if (rowGuid === target || alias === target || rowGuid === canonical || alias === canonical || alias === aliasTarget) return row;
  }
  return null;
}

function formatTenantResponse(row, requestedGuid) {
  if (!row) return { success: false, message: 'Tenant not found' };
  return {
    success: true,
    guid: String(requestedGuid || row.guid || ''),
    institution: {
      name: String(row.institution_name || ''),
      city: String(row.city || ''),
      logoUrl: String(row.logo_url || ''),
      website: String(row.website || ''),
      address: String(row.city || '')
    },
    application: {
      id: String(row.application_id || ''),
      name: String(row.application_name || ''),
      description: String(row.application_description || '')
    },
    orgType: String(row.org_type || ''),
    spreadsheetUrl: String(row.spreadsheet_url || ''),
    apiUrl: String(row.api_url || ''),
    status: String(row.status || 'active')
  };
}

function getApplicationFromGuid(guid) {
  return formatTenantResponse(getTenantByGuid(guid), guid);
}

function doGet(e) {
  try {
    var p = e && e.parameter ? e.parameter : {};
    var callback = p.callback || '';
    if (p.setup === '1') {
      ensureNerveRegistry();
      return respondOut({ success: true, message: 'Nerve registry ready' }, callback);
    }
    if (p.listTenants === '1') {
      return respondOut({ success: true, tenants: getTenants() }, callback);
    }
    if (p.getApplicationFromGuid) {
      return respondOut(getApplicationFromGuid(p.getApplicationFromGuid), callback);
    }
    if (p.guid) {
      return respondOut(getApplicationFromGuid(p.guid), callback);
    }
    return respondOut({
      success: true,
      message: 'Nerve registry is running',
      endpoints: ['?getApplicationFromGuid=GUID', '?listTenants=1']
    }, callback);
  } catch (err) {
    return jsonOut({ success: false, message: 'Nerve doGet: ' + err });
  }
}

function doPost(e) {
  try {
    var body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    if (body.action === 'setup') {
      ensureNerveRegistry();
      return jsonOut({ success: true, message: 'Nerve registry ready' });
    }
    if (body.action === 'getApplicationFromGuid') {
      return jsonOut(getApplicationFromGuid(body.guid));
    }
    return jsonOut({ success: false, message: 'Unknown action' });
  } catch (err) {
    return jsonOut({ success: false, message: 'Nerve doPost: ' + err });
  }
}

function setupNerveRegistry() {
  var sheet = nerveSheet();
  sheet.clearContents();
  sheet.appendRow(TENANT_HEADERS);
  sheet.setFrozenRows(1);
  var values = DEFAULT_TENANTS.map(function(t) {
    return [
      t.guid, t.alias, t.institution_name, t.org_type, t.city, t.spreadsheet_url, t.api_url,
      t.application_id, t.application_name, t.application_description, t.logo_url, t.website,
      t.status, new Date().toISOString(), new Date().toISOString()
    ];
  });
  sheet.getRange(2, 1, values.length, TENANT_HEADERS.length).setValues(values);
  return { success: true, message: 'Nerve tenants seeded' };
}
