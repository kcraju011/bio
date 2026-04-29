function switchMain(tab) {
  ['signin','register'].forEach(t => {
    const pane = document.getElementById('pane-'+t);
    const tabBtn = document.getElementById('mtab-'+t);
    if (pane) pane.classList.toggle('active', t===tab);
    if (tabBtn) tabBtn.classList.toggle('active', t===tab);
  });
  if (tab === 'register') {
    resetRegisterFlow();
    loadRegisterLookups();
    applyTenantToRegistration();
  }
}

let liveRefreshInFlight = false;
let liveRefreshQueued = false;
let liveRetryDelay = 3000;
let liveLastActivityAt = Date.now();
let liveLastRequestAt = 0;
let liveMap = null;
let liveMapMarkers = {};
let analyticsCharts = { daily: null, weekly: null };

function markLiveActivity() {
  liveLastActivityAt = Date.now();
}

['mousemove', 'keydown', 'touchstart', 'scroll', 'focus'].forEach(function(evt) {
  document.addEventListener(evt, markLiveActivity, { passive: true, capture: true });
});

function isTeacherDashboardVisible() {
  const dash = document.getElementById('t-dashboard');
  return !!dash && dash.style.display !== 'none';
}

function isLiveTabActive() {
  return !!document.getElementById('sp-live')?.classList.contains('active');
}

function getLivePollDelay() {
  if (document.visibilityState !== 'visible') return null;
  const recentActivity = Date.now() - liveLastActivityAt < 15000;
  return recentActivity ? 3000 : 10000;
}

async function livePollTick() {
  livePollTimer = null;
  if (!isTeacherDashboardVisible() || !isLiveTabActive() || document.visibilityState !== 'visible') return;
  await refreshLive(false, true);
  if (livePollTimer) return;
  const delay = getLivePollDelay();
  if (delay !== null) {
    livePollTimer = setTimeout(livePollTick, delay);
  }
}

function switchSub(tab) {
  ['session','live','history','students','analytics'].forEach(t => {
    document.getElementById('sp-'+t).classList.toggle('active', t===tab);
    document.getElementById('stab-'+t).classList.toggle('active', t===tab);
  });
  if (tab==='live')     { refreshLive(true); startLivePolling(); }
  else                  stopLivePolling();
  if (tab==='history')  { document.getElementById('hist-date').value = new Date().toISOString().slice(0,10); loadHistory(); }
  if (tab==='students') loadStudents();
  if (tab==='analytics') loadAnalytics(true);
}

function startLivePolling() {
  if (livePollTimer || !isTeacherDashboardVisible() || !isLiveTabActive()) return;
  const delay = getLivePollDelay();
  if (delay === null) return;
  livePollTimer = setTimeout(livePollTick, delay);
}

function stopLivePolling() {
  if (livePollTimer) {
    clearTimeout(livePollTimer);
    livePollTimer = null;
  }
  liveRefreshInFlight = false;
  liveRefreshQueued = false;
  liveRetryDelay = 3000;
}

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && isTeacherDashboardVisible() && isLiveTabActive()) {
    startLivePolling();
    refreshLive();
  } else if (document.visibilityState !== 'visible') {
    stopLivePolling();
  }
});

function switchAdmin(tab) {
  ['dept','loc','ulmap','atttype'].forEach(t => {
    document.getElementById('ap-'+t).classList.toggle('active', t===tab);
    document.getElementById('atab-'+t).classList.toggle('active', t===tab);
  });
  if (tab==='dept')   loadDepts();
  if (tab==='loc')    loadLocs();
}

// â”€â”€ Device fingerprint â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async function getDeviceId() {
  if (deviceId) return deviceId;
  try { const s = localStorage.getItem('ba_did'); if(s){deviceId=s;return s;} } catch(e){}
  const cv = document.createElement('canvas'), c = cv.getContext('2d');
  c.fillText('BioAttend',2,2);
  const raw = [navigator.userAgent,navigator.language,screen.width+'x'+screen.height,
    new Date().getTimezoneOffset(),cv.toDataURL().slice(-40)].join('|');
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(raw));
  deviceId = Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('').slice(0,16);
  try { localStorage.setItem('ba_did', deviceId); } catch(e) {}
  return deviceId;
}

// â”€â”€ GPS (fine, watchPosition) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function getLocation() {
  return new Promise(resolve => {
    if (!navigator.geolocation) return resolve({latitude:'',longitude:'',accuracy:null,denied:false});
    let best=null, wid=null, done=false;
    const finish = r => {
      if(done)return; done=true;
      if(wid!==null)try{navigator.geolocation.clearWatch(wid);}catch(e){}
      resolve(r||{latitude:'',longitude:'',accuracy:null,denied:false});
    };
    wid = navigator.geolocation.watchPosition(
      pos => {
        const r={latitude:pos.coords.latitude,longitude:pos.coords.longitude,accuracy:pos.coords.accuracy,denied:false,address:''};
        if(!best||r.accuracy<best.accuracy)best=r;
        if(r.accuracy<=50)finish(best);
      },
      err => {
        if(err.code===1)finish({latitude:'',longitude:'',accuracy:null,denied:true});
        else finish(best||{latitude:'',longitude:'',accuracy:null,denied:false});
      },
      {enableHighAccuracy:true,timeout:10000,maximumAge:0}
    );
    setTimeout(()=>finish(best),10000);
  });
}

async function getLocationWithAddress() {
  const loc = await getLocation();
  if(loc.denied){showLocBar('fail','Location blocked â€” allow in browser settings and retry');return loc;}
  if(!loc.latitude)return loc;
  try {
    const r = await fetch(`https://nominatim.openstreetmap.org/reverse?lat=${loc.latitude}&lon=${loc.longitude}&format=json`,{headers:{'Accept-Language':'en'}});
    const d = await r.json();
    loc.address = d.display_name ? d.display_name.split(',').slice(0,3).join(',').trim() : '';
  } catch(e){loc.address='';}
  return loc;
}

function showLocBar(state,msg,accuracy){
  const el=document.getElementById('loc-status-bar');if(!el)return;
  const icons={getting:'ðŸ“¡',ok:'ðŸ“',fail:'âš ï¸'};
  const acc=accuracy?` <span style="opacity:.6;font-size:10px">Â±${Math.round(accuracy)}m</span>`:'';
  el.innerHTML=`<div class="loc-bar ${state}">${icons[state]||''} ${msg}${acc}</div>`;
}

function startTracking(userId) {
  if (!userId || !navigator.geolocation) return;
  stopTracking();
  window._trackingUserId = String(userId);
  window._trackingLastSentAt = 0;
  try { sessionStorage.setItem('ba_tracking_user', String(userId)); } catch(e) {}
  window._trackingWatchId = navigator.geolocation.watchPosition(
    pos => trackLocation(pos),
    err => console.error('Tracking GPS error', err),
    { enableHighAccuracy: true, maximumAge: 0, timeout: 10000 }
  );
}

function stopTracking() {
  if (window._trackingWatchId != null) {
    try { navigator.geolocation.clearWatch(window._trackingWatchId); } catch(e) {}
    window._trackingWatchId = null;
  }
  if (window._trackingInterval) {
    clearInterval(window._trackingInterval);
    window._trackingInterval = null;
  }
  window._trackingUserId = null;
  window._trackingLastSentAt = 0;
  try { sessionStorage.removeItem('ba_tracking_user'); } catch(e) {}
}

async function trackLocation(pos) {
  if (!window._trackingUserId || !pos?.coords) return;
  const now = Date.now();
  if (window._trackingLastSentAt && (now - window._trackingLastSentAt) < 30000) return;
  window._trackingLastSentAt = now;
  const payload = {
    action: 'trackLocation',
    user_id: window._trackingUserId,
    latitude: pos.coords.latitude,
    longitude: pos.coords.longitude
  };
  try {
    const res = await api(payload);
    if (res?.exitMarked) {
      stopTracking();
      toast('Auto exit recorded','success');
      restoreSignInForm();
    }
  } catch (e) {
    console.error('Tracking error', e);
  }
}

// â”€â”€ Attendance card (shown after marking) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function showAttendanceCard(data, userId) {
  markedUserId = userId;
  const el = document.getElementById('att-success-card');
  el.style.display = 'block';
  el.innerHTML = `
    <div class="att-card">
      <div class="att-name">âœ“ Attendance Marked</div>
      <div class="att-row"><span>full_name</span><span class="att-val">${data.name||''}</span></div>
      <div class="att-row"><span>attendance_date</span><span class="att-val">${data.date||''}</span></div>
      <div class="att-row"><span>entry_time</span><span class="att-val">${data.time||''}</span></div>
      <div class="att-row"><span>latitude</span><span class="att-val">${data.latitude||'not captured'}</span></div>
      <div class="att-row"><span>longitude</span><span class="att-val">${data.longitude||'not captured'}</span></div>
      <div class="att-row"><span>address</span><span class="att-val">${data.location||'not captured'}</span></div>
      <div class="att-row"><span>distance_from_centre</span><span class="att-val">${data.distanceFromCentre||'â€”'} m</span></div>
      <div class="att-row"><span>login_method</span><span class="att-val">${data.method||'biometric'}</span></div>
      <div class="att-row"><span>type_attendance</span><span class="att-val">entry</span></div>
    </div>
    <button class="btn-exit" id="btn-exit" onclick="handleExit()">
      <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9"/></svg>
      Mark Exit (leaving classroom)
    </button>`;
  // Hide form
  const ef = document.getElementById('si-email');
  if(ef&&ef.closest('.field'))ef.closest('.field').style.display='none';
  const ib = document.querySelector('#pane-signin .info-box');
  if(ib)ib.style.display='none';
  const bb = document.getElementById('btn-bio-signin');
  if(bb)bb.style.display='none';
  try{
    sessionStorage.setItem('ba_uid',   userId);
    sessionStorage.setItem('ba_name',  data.name||'');
    sessionStorage.setItem('ba_date',  data.date||'');
    sessionStorage.setItem('ba_time',  data.time||'');
    sessionStorage.setItem('ba_lat',   data.latitude||'');
    sessionStorage.setItem('ba_lng',   data.longitude||'');
    sessionStorage.setItem('ba_loc',   data.location||'');
    sessionStorage.setItem('ba_meth',  data.method||'biometric');
    sessionStorage.setItem('ba_dist',  data.distanceFromCentre||'');
  }catch(e){}
  startTracking(userId);
}

function restoreSignInForm() {
  stopTracking();
  const card=document.getElementById('att-success-card');if(card)card.style.display='none';
  const lb=document.getElementById('loc-status-bar');if(lb)lb.innerHTML='';
  const ef=document.getElementById('si-email');
  if(ef&&ef.closest('.field'))ef.closest('.field').style.display='';
  const ib=document.querySelector('#pane-signin .info-box');if(ib)ib.style.display='';
  const bb=document.getElementById('btn-bio-signin');if(bb){bb.style.display='';bb.disabled=false;}
  const pb=document.getElementById('btn-pass-signin');if(pb){pb.style.display='';pb.disabled=false;}
  if(ef)ef.value='';
  const pw=document.getElementById('si-password');if(pw)pw.value='';
  markedUserId=null;
  try{['ba_uid','ba_name','ba_date','ba_time','ba_lat','ba_lng','ba_loc','ba_meth','ba_dist'].forEach(k=>sessionStorage.removeItem(k));}catch(e){}
}

function base64UrlToUint8Array(value) {
  const padded = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  const withPad = padded + '='.repeat((4 - (padded.length % 4)) % 4);
  const raw = atob(withPad);
  return Uint8Array.from(raw, c => c.charCodeAt(0));
}

function bufferToBase64Url(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer || []);
  return btoa(String.fromCharCode(...bytes))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// BIOMETRIC SIGN IN
// Sends: getBiometric â†’ markEntry â†’ Attendance + LocationMonitor sheets
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function isValidRegistrationName(value) {
  return /^[A-Za-z][A-Za-z .'-]{1,79}$/.test(value);
}

function isValidRegistrationEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function isValidRegistrationMobile(value) {
  return /^[0-9+\-\s]{7,20}$/.test(value);
}

function isValidOrganizationName(value) {
  return /^[A-Za-z0-9][A-Za-z0-9 .,&()'/-]{1,119}$/.test(value);
}

function isValidDepartmentValue(value) {
  return /^[A-Za-z0-9_. -]{2,120}$/.test(String(value || '').trim());
}

function syncOrganizationName() {
  const orgEl = document.getElementById('r-institute');
  if (!orgEl) return;
  if (!orgEl.value) orgEl.value = '';
}

const COLLEGE_ROLE_OPTIONS = [
  { value: '1', label: 'Admin', roleKey: 'admin' },
  { value: '2', label: 'Faculty', roleKey: 'faculty' },
  { value: '3', label: 'Student', roleKey: 'student' },
  { value: '4', label: 'Employee', roleKey: 'employee' }
];

const REGISTER_BRANCH_OPTIONS = [
  { value: 'main-campus', label: 'Main Campus' },
  { value: 'north-campus', label: 'North Campus' },
  { value: 'south-campus', label: 'South Campus' },
  { value: 'city-branch', label: 'City Branch' },
  { value: 'research-center', label: 'Research Center' },
  { value: 'online', label: 'Online / Distance Learning' }
];

const DEFAULT_DEPARTMENT_OPTIONS = [
  { value: 'cse', label: 'Computer Science' },
  { value: 'ece', label: 'Electronics & Communication' },
  { value: 'mech', label: 'Mechanical Engineering' },
  { value: 'civil', label: 'Civil Engineering' },
  { value: 'hr', label: 'Human Resources' },
  { value: 'finance', label: 'Finance' },
  { value: 'frontend', label: 'Frontend' },
  { value: 'backend', label: 'Backend' }
];

let registerLookupState = {
  roles: [],
  departments: [],
  locations: []
};

let registerFlowState = {
  step: 1,
  accountCreated: false
};

function normalizeCode(value) {
  return String(value || '').replace(/\s+/g, '').toUpperCase();
}

function normalizeRoleKey(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return '';
  if (raw.includes('student')) return 'student';
  if (raw.includes('faculty') || raw.includes('teacher') || raw.includes('lecturer')) return 'faculty';
  if (raw.includes('admin')) return 'admin';
  if (raw.includes('employee') || raw.includes('staff') || raw.includes('assistant')) return 'employee';
  return raw;
}

function roleLabelFromKey(key) {
  const raw = normalizeRoleKey(key);
  if (raw === 'faculty') return 'Faculty';
  if (raw === 'employee') return 'Employee';
  if (raw === 'student') return 'Student';
  if (raw === 'admin') return 'Admin';
  return key || '';
}

function renderRoleOptions(options, placeholder) {
  const roleEl = document.getElementById('r-role');
  if (!roleEl) return;
  roleEl.innerHTML = `<option value="">${placeholder || 'Select...'}</option>` + options.map(opt => {
    const value = opt.role_id || opt.value || '';
    const label = roleLabelFromKey(opt.label || opt.name || value);
    const roleKey = normalizeRoleKey(opt.roleKey || opt.name || opt.label || value);
    return `<option value="${value}" data-role-key="${roleKey}">${label}</option>`;
  }).join('');
}

function renderBranchOptions() {
  const branchEl = document.getElementById('r-branch');
  if (!branchEl) return;
  branchEl.innerHTML = '<option value="">Select campus or branch...</option>' + REGISTER_BRANCH_OPTIONS.map(opt =>
    `<option value="${opt.value}">${opt.label}</option>`
  ).join('');
}

function renderDepartmentOptions(options, placeholder) {
  const deptEl = document.getElementById('r-dept');
  if (!deptEl) return;
  deptEl.innerHTML = `<option value="">${placeholder || 'Select...'}</option>` + options.map(opt => {
    const value = opt.department_id || opt.value || '';
    const label = opt.name || opt.label || value;
    return `<option value="${value}">${label}</option>`;
  }).join('');
}

function roleNameOf(userOrRole) {
  const raw = String(
    userOrRole?.roleKey ||
    userOrRole?.roleName ||
    userOrRole?.roleId ||
    userOrRole?.name ||
    userOrRole ||
    ''
  ).trim().toLowerCase();
  return raw;
}

function isAdminRole(userOrRole) {
  return roleNameOf(userOrRole) === 'admin';
}

function isTeacherRole(userOrRole) {
  return roleNameOf(userOrRole) === 'teacher';
}

function getTenantRoleOptions() {
  const mapped = (registerLookupState.roles || []).map(role => {
    const value = role.role_id || role.value || '';
    const key = normalizeRoleKey(role.name || role.label || role.roleKey || value);
    return {
      value,
      label: roleLabelFromKey(role.name || role.label || value),
      roleKey: key
    };
  }).filter(role => role.value);

  const fallbackByKey = new Map(COLLEGE_ROLE_OPTIONS.map(role => [role.roleKey, role]));
  mapped.forEach(role => fallbackByKey.delete(role.roleKey));
  return mapped.concat(Array.from(fallbackByKey.values()));
}

function getRolePlaceholder() {
  return 'Select institution role...';
}

function getRegisterRoleKey() {
  const roleEl = document.getElementById('r-role');
  if (!roleEl || roleEl.selectedIndex < 0) return '';
  const selected = roleEl.options[roleEl.selectedIndex];
  return normalizeRoleKey(selected?.dataset?.roleKey || selected?.textContent || selected?.value || '');
}

function setFieldState(fieldId, message) {
  const field = document.getElementById(fieldId);
  if (!field) return;
  const error = field.querySelector('.field-error');
  if (message) {
    field.classList.add('has-error');
    if (error) error.textContent = message;
    const input = field.querySelector('input, select, textarea');
    if (input) input.setAttribute('aria-invalid', 'true');
  } else {
    field.classList.remove('has-error');
    if (error) error.textContent = '';
    const input = field.querySelector('input, select, textarea');
    if (input) input.removeAttribute('aria-invalid');
  }
}

function clearRegisterErrors() {
  [
    'field-r-name','field-r-dob','field-r-email','field-r-mobile','field-r-emp-id',
    'field-r-institute','field-r-org-type','field-r-role','field-r-dept',
    'field-r-study-level','field-r-designation','field-r-password','field-r-confirm-password'
  ].forEach(id => setFieldState(id, ''));
}

function setRegisterBiometricVisible(visible) {
  const panel = document.getElementById('bio-panel');
  const btn = document.getElementById('btn-bio-reg');
  if (panel) panel.classList.toggle('hidden', !visible);
  if (btn) btn.disabled = !visible;
}

function refreshRegisterConditionalFields() {
  const roleKey = getRegisterRoleKey();
  const studyField = document.getElementById('field-r-study-level');
  const studyInput = document.getElementById('r-study-level');
  const desigField = document.getElementById('field-r-designation');
  const desigInput = document.getElementById('r-designation');

  const showStudent = roleKey === 'student';
  const showFacultyOrStaff = roleKey === 'faculty' || roleKey === 'employee';

  if (studyField) studyField.classList.toggle('hidden', !showStudent);
  if (studyInput) {
    studyInput.required = showStudent;
    if (!showStudent) {
      studyInput.value = '';
      setFieldState('field-r-study-level', '');
    }
  }

  if (desigField) desigField.classList.toggle('hidden', !showFacultyOrStaff);
  if (desigInput) {
    desigInput.required = showFacultyOrStaff;
    if (!showFacultyOrStaff) {
      desigInput.value = '';
      setFieldState('field-r-designation', '');
    }
  }
}

function applyTenantToRegistration() {
  const tenant = (window.TENANT || tenantState.institution) ? {
    institution: window.TENANT?.institution || tenantState.institution || {},
    orgType: window.TENANT?.orgType || tenantState.orgType || ''
  } : null;
  if (!tenant) return false;

  const orgNameInput = document.getElementById('r-institute');
  const orgTypeSelect = document.getElementById('r-org-type');
  const orgLabel = document.getElementById('r-org-name-label');
  const orgTypeLabel = document.querySelector('label[for="r-org-type"]');
  const orgName = String(tenant.institution?.name || '').trim();
  const orgType = String(tenant.orgType || tenant.institution?.orgType || 'college').toLowerCase();
  const orgTypeLabelText = orgType ? orgType.charAt(0).toUpperCase() + orgType.slice(1) : 'College';
  const city = String(tenant.institution?.city || '').trim();

  if (orgNameInput) {
    orgNameInput.value = orgName;
    orgNameInput.readOnly = true;
    orgNameInput.setAttribute('readonly', 'readonly');
    orgNameInput.setAttribute('aria-readonly', 'true');
  }

  if (orgTypeSelect) {
    orgTypeSelect.value = orgType || 'college';
    orgTypeSelect.disabled = true;
    orgTypeSelect.setAttribute('disabled', 'disabled');
  }

  if (orgLabel) {
    orgLabel.innerHTML = `Organization: ${orgName || 'Tenant'}${city ? ` (${city})` : ''}`;
  }

  if (orgTypeLabel) {
    orgTypeLabel.textContent = `Type: ${orgTypeLabelText}`;
  }

  const step2Title = document.getElementById('register-step-2-title');
  if (step2Title) {
    step2Title.textContent = `${orgName || 'Tenant'} · ${orgTypeLabelText}`;
  }

  return true;
}

function updateRegisterFormByRole() {
  const roleEl = document.getElementById('r-role');
  const orgTypeEl = document.getElementById('r-org-type');
  const orgInput = document.getElementById('r-institute');
  const orgLabel = document.getElementById('r-org-name-label');
  if (!roleEl || !orgTypeEl || !orgInput || !orgLabel) return;

  const selectedRole = String(roleEl.value || '');
  const roleOptions = getTenantRoleOptions();
  const finalRoleOptions = roleOptions.length ? roleOptions : COLLEGE_ROLE_OPTIONS;

  renderRoleOptions(finalRoleOptions, getRolePlaceholder());
  if (finalRoleOptions.some(opt => String(opt.value) === selectedRole)) roleEl.value = selectedRole;

  applyTenantToRegistration();
  refreshRegisterConditionalFields();
}

function renderStudentAttendancePanel(user, locations) {
  const panel = document.getElementById('student-att-panel');
  if (!panel) return;
  const deptCode = normalizeCode(user.deptId);
  const opts = locations.map(loc => `<option value="${loc.attendance_location_id}">${loc.name}</option>`).join('');
  panel.innerHTML = `
    <div class="section-label">Attendance Workspace</div>
    <div class="info-box green">
      <span>Signed in as ${user.name} Â· Department: ${user.deptId || 'N/A'}</span>
    </div>
    <div class="field">
      <label>Classroom / Lab <span class="req">*</span></label>
      <select id="student-location-select">
        <option value="">Select your ${deptCode || 'department'} location...</option>
        ${opts}
      </select>
    </div>
    <div class="hint">Only locations matching your department are shown, for example ${deptCode || 'CSE'}001.</div>
    <button class="btn btn-primary" id="btn-student-attendance" onclick="submitStudentAttendance()">
      Mark Attendance
    </button>
    <button class="btn btn-secondary" onclick="resetBiometricWorkspace()">Back</button>
  `;
  panel.style.display = 'block';
}

function applyRoleBasedUi(user, locations) {
  if (isAdminRole(user)) {
    renderAdminPanel(user);
    return;
  }
  renderStudentAttendancePanel(user, locations || []);
}

function renderAdminPanel(user) {
  const panel = document.getElementById('admin-panel');
  if (!panel) return;
  panel.innerHTML = `
    <div class="section-label">Admin Workspace</div>
    <div class="info-box green">
      <span>Signed in as ${user.name}. Add departments and department locations like labs or classrooms.</span>
    </div>
    <div class="section-label">Add Department</div>
    <div class="row2">
      <div class="field">
        <label>Department Code <span class="req">*</span></label>
        <input type="text" id="admin-dept-id" placeholder="e.g. CSE"/>
      </div>
      <div class="field">
        <label>Department Name <span class="req">*</span></label>
        <input type="text" id="admin-dept-name" placeholder="e.g. Computer Science"/>
      </div>
    </div>
    <button class="btn btn-primary" onclick="addDepartmentFromWorkspace()">Add Department</button>
    <div id="dept-list" class="att-list" style="margin-top:12px"></div>
    <div class="section-label" style="margin-top:14px">Add Classroom / Lab</div>
    <div class="row2">
      <div class="field">
        <label>Department <span class="req">*</span></label>
        <select id="admin-location-dept"></select>
      </div>
      <div class="field">
        <label>Classroom / Lab Name <span class="req">*</span></label>
        <input type="text" id="admin-location-name" placeholder="e.g. CSE001"/>
      </div>
    </div>
    <div class="row2">
      <div class="field">
        <label>Latitude <span class="req">*</span></label>
        <input type="number" id="admin-location-lat" step="0.00001" placeholder="13.32603"/>
      </div>
      <div class="field">
        <label>Longitude <span class="req">*</span></label>
        <input type="number" id="admin-location-lng" step="0.00001" placeholder="77.12621"/>
      </div>
    </div>
    <button class="btn btn-primary" onclick="addLocationFromWorkspace()">Add Location</button>
    <div id="loc-list" class="att-list" style="margin-top:12px"></div>
    <button class="btn btn-secondary" onclick="resetBiometricWorkspace()">Logout</button>
  `;
  panel.style.display = 'block';
  loadAdminWorkspaceData();
}

async function loadAdminWorkspaceData() {
  await Promise.all([loadDepts(), loadLocs()]);
  populateAdminDepartmentSelect();
}

function populateAdminDepartmentSelect() {
  const select = document.getElementById('admin-location-dept');
  if (!select) return;
  const departments = registerLookupState.departments || [];
  select.innerHTML = '<option value="">Select department...</option>' + departments.map(d =>
    `<option value="${d.department_id}">${d.name || d.department_id}</option>`
  ).join('');
}

function resetBiometricWorkspace() {
  const studentPanel = document.getElementById('student-att-panel');
  const adminPanel = document.getElementById('admin-panel');
  if (studentPanel) { studentPanel.style.display = 'none'; studentPanel.innerHTML = ''; }
  if (adminPanel) { adminPanel.style.display = 'none'; adminPanel.innerHTML = ''; }
  signedInUser = null;
  restoreSignInForm();
}

function resetRegisterFlow() {
  registerFlowState.step = 1;
  registerFlowState.accountCreated = false;
  clearRegisterErrors();
  setRegisterStep(1, { silent: true });
}

function setRegisterStep(step, options = {}) {
  const next = Math.min(Math.max(Number(step) || 1, 1), 3);
  registerFlowState.step = next;

  for (let i = 1; i <= 3; i++) {
    const panel = document.getElementById(`register-step-${i}`);
    const tab = document.getElementById(`register-step-tab-${i}`);
    const active = i === next;
    if (panel) panel.classList.toggle('active', active);
    if (tab) {
      tab.classList.toggle('active', active);
      tab.classList.toggle('completed', i < next);
      tab.setAttribute('aria-current', active ? 'step' : 'false');
    }
  }

  const bar = document.getElementById('register-progress-bar');
  if (bar) bar.style.width = next === 1 ? '33.33%' : next === 2 ? '66.66%' : '100%';

  if (!options.silent && !options.skipFocus) {
    const panel = document.getElementById(`register-step-${next}`);
    const focusTarget = panel?.querySelector('input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), button:not([disabled])');
    if (focusTarget) focusTarget.focus();
  }
}

function goRegisterStep(step) {
  const target = Math.min(Math.max(Number(step) || 1, 1), 3);
  if (target > registerFlowState.step) {
    for (let i = registerFlowState.step; i < target; i++) {
      if (!validateRegisterStep(i)) {
        setRegisterStep(i);
        return false;
      }
    }
  }
  setRegisterStep(target);
  return true;
}

function getRegisterValue(id) {
  const el = document.getElementById(id);
  return String(el?.value || '').trim();
}

function validateRegisterStep(step) {
  const current = Number(step) || 1;
  let valid = true;

  if (current === 1) {
    const name = getRegisterValue('r-name');
    const dob = getRegisterValue('r-dob');
    const email = getRegisterValue('r-email');
    const mobile = getRegisterValue('r-mobile');
    const memberId = getRegisterValue('r-employee-id');

    if (!name) { setFieldState('field-r-name', 'Name is required.'); valid = false; } else if (!isValidRegistrationName(name)) { setFieldState('field-r-name', 'Enter a valid full name.'); valid = false; } else setFieldState('field-r-name', '');
    if (!dob) { setFieldState('field-r-dob', 'Date of birth is required.'); valid = false; } else if (!isValidDob(dob)) { setFieldState('field-r-dob', 'Enter a valid date of birth.'); valid = false; } else setFieldState('field-r-dob', '');
    if (!email) { setFieldState('field-r-email', 'Email is required.'); valid = false; } else if (!isValidRegistrationEmail(email)) { setFieldState('field-r-email', 'Enter a valid email address.'); valid = false; } else setFieldState('field-r-email', '');
    if (!mobile) { setFieldState('field-r-mobile', 'Mobile number is required.'); valid = false; } else if (!isValidRegistrationMobile(mobile)) { setFieldState('field-r-mobile', 'Enter a valid mobile number.'); valid = false; } else setFieldState('field-r-mobile', '');
    if (!memberId) { setFieldState('field-r-emp-id', 'Student / Employee ID is required.'); valid = false; } else if (!/^[A-Za-z0-9][A-Za-z0-9_.-]{2,39}$/.test(memberId)) { setFieldState('field-r-emp-id', 'Use 3-40 letters, numbers, dots, dashes, or underscores.'); valid = false; } else setFieldState('field-r-emp-id', '');
  }

  if (current === 2) {
    const role = getRegisterValue('r-role');
    const dept = getRegisterValue('r-dept');
    const roleKey = getRegisterRoleKey();
    const studyLevel = getRegisterValue('r-study-level');
    const designation = getRegisterValue('r-designation');
    const tenantOrgName = String(window.TENANT?.institution?.name || tenantState.institution?.name || '').trim();
    const tenantOrgType = String(window.TENANT?.orgType || tenantState.orgType || '').trim();

    if (!tenantOrgName) { setFieldState('field-r-institute', 'Tenant organization is not loaded.'); valid = false; } else setFieldState('field-r-institute', '');
    if (!tenantOrgType) { setFieldState('field-r-org-type', 'Tenant organization type is not loaded.'); valid = false; } else setFieldState('field-r-org-type', '');
    if (!role) { setFieldState('field-r-role', 'Select a role.'); valid = false; } else setFieldState('field-r-role', '');
    if (!dept) { setFieldState('field-r-dept', 'Department is required.'); valid = false; } else if (!isValidDepartmentValue(dept)) { setFieldState('field-r-dept', 'Enter a valid department.'); valid = false; } else setFieldState('field-r-dept', '');

    if (roleKey === 'student') {
      if (!studyLevel) { setFieldState('field-r-study-level', 'Select a semester or year of study.'); valid = false; } else setFieldState('field-r-study-level', '');
    }
    if (roleKey === 'faculty' || roleKey === 'employee') {
      if (!designation) { setFieldState('field-r-designation', 'Designation is required for faculty or employee.'); valid = false; } else setFieldState('field-r-designation', '');
    }
  }

  if (current === 3) {
    const pass = getRegisterValue('r-password');
    const confirmPass = getRegisterValue('r-confirm-password');

    if (!pass) { setFieldState('field-r-password', 'Password is required.'); valid = false; } else if (pass.length < 8) { setFieldState('field-r-password', 'Password must be at least 8 characters.'); valid = false; } else if (!/[A-Za-z]/.test(pass) || !/[0-9]/.test(pass)) { setFieldState('field-r-password', 'Password must include at least one letter and one number.'); valid = false; } else setFieldState('field-r-password', '');
    if (!confirmPass) { setFieldState('field-r-confirm-password', 'Confirm your password.'); valid = false; } else if (pass !== confirmPass) { setFieldState('field-r-confirm-password', 'Passwords do not match.'); valid = false; } else setFieldState('field-r-confirm-password', '');
  }

  return valid;
}

async function loadRegisterLookups() {
  syncOrganizationName();
  try {
    const [roleRes, deptRes, locRes] = await Promise.all([
      api({ action: 'getRoles' }),
      api({ action: 'getDepartments' }),
      api({ action: 'getAttendanceLocations' })
    ]);

    const roles = (roleRes.data || []).filter(r => r && (r.role_id || r.name));
    registerLookupState.roles = roles;
    tenantState.roles = roles;

    const departments = (deptRes.data || []).filter(d => d && d.department_id);
    registerLookupState.departments = departments;
    tenantState.departments = departments;

    const locations = (locRes.data || []).filter(l => l && l.attendance_location_id);
    registerLookupState.locations = locations;
    tenantState.attendanceLocations = locations;
    updateRegisterFormByRole();
    populateAdminDepartmentSelect();
  } catch (e) {
    if (!registerLookupState.departments || !registerLookupState.departments.length) {
      registerLookupState.departments = DEFAULT_DEPARTMENT_OPTIONS;
    }
    renderRoleOptions(COLLEGE_ROLE_OPTIONS, getRolePlaceholder());
    refreshRegisterConditionalFields();
    toast('Could not load register form options', 'error');
  }
}

function isValidDob(value) {
  if (!value) return false;
  const dob = new Date(value + 'T00:00:00');
  if (Number.isNaN(dob.getTime())) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (dob >= today) return false;
  const minDob = new Date(today);
  minDob.setFullYear(minDob.getFullYear() - 100);
  return dob >= minDob;
}

async function handleBiometricSignIn() {
  if(!window.PublicKeyCredential){toast('WebAuthn not supported on this browser','error');return;}
  const email = document.getElementById('si-email').value.trim();
  if(!email){toast('Enter your email first','error');return;}

  if(navigator.permissions){
    try{
      const p=await navigator.permissions.query({name:'geolocation'});
      if(p.state==='denied'){showLocBar('fail','Location blocked â€” allow in Settings â†’ Site permissions');toast('ðŸ“ Allow location in browser settings and retry','error');return;}
    }catch(e){}
  }

  const btn=document.getElementById('btn-bio-signin');
  if(btn){btn._h=btn.innerHTML;btn.innerHTML='<span class="spin"></span> Verifyingâ€¦';btn.disabled=true;}
  try{
    const begin = await api({action:'beginWebAuthnLogin', email, guid: tenantState.guid});
    if(!begin.success||!begin.credentialId||!begin.challenge){toast(begin.message||'No biometric registered','error');return;}

    const cred = await navigator.credentials.get({publicKey:{
      challenge: base64UrlToUint8Array(begin.challenge),
      userVerification:'required', timeout:60000,
      allowCredentials:[{type:'public-key',id:base64UrlToUint8Array(begin.credentialId)}]
    }});

    const deviceId = await getDeviceId();
    const session = await api({
      action: 'finishWebAuthnLogin',
      email,
      challengeId: begin.challengeId,
      challenge: begin.challenge,
      credentialId: bufferToBase64Url(cred.rawId),
      clientDataJSON: cred.response?.clientDataJSON ? bufferToBase64Url(cred.response.clientDataJSON) : '',
      authenticatorData: cred.response?.authenticatorData ? bufferToBase64Url(cred.response.authenticatorData) : '',
      signature: cred.response?.signature ? bufferToBase64Url(cred.response.signature) : '',
      userHandle: cred.response?.userHandle ? bufferToBase64Url(cred.response.userHandle) : '',
      deviceId,
      guid: tenantState.guid
    });
    if(!session.success){toast(session.message||'Biometric login failed','error');return;}

    signedInUser = session;
    persistTeacherSession(session);
    const roleValue = normalizeRoleKey(session.roleKey || session.roleId || '');
    showLocBar('ok','Fingerprint / Face ID verified');

    if (isAdminRole(roleValue)) {
      toast('âœ“ Admin signed in','success');
      return;
    }

    await submitStudentAttendance('biometric');
  }catch(e){
    if(e.name==='NotAllowedError')toast('Biometric cancelled','warn');
    else toast('Error: '+e.message,'error');
  }finally{
    if(btn){btn.innerHTML=btn._h||btn.innerHTML;btn.disabled=false;}
  }
}

async function handlePasswordSignIn() {
  const email = document.getElementById('si-email').value.trim();
  const password = document.getElementById('si-password').value;
  if(!email || !password){toast('Enter email and password first','error');return;}

  const btn = document.getElementById('btn-pass-signin');
  if(btn){btn._h=btn.innerHTML;btn.innerHTML='<span class="spin"></span> Verifyingâ€¦';btn.disabled=true;}
  try{
    const deviceId = await getDeviceId();
    const info = await api({action:'loginUser', email, password, deviceId, guid: tenantState.guid});
    if(!info.success){toast(info.message||'Invalid credentials','error');return;}

    signedInUser = info;
    persistTeacherSession(info);
    const roleValue = normalizeRoleKey(info.roleKey || info.roleId || '');
    showLocBar('ok','Password verified');

    if (isAdminRole(roleValue)) {
      toast('âœ“ Admin signed in','success');
      return;
    }

    await submitStudentAttendance('password');
  }catch(e){
    toast('Error: '+e.message,'error');
  }finally{
    if(btn){btn.innerHTML=btn._h||btn.innerHTML;btn.disabled=false;}
  }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// MARK EXIT
// Updates Attendance row: typeâ†’exit, exit_time
// Also writes to LocationMonitor
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function handleExit() {
  if(!markedUserId){toast('Mark attendance first','error');return;}
  const btn=document.getElementById('btn-exit');
  if(btn){btn.disabled=true;btn.textContent='Getting exit locationâ€¦';}
  try{
    showLocBar('getting','Getting exit locationâ€¦');
    const loc = await getLocationWithAddress();
    if(loc.latitude)showLocBar('ok','Exit: '+(loc.address||`${loc.latitude}, ${loc.longitude}`),loc.accuracy);
    else showLocBar('fail','Exit location not captured');

    const res = await api({
      action:    'markExit',
      userId:    markedUserId,
      latitude:  loc.latitude,
      longitude: loc.longitude,
      address:   loc.address
    });

    if(res.success){
      toast('âœ“ '+res.message,'success');
      const card=document.querySelector('.att-card');
      if(card){
        card.querySelector('.att-name').textContent='âœ“ Entry & Exit Recorded';
        card.innerHTML+=`
          <div class="att-row"><span>exit_time</span><span class="att-val">${res.exitTime||''}</span></div>
          <div class="att-row"><span>type_attendance</span><span class="att-val">exit</span></div>
          <div class="att-row"><span>duration</span><span class="att-val">${res.duration||res.durationMins+' min'||''}</span></div>
          <div class="att-row"><span>address (exit)</span><span class="att-val">${res.location||'not captured'}</span></div>`;
      }
      stopTracking();
      if(btn){btn.disabled=true;btn.textContent='Exit recorded âœ“';btn.style.opacity='.5';}
      markedUserId=null;
      setTimeout(()=>restoreSignInForm(),4000);
    }else{
      toast(res.message,'error');
      if(btn){btn.disabled=false;btn.innerHTML='<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9"/></svg> Mark Exit (leaving classroom)';}
    }
  }catch(e){toast('Error: '+e.message,'error');if(btn){btn.disabled=false;btn.textContent='Mark Exit (leaving classroom)';}}
}

async function submitStudentAttendance(loginMethod = 'biometric') {
  if (!signedInUser || !signedInUser.userId) {
    toast('Sign in first','error');
    return;
  }
  const btn = document.getElementById('btn-student-attendance');
  if (btn) setLoading('btn-student-attendance', true);
  try {
    showLocBar('getting','Getting your locationâ€¦');
    const loc = await getLocationWithAddress();
    if(loc.denied){ if (btn) setLoading('btn-student-attendance', false); return; }
    if(!loc.latitude || !loc.longitude){
      showLocBar('fail','Location not captured â€” attendance blocked');
      toast('GPS location is required to mark attendance','error');
      if (btn) setLoading('btn-student-attendance', false);
      return;
    }
    showLocBar('ok',loc.address||`${loc.latitude}, ${loc.longitude}`,loc.accuracy);

    const att = await api({
      action: 'markEntry',
      userId: signedInUser.userId,
      loginMethod,
      latitude: loc.latitude,
      longitude: loc.longitude,
      address: loc.address
    });

    if(att.success){
      toast('âœ“ '+att.message,'success');
      showAttendanceCard({...att, method: loginMethod}, signedInUser.userId);
    } else if(att.code==='TOO_FAR'){
      showLocBar('fail',`${att.distance}m from location â€” must be within ${att.allowed}m`);
      toast(`ðŸ“ Too far (${att.distance}m). Move closer and try again`,'error');
    } else {
      toast(att.message,'error');
    }
  } catch(e) {
    toast('Error: ' + e.message,'error');
  } finally {
    if (btn) setLoading('btn-student-attendance', false);
  }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// MY ATTENDANCE
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function toggleMyAtt() {
  const list=document.getElementById('my-att-list');
  if(list.style.display==='block'){list.style.display='none';return;}
  list.style.display='block';
  const uid=sessionStorage.getItem('ba_uid')||markedUserId;
  if(!uid){list.innerHTML='<div style="color:var(--muted);font-size:11.5px;text-align:center;padding:8px">Sign in first to view history</div>';return;}
  list.innerHTML='<div style="text-align:center;color:var(--muted);font-size:11.5px;padding:8px">Loadingâ€¦</div>';
  try{
    const d=await api({action:'getMyAttendance',userId:uid});
    if(!d.records||!d.records.length){list.innerHTML='<div style="text-align:center;color:var(--muted);font-size:11.5px;padding:8px">No records yet</div>';return;}
    list.innerHTML=d.records.map(r=>`
      <div class="my-item">
        <div class="my-date">${r.date}</div>
        <div class="my-row"><span>entry_time</span><span>${r.entryTime||'â€”'}</span></div>
        <div class="my-row"><span>exit_time</span><span>${r.exitTime||'â€”'}</span></div>
        <div class="my-row"><span>duration</span><span>${r.duration||'â€”'}</span></div>
        <div class="my-row"><span>login_method</span><span>${r.loginMethod||'â€”'}</span></div>
        <div class="my-row"><span>address</span><span>${r.address||'not captured'}</span></div>
        <div class="my-row"><span>distance_from_centre</span><span>${r.distanceFromCentre||'â€”'} m</span></div>
      </div>`).join('');
  }catch(e){list.innerHTML='<div style="color:var(--danger);font-size:11.5px;text-align:center;padding:8px">Error: '+e.message+'</div>';}
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// REGISTER â€” biometric first, then account creation
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function handleRegister() {
  return handleRegisterV2();
}

async function registerBiometric(userId) {
  if (!window.PublicKeyCredential) {
    throw new Error('WebAuthn not supported');
  }
  const cred = await navigator.credentials.create({ publicKey: {
    challenge: crypto.getRandomValues(new Uint8Array(32)),
    rp: { name: `BioAttend ${tenantState.institution?.name || 'BioAttend'}`.trim() },
    user: {
      id: new TextEncoder().encode(String(userId)),
      name: String(userId),
      displayName: 'User'
    },
    pubKeyCredParams: [{ type: 'public-key', alg: -7 }],
    authenticatorSelection: {
      authenticatorAttachment: 'platform',
      userVerification: 'required'
    },
    attestation: 'none',
    timeout: 60000
  }});
  return {
    credentialId: bufferToBase64Url(cred.rawId)
  };
}

async function handleRegisterV2() {
  clearRegisterErrors();
  refreshRegisterConditionalFields();

  if (!validateRegisterStep(1)) {
    goRegisterStep(1);
    toast('Complete Step 1 before continuing', 'error');
    return;
  }
  if (!validateRegisterStep(2)) {
    goRegisterStep(2);
    toast('Complete Step 2 before continuing', 'error');
    return;
  }
  if (!validateRegisterStep(3)) {
    goRegisterStep(3);
    toast('Fix the security fields to continue', 'error');
    return;
  }

  const name   = getRegisterValue('r-name');
  const email  = getRegisterValue('r-email');
  const pass   = document.getElementById('r-password').value;
  const dob    = getRegisterValue('r-dob');
  const mobile = getRegisterValue('r-mobile');
  const dept   = getRegisterValue('r-dept');
  const role   = document.getElementById('r-role').value;
  const inst   = String(window.TENANT?.institution?.name || tenantState.institution?.name || '').trim();
  const orgType = String(window.TENANT?.orgType || tenantState.orgType || '').trim();
  const memberId = getRegisterValue('r-employee-id');
  const studyLevel = document.getElementById('r-study-level').value;
  const designation = getRegisterValue('r-designation');

  setLoading('btn-register',true);
  try{
    const biometric = await registerBiometric(memberId || email);
    const dId = await getDeviceId();
    const d   = await api({
      action:               'registerUser',
      name, email, password:pass, dob, mobile,
      departmentId:         dept,
      roleId:               role,
      instituteId:          inst,
      orgType:              orgType,
      studentEmployeeId:    memberId,
      studyLevel:           studyLevel,
      designation:          designation,
      biometricCode:        biometric.credentialId,
      deviceId:             dId
    });
    if(d.success){
      registeredUid=d.userId;
      registerFlowState.accountCreated = true;
      toast('âœ“ Account created with biometric access.','success');
    }else toast(d.message,'error');
  }catch(e){
    if(e.name==='NotAllowedError') toast('Biometric cancelled','warn');
    else toast('Error: '+e.message,'error');
  }
  setLoading('btn-register',false);
}

async function handleBiometricRegister() {
  try{
    toast('Biometric is now requested automatically before account creation.','info');
  }catch(e){toast('Biometric error: '+e.message,'error');}
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// TEACHER LOGIN & DASHBOARD
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function handleTeacherLogin() {
  const email=document.getElementById('t-email').value.trim();
  const pass=document.getElementById('t-password').value;
  if(!email||!pass){toast('Enter email and password','error');return;}
  setLoading('btn-t-login',true);
  try{
    const deviceId = await getDeviceId();
    const d=await api({action:'loginUser',email,password:pass,deviceId,guid: tenantState.guid});
    if(!d.success){toast(d.message,'error');setLoading('btn-t-login',false);return;}
    if(!isTeacherRole(d) && !isAdminRole(d)){toast('Not a teacher or admin account','error');setLoading('btn-t-login',false);return;}
    teacherData=d;
    persistTeacherSession(d);
    document.getElementById('t-login-section').style.display='none';
    document.getElementById('t-dashboard').style.display='block';
    document.getElementById('t-welcome').textContent='Hello, '+d.name;
    switchSub('session'); checkActiveSess();
  }catch(e){toast('Error: '+e.message,'error');}
  setLoading('btn-t-login',false);
}

async function teacherLogout(silent=false){
  teacherData=null;clearInterval(sessionTimer);liveSessionId=null;
  stopLivePolling();
  liveLastSyncTime='';
  liveData=null;
  liveMapMarkers = {};
  if (liveMap) {
    try { liveMap.remove(); } catch (e) {}
    liveMap = null;
  }
  if (analyticsCharts.daily) { try { analyticsCharts.daily.destroy(); } catch (e) {} analyticsCharts.daily = null; }
  if (analyticsCharts.weekly) { try { analyticsCharts.weekly.destroy(); } catch (e) {} analyticsCharts.weekly = null; }
  if(!silent){
    try{await api({action:'logout', guid: tenantState.guid});}catch(e){}
  }
  clearTeacherSession();
  const loginSection = document.getElementById('t-login-section');
  const dashboard = document.getElementById('t-dashboard');
  const email = document.getElementById('t-email');
  const password = document.getElementById('t-password');
  if(loginSection) loginSection.style.display='block';
  if(dashboard) dashboard.style.display='none';
  if(email) email.value='';
  if(password) password.value='';
}

async function checkActiveSess(){
  try{
    const d=await api({action:'getActiveSession'});
    const el=document.getElementById('active-sess-display');
    const form=document.getElementById('open-sess-form');
    if(d.active){
      let secs=d.secondsLeft; liveSessionId=d.session.session_id;
      el.innerHTML=`<div class="sess-card">
        <div class="sess-subj">ðŸŸ¢ ${d.session.subject} â€” LIVE</div>
        <div class="sess-meta">Closes in <span id="t-timer" style="font-weight:700;color:var(--success)">${fmtTime(secs)}</span></div>
        <button onclick="closeSess('${d.session.session_id}')" class="btn btn-danger" style="margin-top:9px;padding:8px">Stop Session</button>
      </div>`;
      form.style.display='none';
      clearInterval(sessionTimer);
      sessionTimer=setInterval(()=>{secs--;const t=document.getElementById('t-timer');if(t)t.textContent=fmtTime(secs);if(secs<=0){clearInterval(sessionTimer);checkActiveSess();}},1000);
    }else{liveSessionId=null;el.innerHTML='';form.style.display='block';}
  }catch(e){}
}

async function openSession(){
  const subj=document.getElementById('t-subject').value.trim();
  const win=parseInt(document.getElementById('t-window').value);
  if(!subj){toast('Enter subject name','error');return;}
  setLoading('btn-open-sess',true);
  try{
    const d=await api({action:'createSession',userId:teacherData.userId,teacherName:teacherData.name,
                       roleId:teacherData.roleId,subject:subj,windowMinutes:win});
    if(d.success){toast('âœ“ Session opened','success');document.getElementById('t-subject').value='';liveLastSyncTime='';checkActiveSess();if(document.getElementById('sp-live')?.classList.contains('active')) refreshLive(true);}
    else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
  setLoading('btn-open-sess',false);
}

async function closeSess(sid){
  try{await api({action:'closeSession',sessionId:sid});toast('Session closed','success');clearInterval(sessionTimer);liveSessionId=null;liveLastSyncTime='';checkActiveSess();stopLivePolling();}
  catch(e){toast('Error','error');}
}

// â”€â”€ Live dashboard â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function normalizeLivePayload(payload) {
  const attendance = payload?.liveAttendance || payload || {};
  const activeUsers = attendance.activeUsers || attendance.present || [];
  const offlineUsers = attendance.offlineUsers || attendance.absent || [];
  const recentlyExited = attendance.recentlyExited || [];
  return {
    sessionId: attendance.sessionId || payload?.sessionId || liveSessionId || '',
    session: attendance.session || payload?.session || null,
    totalIn: attendance.totalIn ?? activeUsers.length,
    totalOut: attendance.totalOut ?? offlineUsers.length,
    activeUsers,
    offlineUsers,
    recentlyExited,
    locations: payload?.locations?.locations || payload?.locations || [],
    updatedAt: attendance.updatedAt || payload?.syncedAt || payload?.updatedAt || new Date().toISOString()
  };
}

async function refreshLive(force = false, internal = false){
  if (liveRefreshInFlight) {
    liveRefreshQueued = true;
    return liveData;
  }
  liveRefreshInFlight = true;
  markLiveActivity();
  const listEl=document.getElementById('live-list');
  const infoEl=document.getElementById('live-info');
  const statEl=document.getElementById('live-stats');
  const toolEl=document.getElementById('live-toolbar');
  const refEl=document.getElementById('live-refresh');
  try{const chk=await api({action:'getActiveSession'});if(chk.active)liveSessionId=chk.session.session_id;}catch(e){}
  if(!liveSessionId){
    infoEl.className='no-session';infoEl.style.display='block';infoEl.textContent='No active session â€” open one from Session tab';
    statEl.style.display='none';toolEl.style.display='none';refEl.style.display='none';listEl.innerHTML='';
    renderLiveMap([]);
    stopLivePolling();
    liveRefreshInFlight = false;
    return null;
  }
  try{
    const payload = !liveLastSyncTime || force
      ? { action:'getLiveAttendance', sessionId: liveSessionId }
      : { action:'getLiveUpdates', sessionId: liveSessionId, lastSyncTime: liveLastSyncTime, waitMs: internal ? 4000 : 0 };
    const d=await api(payload);
    if(!d.success){throw new Error(d.message || 'Live refresh failed');}
    liveData=normalizeLivePayload(d);
    liveLastSyncTime=liveData.updatedAt;
    liveRetryDelay = 3000;
    infoEl.style.display='none';statEl.style.display='grid';toolEl.style.display='flex';refEl.style.display='block';
    const activeCount = liveData.activeUsers.length;
    const offlineCount = liveData.offlineUsers.length;
    const total = activeCount + offlineCount;
    document.getElementById('sp').textContent=activeCount;
    document.getElementById('sa').textContent=offlineCount;
    document.getElementById('st').textContent=total;
    document.getElementById('spc').textContent=total?Math.round(activeCount/total*100)+'%':'0%';
    const updatedEl = document.getElementById('live-updated');
    if (updatedEl) updatedEl.textContent='Updated '+new Date(liveData.updatedAt).toLocaleTimeString();
    renderLiveList();
    renderLiveMap(liveData.locations || []);
    return liveData;
  }catch(e){
    liveRetryDelay = Math.min(liveRetryDelay * 2, 30000);
    if (internal && isTeacherDashboardVisible() && isLiveTabActive() && document.visibilityState === 'visible') {
      if (livePollTimer) clearTimeout(livePollTimer);
      livePollTimer = setTimeout(livePollTick, liveRetryDelay);
    }
    toast('Error: '+e.message,'error');
    return null;
  } finally {
    liveRefreshInFlight = false;
    if (liveRefreshQueued) {
      liveRefreshQueued = false;
      if (!livePollTimer && isTeacherDashboardVisible() && isLiveTabActive() && document.visibilityState === 'visible') {
        livePollTimer = setTimeout(livePollTick, getLivePollDelay() || liveRetryDelay);
      }
    }
  }
}

function showLive(tab){
  liveTab=tab;
  document.getElementById('seg-p').classList.toggle('active',tab==='present');
  document.getElementById('seg-a').classList.toggle('active',tab==='absent');
  const segR = document.getElementById('seg-r');
  if (segR) segR.classList.toggle('active',tab==='recent');
  renderLiveList();
  if (liveData) renderLiveMap(liveData.locations || []);
}

function renderLiveList(){
  if(!liveData)return;
  const el=document.getElementById('live-list');
  let list=[];
  if(liveTab==='present') list=liveData.activeUsers||[];
  else if(liveTab==='absent') list=liveData.offlineUsers||[];
  else list=liveData.recentlyExited||[];
  if(!list||!list.length){
    el.innerHTML=`<div style="text-align:center;color:var(--muted);padding:16px;font-size:12px">${liveTab==='present'?'No one is currently inside campus':(liveTab==='absent'?'Everyone is online or no active attendance yet':'No recent exits')}</div>`;
    return;
  }
  el.innerHTML=list.map((s,i)=>{
    const statusBadge = liveTab==='present'
      ? `<span class="badge" style="background:#dcfce7;color:#166534">🟢 Online</span>`
      : liveTab==='recent'
        ? `<span class="badge" style="background:#fef3c7;color:#92400e">🟡 Recently exited</span>`
        : `<span class="badge absent" style="background:#fee2e2;color:#991b1b">🔴 Offline</span>`;
    const loc = s.location ? `${s.location.latitude||''}, ${s.location.longitude||''}` : 'no live location';
    const forceBtn = liveTab==='present'
      ? `<button class="export-btn" style="margin-left:8px;padding:5px 8px" onclick="event.stopPropagation();forceExitLiveUser('${s.userId}')">Force exit</button>`
      : '';
    return `<div class="att-item">
      <div style="flex:1">
        <div class="iname">${i+1}. ${s.name}</div>
        <div class="imeta">${s.email} · ${s.department||'—'} · ${s.entryTime||''}${s.exitTime?' → exit '+s.exitTime:''} · ${loc}</div>
        <div class="imeta" style="margin-top:3px">${s.lastSeenAt ? 'last seen ' + new Date(s.lastSeenAt).toLocaleTimeString() : ''}</div>
      </div>
      <div style="display:flex;flex-direction:column;align-items:flex-end;gap:6px">
        ${statusBadge}
        ${forceBtn}
      </div>
    </div>`;
  }).join('');
}

function ensureLiveMap() {
  const el = document.getElementById('live-map');
  if (!el || typeof L === 'undefined') return null;
  if (!liveMap) {
    liveMap = L.map('live-map', { zoomControl: true, attributionControl: false }).setView([13.32609, 77.12623], 16);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    }).addTo(liveMap);
  }
  setTimeout(() => {
    try { liveMap.invalidateSize(); } catch (e) {}
  }, 60);
  return liveMap;
}

function renderLiveMap(locations) {
  const map = ensureLiveMap();
  if (!map) return;
  const rows = Array.isArray(locations) ? locations : [];
  Object.keys(liveMapMarkers || {}).forEach(key => {
    try { map.removeLayer(liveMapMarkers[key]); } catch (e) {}
  });
  liveMapMarkers = {};

  const activeSet = new Set((liveData?.activeUsers || []).map(u => String(u.userId || '')));
  const recentSet = new Set((liveData?.recentlyExited || []).map(u => String(u.userId || '')));
  const points = [];

  rows.forEach(row => {
    const lat = parseFloat(row.latitude);
    const lng = parseFloat(row.longitude);
    if (Number.isNaN(lat) || Number.isNaN(lng)) return;
    const uid = String(row.userId || row.user_id || '');
    let color = '#dc2626';
    let label = 'Offline';
    if (activeSet.has(uid)) {
      color = '#16a34a';
      label = 'Inside campus';
    } else if (recentSet.has(uid)) {
      color = '#ca8a04';
      label = 'Recently exited';
    } else if (row.online) {
      color = '#0ea5e9';
      label = 'Online';
    }
    const marker = L.circleMarker([lat, lng], {
      radius: 9,
      color,
      fillColor: color,
      fillOpacity: 0.8,
      weight: 2
    }).addTo(map);
    marker.bindPopup(`<strong>${row.name || uid}</strong><br>${label}<br>${lat.toFixed(5)}, ${lng.toFixed(5)}${row.timestamp ? '<br>' + row.timestamp : ''}`);
    liveMapMarkers[uid || `${lat}-${lng}`] = marker;
    points.push([lat, lng]);
  });

  if (points.length) {
    const bounds = L.latLngBounds(points);
    map.fitBounds(bounds.pad(0.2));
  } else {
    map.setView([13.32609, 77.12623], 16);
  }
}

function renderAnalyticsCharts(daily, weekly) {
  if (typeof Chart === 'undefined') return;
  const dailyCtx = document.getElementById('daily-chart');
  const weeklyCtx = document.getElementById('weekly-chart');
  if (analyticsCharts.daily) { analyticsCharts.daily.destroy(); analyticsCharts.daily = null; }
  if (analyticsCharts.weekly) { analyticsCharts.weekly.destroy(); analyticsCharts.weekly = null; }

  if (dailyCtx && daily) {
    const labels = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0') + ':00');
    analyticsCharts.daily = new Chart(dailyCtx, {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'Entries', data: daily.entryTrend || [], borderColor: '#16a34a', backgroundColor: 'rgba(22,163,74,0.15)', tension: 0.35, fill: true },
          { label: 'Exits', data: daily.exitTrend || [], borderColor: '#dc2626', backgroundColor: 'rgba(220,38,38,0.12)', tension: 0.35, fill: true }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
  }

  if (weeklyCtx && weekly?.series) {
    analyticsCharts.weekly = new Chart(weeklyCtx, {
      type: 'bar',
      data: {
        labels: weekly.series.map(r => r.date),
        datasets: [
          { label: 'Present', data: weekly.series.map(r => r.present), backgroundColor: '#2563eb' },
          { label: 'Late', data: weekly.series.map(r => r.late), backgroundColor: '#f59e0b' },
          { label: 'Early Exit', data: weekly.series.map(r => r.early), backgroundColor: '#ef4444' }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
  }
}

function renderAnalyticsSummary(stats, weekly, insights, summary) {
  const daily = stats || {};
  const presentEl = document.getElementById('an-present');
  const absentEl = document.getElementById('an-absent');
  const lateEl = document.getElementById('an-late');
  const rateEl = document.getElementById('an-rate');
  if (presentEl) presentEl.textContent = daily.totalPresent ?? 0;
  if (absentEl) absentEl.textContent = daily.totalAbsent ?? 0;
  if (lateEl) lateEl.textContent = daily.lateEntries ?? 0;
  if (rateEl) rateEl.textContent = (daily.avgAttendancePercent ?? 0) + '%';

  renderAnalyticsCharts(daily, weekly);

  const insightsBox = document.getElementById('insights-box');
  if (insightsBox && insights) {
    const late = (insights.lateEntries || []).slice(0, 5);
    const early = (insights.earlyExits || []).slice(0, 5);
    const abs = (insights.frequentAbsentees || []).slice(0, 5);
    insightsBox.innerHTML = [
      ...late.map(r => `<div class="att-item"><div><div class="iname">Late: ${r.name || r.userId}</div><div class="imeta">${r.date} · ${r.entryTime || ''}</div></div><span class="badge" style="background:#fef3c7;color:#92400e">Late</span></div>`),
      ...early.map(r => `<div class="att-item"><div><div class="iname">Early exit: ${r.name || r.userId}</div><div class="imeta">${r.date} · ${r.exitTime || ''}</div></div><span class="badge absent" style="background:#fee2e2;color:#991b1b">Early</span></div>`),
      ...abs.map(r => `<div class="att-item"><div><div class="iname">Frequent absentee: ${r.name || r.userId}</div><div class="imeta">${r.absentDays} absent days</div></div><span class="badge" style="background:#e0f2fe;color:#075985">Risk</span></div>`)
    ].join('') || '<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">No insights yet</div>';
  }

  const summaryBox = document.getElementById('summary-box');
  if (summaryBox) {
    if (!summary) {
      summaryBox.innerHTML = '<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">Enter a user ID to load a student summary</div>';
    } else {
      summaryBox.innerHTML = `
        <div class="att-item" style="flex-direction:column;align-items:stretch">
          <div style="display:flex;justify-content:space-between;gap:10px">
            <div>
              <div class="iname">${summary.name || summary.userId}</div>
              <div class="imeta">${summary.email || ''}</div>
            </div>
            <span class="badge" style="background:#dcfce7;color:#166534">${summary.attendancePercent || 0}%</span>
          </div>
          <div class="imeta" style="margin-top:8px">Present ${summary.presentCount || 0} · Late ${summary.lateEntries || 0} · Early exits ${summary.earlyExits || 0}</div>
          <div class="imeta" style="margin-top:4px">Recent records: ${summary.totalRecords || 0}</div>
        </div>`;
    }
  }
}

async function loadAnalytics(force = false){
  const dateEl = document.getElementById('analytics-date');
  if (dateEl && !dateEl.value) dateEl.value = new Date().toISOString().slice(0,10);
  try{
    const [daily, weekly, insights] = await Promise.all([
      api({action:'getDailyStats', date: dateEl ? dateEl.value : '', guid: tenantState.guid}),
      api({action:'getWeeklyStats', guid: tenantState.guid}),
      api({action:'getAttendanceInsights', days: 14, guid: tenantState.guid})
    ]);
    if (!daily.success) throw new Error(daily.message || 'Daily stats failed');
    if (!weekly.success) throw new Error(weekly.message || 'Weekly stats failed');
    if (!insights.success) throw new Error(insights.message || 'Insights failed');
    renderAnalyticsSummary(daily, weekly, insights, null);
  }catch(e){
    toast('Error: '+e.message,'error');
  }
}

async function loadUserAttendanceSummary(){
  const userId = document.getElementById('analytics-user-id')?.value.trim();
  if (!userId) {
    toast('Enter a user ID first','error');
    return;
  }
  try{
    const d = await api({action:'getUserAttendanceSummary', userId, guid: tenantState.guid});
    if(!d.success) throw new Error(d.message || 'Summary failed');
    const dateValue = document.getElementById('analytics-date')?.value || new Date().toISOString().slice(0,10);
    const daily = await api({action:'getDailyStats', date: dateValue, guid: tenantState.guid});
    const weekly = await api({action:'getWeeklyStats', guid: tenantState.guid});
    const insights = await api({action:'getAttendanceInsights', days: 14, guid: tenantState.guid});
    renderAnalyticsSummary(daily.success ? daily : null, weekly.success ? weekly : null, insights.success ? insights : null, d);
  }catch(e){
    toast('Error: '+e.message,'error');
  }
}

function exportLive(){
  if(!liveData)return;
  const rows=[['full_name','email','department_id','entry_time','exit_time','login_method','type_attendance']];
  (liveData.activeUsers||[]).forEach(s=>rows.push([s.name,s.email,s.department||'',s.entryTime,s.exitTime||'','','present']));
  (liveData.offlineUsers||[]).forEach(s=>rows.push([s.name,s.email,s.department||'','','','','absent']));
  dlCSV(rows,'attendance_live.csv');
}

async function forceExitLiveUser(userId){
  if(!userId)return;
  if(!confirm('Force this user to exit?')) return;
  try{
    const d=await api({action:'forceExitUser',userId,sessionId:liveSessionId,guid:tenantState.guid});
    if(d.success){toast('✓ User forced out','success');liveLastSyncTime='';refreshLive(true);}
    else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
}

// â”€â”€ History â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async function loadHistory(){
  const date=document.getElementById('hist-date').value;
  const el=document.getElementById('hist-list');
  el.innerHTML='<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">Loadingâ€¦</div>';
  try{
    const d=await api({action:'getSessions',userId:teacherData.userId});
    let sessions=d.sessions||[];
    if(date)sessions=sessions.filter(s=>s.date===date);
    historyData=sessions;
    if(!sessions.length){el.innerHTML='<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">No sessions found</div>';document.getElementById('hist-export').style.display='none';return;}
    document.getElementById('hist-export').style.display='block';
    el.innerHTML=sessions.map(s=>`
      <div class="att-item clickable" style="flex-direction:column;align-items:stretch" onclick="toggleSessDet('${s.sessionId}','${(s.subject||'').replace(/'/g,"\\'")}')">
        <div style="display:flex;justify-content:space-between;align-items:center">
          <div><div class="iname">${s.subject}</div><div class="imeta">${s.date} Â· ${s.startTime}â€“${s.endTime}</div></div>
          <div style="text-align:right"><span class="badge ${s.status==='open'?'':'closed'}">${s.presentCount||0} present</span><div style="font-size:9.5px;margin-top:2px;color:var(--muted)">${s.status}</div></div>
        </div>
        <div id="sd-${s.sessionId}" style="display:none;margin-top:9px;border-top:1px solid var(--border);padding-top:9px"></div>
      </div>`).join('');
  }catch(e){toast('Error: '+e.message,'error');}
}

async function toggleSessDet(sid,subj){
  const el=document.getElementById('sd-'+sid);if(!el)return;
  if(el.style.display==='block'){el.style.display='none';return;}
  el.style.display='block';el.innerHTML='<div style="color:var(--muted);font-size:11px">Loadingâ€¦</div>';
  try{
    const d=await api({action:'getDashboard',sessionId:sid});
    if(!d.success){el.innerHTML='<div style="color:var(--danger);font-size:11px">'+d.message+'</div>';return;}
    el.innerHTML=`
      <div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px">
        <span style="color:var(--success)">âœ“ ${d.presentCount}</span>
        <span style="color:var(--danger)">âœ— ${d.absentCount}</span>
        <span style="color:var(--muted)">${d.total} total Â· ${d.total?Math.round(d.presentCount/d.total*100):0}%</span>
      </div>
      ${d.present.map(s=>`<div style="font-size:11px;padding:3px 0;border-bottom:1px solid var(--border);color:var(--text)">${s.name} <span style="color:var(--muted)">${s.entryTime}${s.exitTime?' â†’ '+s.exitTime:''} Â· ${s.method}</span></div>`).join('')}
      ${!d.present.length?'<div style="font-size:11px;color:var(--muted)">No students marked</div>':''}
      <button onclick="event.stopPropagation();exportSession('${sid}','${subj}')" class="export-btn" style="float:none;margin-top:9px">â†“ Export CSV</button>`;
  }catch(e){el.innerHTML='<div style="color:var(--danger);font-size:11px">Error</div>';}
}

async function exportSession(sid,subj){
  try{
    const d=await api({action:'exportAttendance',sessionId:sid});
    if(d.success)dlCSV(null,'att_'+(subj||'export').replace(/\s+/g,'_')+'.csv',d.csv);
    else toast(d.message,'error');
  }catch(e){toast('Export failed','error');}
}

async function exportHistory(){
  if(!historyData.length)return;
  try{
    const d=await api({action:'exportAttendance',date:historyData[0]?.date});
    if(d.success)dlCSV(null,'att_'+(historyData[0]?.date||'all')+'.csv',d.csv);
    else toast(d.message,'error');
  }catch(e){toast('Export failed','error');}
}

// â”€â”€ Students roster â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async function loadStudents(){
  const el=document.getElementById('stud-list');
  el.innerHTML='<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">Loadingâ€¦</div>';
  try{
    const d=await api({action:'getStudents'});
    allStudents=d.students||[];
    document.getElementById('stud-count').textContent=allStudents.length+' students in Users table';
    renderStudents(allStudents);
  }catch(e){toast('Error','error');}
}

function filterStudents(){
  const q=document.getElementById('student-search').value.toLowerCase();
  const f=allStudents.filter(s=>(s.name||'').toLowerCase().includes(q)||(s.email||'').toLowerCase().includes(q)||(s.department||'').toLowerCase().includes(q));
  document.getElementById('stud-count').textContent=f.length+' of '+allStudents.length+' students';
  renderStudents(f);
}

function renderStudents(list){
  const el=document.getElementById('stud-list');
  if(!list.length){el.innerHTML='<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">No students found</div>';return;}
  el.innerHTML=list.map((s,i)=>`
    <div class="att-item">
      <div><div class="iname">${i+1}. ${s.name}</div><div class="imeta">${s.email} Â· dept: ${s.department||'â€”'}</div></div>
      <div style="text-align:right;font-size:10px;line-height:1.8">
        <span style="color:${s.hasBio?'var(--success)':'var(--muted)'}">${s.hasBio?'ðŸ” bio':'ðŸ” none'}</span><br>
        <span style="color:${s.hasDevice?'var(--success)':'var(--muted)'}">${s.hasDevice?'ðŸ“± bound':'ðŸ“± none'}</span>
      </div>
    </div>`).join('');
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ADMIN â€” lookup table management
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// Departments
async function addDepartment(){
  const id=document.getElementById('ad-id').value.trim();
  const name=document.getElementById('ad-name').value.trim();
  const incharge=document.getElementById('ad-incharge').value.trim();
  const email=document.getElementById('ad-email').value.trim();
  if(!id||!name){toast('department_id and name are required','error');return;}
  try{
    // Write directly to Departments sheet via a dedicated action
    const d=await api({action:'addDepartment',departmentId:id,name,inCharge:incharge,email});
    if(d.success){toast('âœ“ Department added','success');document.getElementById('ad-id').value='';document.getElementById('ad-name').value='';loadDepts();}
    else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
}

async function addDepartmentFromWorkspace() {
  const deptId = document.getElementById('admin-dept-id')?.value.trim();
  const name = document.getElementById('admin-dept-name')?.value.trim();
  if(!deptId||!name){toast('Department code and name are required','error');return;}
  try{
    const d=await api({action:'addDepartment',departmentId:normalizeCode(deptId),name});
    if(d.success){
      toast('âœ“ Department added','success');
      document.getElementById('admin-dept-id').value='';
      document.getElementById('admin-dept-name').value='';
      await loadRegisterLookups();
      await loadDepts();
    } else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
}

async function loadDepts(){
  const el=document.getElementById('dept-list');
  try{
    const d=await api({action:'getDepartments'});
    const rows=d.data||[];
    el.innerHTML=rows.length?rows.map(r=>`<div class="att-item"><div><div class="iname">${r.department_id} â€” ${r.name}</div><div class="imeta">in_charge: ${r.in_charge||'â€”'} Â· ${r.email||'â€”'}</div></div></div>`).join(''):'<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">No departments yet</div>';
  }catch(e){}
}

// Attendance Locations
async function addLocation(){
  const id=document.getElementById('al-id').value.trim();
  const name=document.getElementById('al-name').value.trim();
  const lat=document.getElementById('al-lat').value.trim();
  const lng=document.getElementById('al-lng').value.trim();
  if(!id||!name||!lat||!lng){toast('All location fields are required','error');return;}
  try{
    const d=await api({action:'addAttendanceLocation',locationId:id,name,latitude:parseFloat(lat),longitude:parseFloat(lng)});
    if(d.success){toast('âœ“ Location added','success');document.getElementById('al-id').value='';document.getElementById('al-name').value='';loadLocs();}
    else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
}

async function addLocationFromWorkspace() {
  const deptId = document.getElementById('admin-location-dept')?.value.trim();
  const name = document.getElementById('admin-location-name')?.value.trim();
  const lat = document.getElementById('admin-location-lat')?.value.trim();
  const lng = document.getElementById('admin-location-lng')?.value.trim();
  if(!deptId||!name||!lat||!lng){toast('Department, classroom/lab name, latitude and longitude are required','error');return;}
  const normalizedName = normalizeCode(name);
  if (!normalizedName.startsWith(normalizeCode(deptId))) {
    toast('Location name should start with department code, for example CSE001','error');
    return;
  }
  const locationId = `${normalizeCode(deptId)}_${Date.now().toString().slice(-6)}`;
  try{
    const d=await api({
      action:'addAttendanceLocation',
      locationId,
      name: normalizedName,
      latitude: parseFloat(lat),
      longitude: parseFloat(lng)
    });
    if(d.success){
      toast('âœ“ Location added','success');
      document.getElementById('admin-location-name').value='';
      document.getElementById('admin-location-lat').value='';
      document.getElementById('admin-location-lng').value='';
      await loadRegisterLookups();
      await loadLocs();
    } else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
}

async function loadLocs(){
  const el=document.getElementById('loc-list');
  try{
    const d=await api({action:'getLocations'});
    const rows=d.data||[];
    el.innerHTML=rows.length?rows.map(r=>`<div class="att-item"><div><div class="iname">${r.attendance_location_id} â€” ${r.name}</div><div class="imeta">lat: ${r.latitude} Â· lng: ${r.longitude}</div></div></div>`).join(''):'<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">No locations yet</div>';
  }catch(e){}
}

// User â†’ Location map
async function addUserLocMap(){
  const uid=document.getElementById('ulm-uid').value.trim();
  const lid=document.getElementById('ulm-lid').value.trim();
  const dist=document.getElementById('ulm-dist').value.trim();
  if(!uid||!lid){toast('user_id and attendance_location_id are required','error');return;}
  try{
    const d=await api({action:'addUserLocMap',userId:uid,locationId:lid,allowedDistance:parseInt(dist)||200});
    if(d.success){toast('âœ“ Mapping added','success');document.getElementById('ulm-uid').value='';}
    else toast(d.message,'error');
  }catch(e){toast('Error: '+e.message,'error');}
}

// â”€â”€ CSV download â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function dlCSV(rows, filename, csvStr){
  const c=csvStr||rows.map(r=>r.map(x=>'"'+(String(x||'')).replace(/"/g,'""')+'"').join(',')).join('\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([c],{type:'text/csv'}));
  a.download=filename;a.click();
}

// â”€â”€ Init â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Restore session state (if student already marked today)
try{
  const uid=sessionStorage.getItem('ba_uid');
  if(uid){
    markedUserId=uid;
    showAttendanceCard({
      name:             sessionStorage.getItem('ba_name'),
      date:             sessionStorage.getItem('ba_date'),
      time:             sessionStorage.getItem('ba_time'),
      location:         sessionStorage.getItem('ba_loc'),
      method:           sessionStorage.getItem('ba_meth'),
      distanceFromCentre: sessionStorage.getItem('ba_dist')
    }, uid);
  }
}catch(e){}

// Online/offline detection
const offBar=document.getElementById('offline-bar');
function chkOnline(){offBar.classList.toggle('show',!navigator.onLine);}
window.addEventListener('online',chkOnline);window.addEventListener('offline',chkOnline);chkOnline();
const rDob=document.getElementById('r-dob');if(rDob)rDob.max=new Date().toISOString().slice(0,10);
syncOrganizationName();
try{
  const savedTeacher = localStorage.getItem('ba_teacher_session');
  if (savedTeacher) teacherData = JSON.parse(savedTeacher);
  if (teacherData?.authToken) scheduleAuthExpiry();
}catch(e){}
(async () => {
  const ok = await bootTenant();
  if (ok !== false) applyTenantToRegistration();
})();

