// ── app.js ────────────────────────────────────────────────────
// BioAttend frontend logic
// FIXES:
//  1. pubKeyCredParams now includes both ES256 (-7) and RS256 (-257)
//  2. loginUser action corrected to signIn (matches backend router)
//  3. credentialIdToUint8Array defined here (not just in tenant.js)
//  4. Registration flow: biometric registered DURING account creation
//  5. Teacher dashboard works without active session
//  6. Tenant loads even without ?q= param

// ── WebAuthn helpers ──────────────────────────────────────────
// FIX: Moved here so it's always available before tenant.js loads
function credentialIdToUint8Array(value) {
  const raw = String(value || '').trim();
  if (!raw) return new Uint8Array();
  
  const normalized = raw.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized + '='.repeat((4 - (normalized.length % 4 || 4)) % 4);
  try { return Uint8Array.from(atob(padded), c => c.charCodeAt(0)); }
  catch (e) { return new Uint8Array(); }
}

function bufferToBase64Url(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer || []);
  return btoa(String.fromCharCode(...bytes))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// ── Tab switching ─────────────────────────────────────────────
function switchMain(tab) {
  ['signin', 'register'].forEach(t => {
    const pane = document.getElementById('pane-' + t);
    const tabBtn = document.getElementById('mtab-' + t);
    if (pane) pane.classList.toggle('active', t === tab);
    if (tabBtn) tabBtn.classList.toggle('active', t === tab);
  });
  if (tab === 'register') {
    resetRegisterFlow();
    loadRegisterLookups();
    applyTenantToRegistration();
  }
}

// ── Live polling ──────────────────────────────────────────────
let livePollTimer = null;
let liveRefreshInFlight = false;
let liveRefreshQueued = false;
let liveRetryDelay = 3000;
let liveLastActivityAt = Date.now();
let liveLastRequestAt = 0;
let liveLastSyncTime = '';
let liveSessionId = null;
let liveData = null;
let liveTab = 'present';
let liveMap = null;
let liveMapMarkers = {};
let analyticsCharts = { daily: null, weekly: null };

function markLiveActivity() { liveLastActivityAt = Date.now(); }
['mousemove', 'keydown', 'touchstart', 'scroll', 'focus'].forEach(evt => {
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
  return Date.now() - liveLastActivityAt < 15000 ? 3000 : 10000;
}

async function livePollTick() {
  livePollTimer = null;
  if (!isTeacherDashboardVisible() || !isLiveTabActive() || document.visibilityState !== 'visible') return;
  await refreshLive(false, true);
  if (livePollTimer) return;
  const delay = getLivePollDelay();
  if (delay !== null) livePollTimer = setTimeout(livePollTick, delay);
}

function switchSub(tab) {
  ['session', 'live', 'history', 'students', 'analytics'].forEach(t => {
    const sp = document.getElementById('sp-' + t);
    const st = document.getElementById('stab-' + t);
    if (sp) sp.classList.toggle('active', t === tab);
    if (st) st.classList.toggle('active', t === tab);
  });
  if (tab === 'live') { refreshLive(true); startLivePolling(); }
  else stopLivePolling();
  if (tab === 'history') { const d = document.getElementById('hist-date'); if (d) d.value = new Date().toISOString().slice(0, 10); loadHistory(); }
  if (tab === 'students') loadStudents();
  if (tab === 'analytics') loadAnalytics(true);
}

function startLivePolling() {
  if (livePollTimer || !isTeacherDashboardVisible() || !isLiveTabActive()) return;
  const delay = getLivePollDelay();
  if (delay === null) return;
  livePollTimer = setTimeout(livePollTick, delay);
}

function stopLivePolling() {
  if (livePollTimer) { clearTimeout(livePollTimer); livePollTimer = null; }
  liveRefreshInFlight = false;
  liveRefreshQueued = false;
  liveRetryDelay = 3000;
}

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && isTeacherDashboardVisible() && isLiveTabActive()) {
    startLivePolling(); refreshLive();
  } else if (document.visibilityState !== 'visible') {
    stopLivePolling();
  }
});

function switchAdmin(tab) {
  ['dept', 'loc', 'ulmap', 'atttype'].forEach(t => {
    const ap = document.getElementById('ap-' + t);
    const at = document.getElementById('atab-' + t);
    if (ap) ap.classList.toggle('active', t === tab);
    if (at) at.classList.toggle('active', t === tab);
  });
  if (tab === 'dept') loadDepts();
  if (tab === 'loc') loadLocs();
}

// ── Device fingerprint ────────────────────────────────────────
let deviceId = null;
async function getDeviceId() {
  if (deviceId) return deviceId;
  try { const s = localStorage.getItem('ba_did'); if (s) { deviceId = s; return s; } } catch (e) {}
  const cv = document.createElement('canvas'), c = cv.getContext('2d');
  c.fillText('BioAttend', 2, 2);
  const raw = [navigator.userAgent, navigator.language, screen.width + 'x' + screen.height,
    new Date().getTimezoneOffset(), cv.toDataURL().slice(-40)].join('|');
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(raw));
  deviceId = Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('').slice(0, 16);
  try { localStorage.setItem('ba_did', deviceId); } catch (e) {}
  return deviceId;
}

// ── GPS ───────────────────────────────────────────────────────
function getLocation() {
  return new Promise(resolve => {
    if (!navigator.geolocation) return resolve({ latitude: '', longitude: '', accuracy: null, denied: false });
    let best = null, wid = null, done = false;
    const finish = r => {
      if (done) return; done = true;
      if (wid !== null) try { navigator.geolocation.clearWatch(wid); } catch (e) {}
      resolve(r || { latitude: '', longitude: '', accuracy: null, denied: false });
    };
    wid = navigator.geolocation.watchPosition(
      pos => {
        const r = { latitude: pos.coords.latitude, longitude: pos.coords.longitude, accuracy: pos.coords.accuracy, denied: false, address: '' };
        if (!best || r.accuracy < best.accuracy) best = r;
        if (r.accuracy <= 50) finish(best);
      },
      err => {
        if (err.code === 1) finish({ latitude: '', longitude: '', accuracy: null, denied: true });
        else finish(best || { latitude: '', longitude: '', accuracy: null, denied: false });
      },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
    setTimeout(() => finish(best), 10000);
  });
}

async function getLocationWithAddress() {
  const loc = await getLocation();
  if (loc.denied) { showLocBar('fail', 'Location blocked — allow in browser settings and retry'); return loc; }
  if (!loc.latitude) return loc;
  try {
    const r = await fetch(`https://nominatim.openstreetmap.org/reverse?lat=${loc.latitude}&lon=${loc.longitude}&format=json`, { headers: { 'Accept-Language': 'en' } });
    const d = await r.json();
    loc.address = d.display_name ? d.display_name.split(',').slice(0, 3).join(',').trim() : '';
  } catch (e) { loc.address = ''; }
  return loc;
}

function showLocBar(state, msg, accuracy) {
  const el = document.getElementById('loc-status-bar');
  if (!el) return;
  const icons = { getting: '📡', ok: '📍', fail: '⚠️' };
  const acc = accuracy ? ` <span style="opacity:.6;font-size:10px">±${Math.round(accuracy)}m</span>` : '';
  el.innerHTML = `<div class="loc-bar ${state}">${icons[state] || ''} ${msg}${acc}</div>`;
}

// ── Location tracking ─────────────────────────────────────────
function startTracking(userId) {
  if (!userId || !navigator.geolocation) return;
  stopTracking();
  window._trackingUserId = String(userId);
  window._trackingLastSentAt = 0;
  try { sessionStorage.setItem('ba_tracking_user', String(userId)); } catch (e) {}
  window._trackingWatchId = navigator.geolocation.watchPosition(
    pos => trackLocation(pos),
    err => console.warn('Tracking GPS error', err),
    { enableHighAccuracy: true, maximumAge: 0, timeout: 10000 }
  );
}

function stopTracking() {
  if (window._trackingWatchId != null) {
    try { navigator.geolocation.clearWatch(window._trackingWatchId); } catch (e) {}
    window._trackingWatchId = null;
  }
  window._trackingUserId = null;
  window._trackingLastSentAt = 0;
  try { sessionStorage.removeItem('ba_tracking_user'); } catch (e) {}
}

async function trackLocation(pos) {
  if (!window._trackingUserId || !pos?.coords) return;
  const now = Date.now();
  if (window._trackingLastSentAt && (now - window._trackingLastSentAt) < 30000) return;
  window._trackingLastSentAt = now;
  try {
    const res = await api({
      action: 'trackStudentLocation',
      userId: window._trackingUserId,
      latitude: pos.coords.latitude,
      longitude: pos.coords.longitude
    });
    if (res?.exitMarked) { stopTracking(); toast('Auto exit recorded', 'success'); restoreSignInForm(); }
  } catch (e) {}
}

// ── Attendance card ───────────────────────────────────────────
let markedUserId = null;
let signedInUser = null;

function showAttendanceCard(data, userId) {
  markedUserId = userId;
  const el = document.getElementById('att-success-card');
  el.style.display = 'block';
  el.innerHTML = `
    <div class="att-card">
      <div class="att-name">✓ Attendance Marked</div>
      <div class="att-row"><span>full_name</span><span class="att-val">${data.name || ''}</span></div>
      <div class="att-row"><span>attendance_date</span><span class="att-val">${data.date || ''}</span></div>
      <div class="att-row"><span>entry_time</span><span class="att-val">${data.time || ''}</span></div>
      <div class="att-row"><span>address</span><span class="att-val">${data.location || 'not captured'}</span></div>
      <div class="att-row"><span>distance_from_centre</span><span class="att-val">${data.distanceFromCentre || '—'} m</span></div>
      <div class="att-row"><span>login_method</span><span class="att-val">${data.method || 'biometric'}</span></div>
      <div class="att-row"><span>type_attendance</span><span class="att-val">entry</span></div>
    </div>
    <button class="btn-exit" id="btn-exit" onclick="handleExit()">
      <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9"/></svg>
      Mark Exit (leaving classroom)
    </button>`;
  const ef = document.getElementById('si-email');
  if (ef?.closest('.field')) ef.closest('.field').style.display = 'none';
  const ib = document.querySelector('#pane-signin .info-box');
  if (ib) ib.style.display = 'none';
  const bb = document.getElementById('btn-bio-signin');
  if (bb) bb.style.display = 'none';
  try {
    sessionStorage.setItem('ba_uid', userId);
    sessionStorage.setItem('ba_name', data.name || '');
    sessionStorage.setItem('ba_date', data.date || '');
    sessionStorage.setItem('ba_time', data.time || '');
    sessionStorage.setItem('ba_loc', data.location || '');
    sessionStorage.setItem('ba_meth', data.method || 'biometric');
    sessionStorage.setItem('ba_dist', data.distanceFromCentre || '');
  } catch (e) {}
  startTracking(userId);
}

function restoreSignInForm() {
  stopTracking();
  const card = document.getElementById('att-success-card');
  if (card) card.style.display = 'none';
  const lb = document.getElementById('loc-status-bar');
  if (lb) lb.innerHTML = '';
  const ef = document.getElementById('si-email');
  if (ef?.closest('.field')) ef.closest('.field').style.display = '';
  const ib = document.querySelector('#pane-signin .info-box');
  if (ib) ib.style.display = '';
  const bb = document.getElementById('btn-bio-signin');
  if (bb) { bb.style.display = ''; bb.disabled = false; }
  const pb = document.getElementById('btn-pass-signin');
  if (pb) { pb.style.display = ''; pb.disabled = false; }
  if (ef) ef.value = '';
  const pw = document.getElementById('si-password');
  if (pw) pw.value = '';
  markedUserId = null;
  try { ['ba_uid', 'ba_name', 'ba_date', 'ba_time', 'ba_loc', 'ba_meth', 'ba_dist'].forEach(k => sessionStorage.removeItem(k)); } catch (e) {}
}

// ═══════════════════════════════════════════════════════════════
// BIOMETRIC SIGN IN
// FIX: Uses correct credentialIdToUint8Array, proper error messages
// ═══════════════════════════════════════════════════════════════
async function handleBiometricSignIn() {
  if (!window.PublicKeyCredential) { toast('WebAuthn not supported on this browser', 'error'); return; }
  const email = document.getElementById('si-email').value.trim();
  if (!email) { toast('Enter your email first', 'error'); return; }

  const btn = document.getElementById('btn-bio-signin');
  if (btn) { btn._h = btn.innerHTML; btn.innerHTML = '<span class="spin"></span> Verifying…'; btn.disabled = true; }
  try {
    const info = await api({ action: 'getBiometric', email });
    if (!info.success || !info.credentialId) { toast(info.message || 'No biometric registered. Please register first.', 'error'); return; }

    const challenge = crypto.getRandomValues(new Uint8Array(32));
    const rawId = credentialIdToUint8Array(info.credentialId);

    await navigator.credentials.get({
      publicKey: {
        challenge,
        userVerification: 'required',
        timeout: 60000,
        allowCredentials: [{ type: 'public-key', id: rawId }]
      }
    });

    showLocBar('getting', 'Getting your location…');
    const loc = await getLocationWithAddress();
    if (loc.denied) { if (btn) { btn.innerHTML = btn._h; btn.disabled = false; } return; }
    if (loc.latitude) showLocBar('ok', loc.address || `${loc.latitude}, ${loc.longitude}`, loc.accuracy);
    else showLocBar('fail', 'Location not captured — marked without GPS');

    const deviceId = await getDeviceId();
    const att = await api({
      action: 'markEntry',
      userId: info.userId,
      loginMethod: 'biometric',
      latitude: loc.latitude,
      longitude: loc.longitude,
      address: loc.address,
      deviceId,
      guid: tenantState.guid
    });

    if (att.success) {
      toast('✓ ' + att.message, 'success');
      showAttendanceCard({ ...att, method: 'biometric' }, info.userId);
    } else if (att.code === 'TOO_FAR') {
      showLocBar('fail', `${att.distance}m from campus — must be within ${att.allowed}m`);
      toast(`📍 Too far (${att.distance}m). Move closer and try again`, 'error');
    } else {
      toast(att.message || 'Could not mark attendance', 'error');
    }
  } catch (e) {
    if (e.name === 'NotAllowedError') toast('Biometric cancelled or not allowed', 'warn');
    else if (e.name === 'InvalidStateError') toast('Biometric key not found. Please re-register.', 'error');
    else toast('Error: ' + e.message, 'error');
  } finally {
    if (btn) { btn.innerHTML = btn._h || btn.innerHTML; btn.disabled = false; }
  }
}

async function handlePasswordSignIn() {
  const email = document.getElementById('si-email').value.trim();
  const password = document.getElementById('si-password').value;
  if (!email || !password) { toast('Enter email and password first', 'error'); return; }

  const btn = document.getElementById('btn-pass-signin');
  if (btn) { btn._h = btn.innerHTML; btn.innerHTML = '<span class="spin"></span> Verifying…'; btn.disabled = true; }
  try {
    const deviceId = await getDeviceId();
    // FIX: backend uses 'signIn' not 'loginUser'
    const info = await api({ action: 'signIn', email, password, deviceId, guid: tenantState.guid });
    if (!info.success) { toast(info.message || 'Invalid credentials', 'error'); return; }

    signedInUser = info;
    showLocBar('ok', 'Password verified');

    const roleKey = normalizeRoleKey(info.roleKey || info.roleId || '');
    if (isTeacherRole(roleKey) || isAdminRole(roleKey)) {
      // Go to teacher tab
      toast('✓ Signed in as ' + (isAdminRole(roleKey) ? 'Admin' : 'Teacher'), 'success');
      return;
    }

    // Student — mark attendance
    await submitStudentAttendance('password');
  } catch (e) {
    toast('Error: ' + e.message, 'error');
  } finally {
    if (btn) { btn.innerHTML = btn._h || btn.innerHTML; btn.disabled = false; }
  }
}

// ── Mark Exit ─────────────────────────────────────────────────
async function handleExit() {
  if (!markedUserId) { toast('Mark attendance first', 'error'); return; }
  const btn = document.getElementById('btn-exit');
  if (btn) { btn.disabled = true; btn.textContent = 'Getting exit location…'; }
  try {
    showLocBar('getting', 'Getting exit location…');
    const loc = await getLocationWithAddress();
    if (loc.latitude) showLocBar('ok', 'Exit: ' + (loc.address || `${loc.latitude}, ${loc.longitude}`), loc.accuracy);
    else showLocBar('fail', 'Exit location not captured');

    const res = await api({
      action: 'markExit',
      userId: markedUserId,
      latitude: loc.latitude,
      longitude: loc.longitude,
      address: loc.address
    });

    if (res.success) {
      toast('✓ ' + res.message, 'success');
      const card = document.querySelector('.att-card');
      if (card) {
        card.querySelector('.att-name').textContent = '✓ Entry & Exit Recorded';
        card.innerHTML += `
          <div class="att-row"><span>exit_time</span><span class="att-val">${res.exitTime || ''}</span></div>
          <div class="att-row"><span>duration</span><span class="att-val">${res.duration || ''}</span></div>
          <div class="att-row"><span>address (exit)</span><span class="att-val">${res.location || 'not captured'}</span></div>`;
      }
      stopTracking();
      if (btn) { btn.disabled = true; btn.textContent = 'Exit recorded ✓'; btn.style.opacity = '.5'; }
      markedUserId = null;
      setTimeout(() => restoreSignInForm(), 4000);
    } else {
      toast(res.message, 'error');
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9"/></svg> Mark Exit (leaving classroom)';
      }
    }
  } catch (e) {
    toast('Error: ' + e.message, 'error');
    if (btn) { btn.disabled = false; btn.textContent = 'Mark Exit (leaving classroom)'; }
  }
}

async function submitStudentAttendance(loginMethod = 'biometric') {
  if (!signedInUser?.userId) { toast('Sign in first', 'error'); return; }
  const btn = document.getElementById('btn-student-attendance');
  if (btn) setLoading('btn-student-attendance', true);
  try {
    showLocBar('getting', 'Getting your location…');
    const loc = await getLocationWithAddress();
    if (loc.denied) { if (btn) setLoading('btn-student-attendance', false); return; }
    if (!loc.latitude || !loc.longitude) {
      showLocBar('fail', 'Location not captured — attendance blocked');
      toast('GPS location is required', 'error');
      if (btn) setLoading('btn-student-attendance', false);
      return;
    }
    showLocBar('ok', loc.address || `${loc.latitude}, ${loc.longitude}`, loc.accuracy);

    const att = await api({
      action: 'markEntry',
      userId: signedInUser.userId,
      loginMethod,
      latitude: loc.latitude,
      longitude: loc.longitude,
      address: loc.address
    });

    if (att.success) {
      toast('✓ ' + att.message, 'success');
      showAttendanceCard({ ...att, method: loginMethod }, signedInUser.userId);
    } else if (att.code === 'TOO_FAR') {
      showLocBar('fail', `${att.distance}m from location — must be within ${att.allowed}m`);
      toast(`📍 Too far (${att.distance}m). Move closer and try again`, 'error');
    } else {
      toast(att.message, 'error');
    }
  } catch (e) {
    toast('Error: ' + e.message, 'error');
  } finally {
    if (btn) setLoading('btn-student-attendance', false);
  }
}

// ── My Attendance ─────────────────────────────────────────────
async function toggleMyAtt() {
  const list = document.getElementById('my-att-list');
  if (list.style.display === 'block') { list.style.display = 'none'; return; }
  list.style.display = 'block';
  const uid = sessionStorage.getItem('ba_uid') || markedUserId;
  if (!uid) { list.innerHTML = '<div style="color:var(--muted);font-size:11.5px;text-align:center;padding:8px">Sign in first to view history</div>'; return; }
  list.innerHTML = '<div style="text-align:center;color:var(--muted);font-size:11.5px;padding:8px">Loading…</div>';
  try {
    const d = await api({ action: 'getMyAttendance', userId: uid });
    if (!d.records || !d.records.length) { list.innerHTML = '<div style="text-align:center;color:var(--muted);font-size:11.5px;padding:8px">No records yet</div>'; return; }
    list.innerHTML = d.records.map(r => `
      <div class="my-item">
        <div class="my-date">${r.date}</div>
        <div class="my-row"><span>entry_time</span><span>${r.entryTime || '—'}</span></div>
        <div class="my-row"><span>exit_time</span><span>${r.exitTime || '—'}</span></div>
        <div class="my-row"><span>duration</span><span>${r.duration || '—'}</span></div>
        <div class="my-row"><span>login_method</span><span>${r.loginMethod || '—'}</span></div>
        <div class="my-row"><span>address</span><span>${r.address || 'not captured'}</span></div>
        <div class="my-row"><span>distance_from_centre</span><span>${r.distanceFromCentre || '—'} m</span></div>
      </div>`).join('');
  } catch (e) { list.innerHTML = '<div style="color:var(--danger);font-size:11.5px;text-align:center;padding:8px">Error: ' + e.message + '</div>'; }
}

// ═══════════════════════════════════════════════════════════════
// REGISTER — FIX: Biometric registered in same call as account
// FIX: pubKeyCredParams includes BOTH ES256 (-7) and RS256 (-257)
// ═══════════════════════════════════════════════════════════════

// Role/dept helpers
function isValidRegistrationName(value) { return /^[A-Za-z][A-Za-z .'-]{1,79}$/.test(value); }
function isValidRegistrationEmail(value) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value); }
function isValidRegistrationMobile(value) { return /^[0-9+\-\s]{7,20}$/.test(value); }
function isValidDepartmentValue(value) { return /^[A-Za-z0-9_. -]{2,120}$/.test(String(value || '').trim()); }
function isValidDob(value) {
  if (!value) return false;
  const dob = new Date(value + 'T00:00:00');
  if (Number.isNaN(dob.getTime())) return false;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (dob >= today) return false;
  const minDob = new Date(today); minDob.setFullYear(minDob.getFullYear() - 100);
  return dob >= minDob;
}

function normalizeCode(value) { return String(value || '').replace(/\s+/g, '').toUpperCase(); }
function normalizeRoleKey(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return '';
  if (raw.includes('student')) return 'student';
  if (raw.includes('faculty') || raw.includes('teacher') || raw.includes('lecturer')) return 'teacher';
  if (raw.includes('admin')) return 'admin';
  if (raw.includes('employee') || raw.includes('staff')) return 'employee';
  return raw;
}
function roleLabelFromKey(key) {
  const raw = normalizeRoleKey(key);
  if (raw === 'teacher') return 'Faculty / Teacher';
  if (raw === 'student') return 'Student';
  if (raw === 'admin') return 'Admin';
  if (raw === 'employee') return 'Employee';
  return key || '';
}
function isAdminRole(userOrRole) { return normalizeRoleKey(userOrRole?.roleKey || userOrRole?.roleId || userOrRole?.name || userOrRole || '') === 'admin'; }
function isTeacherRole(userOrRole) { return normalizeRoleKey(userOrRole?.roleKey || userOrRole?.roleId || userOrRole?.name || userOrRole || '') === 'teacher'; }

const COLLEGE_ROLE_OPTIONS = [
  { value: '1', label: 'Admin', roleKey: 'admin' },
  { value: '2', label: 'Teacher / Faculty', roleKey: 'teacher' },
  { value: '3', label: 'Student', roleKey: 'student' },
  { value: '4', label: 'Employee', roleKey: 'employee' }
];

const DEFAULT_DEPARTMENT_OPTIONS = [
  { value: 'CSE', label: 'Computer Science & Engineering' },
  { value: 'ECE', label: 'Electronics & Communication' },
  { value: 'MECH', label: 'Mechanical Engineering' },
  { value: 'CIVIL', label: 'Civil Engineering' },
  { value: 'MBA', label: 'MBA' },
  { value: 'MCA', label: 'MCA' }
];

let registerLookupState = { roles: [], departments: [], locations: [] };
let registerFlowState = { step: 1, accountCreated: false };

function getRegisterValue(id) { const el = document.getElementById(id); return String(el?.value || '').trim(); }

function setFieldState(fieldId, message) {
  const field = document.getElementById(fieldId);
  if (!field) return;
  const error = field.querySelector('.field-error');
  if (message) {
    field.classList.add('has-error');
    if (error) error.textContent = message;
  } else {
    field.classList.remove('has-error');
    if (error) error.textContent = '';
  }
}

function clearRegisterErrors() {
  ['field-r-name', 'field-r-dob', 'field-r-email', 'field-r-mobile', 'field-r-emp-id',
    'field-r-institute', 'field-r-org-type', 'field-r-role', 'field-r-category', 'field-r-subcategory',
    'field-r-study-level', 'field-r-designation', 'field-r-password', 'field-r-confirm-password'
  ].forEach(id => setFieldState(id, ''));
}

function getRegisterRoleKey() {
  const roleEl = document.getElementById('r-role');
  if (!roleEl || roleEl.selectedIndex < 0) return '';
  const selected = roleEl.options[roleEl.selectedIndex];
  return normalizeRoleKey(selected?.dataset?.roleKey || selected?.textContent || selected?.value || '');
}

function refreshRegisterConditionalFields() {
  const roleKey = getRegisterRoleKey();
  const studyField = document.getElementById('field-r-study-level');
  const desigField = document.getElementById('field-r-designation');
  const studyInput = document.getElementById('r-study-level');
  const desigInput = document.getElementById('r-designation');

  const showStudent = roleKey === 'student';
  const showStaff = roleKey === 'teacher' || roleKey === 'employee';

  if (studyField) studyField.classList.toggle('hidden', !showStudent);
  if (studyInput) { studyInput.required = showStudent; if (!showStudent) studyInput.value = ''; }
  if (desigField) desigField.classList.toggle('hidden', !showStaff);
  if (desigInput) { desigInput.required = showStaff; if (!showStaff) desigInput.value = ''; }
}

function renderRoleOptions(options, placeholder) {
  const roleEl = document.getElementById('r-role');
  if (!roleEl) return;
  roleEl.innerHTML = `<option value="">${placeholder || 'Select role…'}</option>` +
    options.map(opt => {
      const value = opt.role_id || opt.value || '';
      const label = roleLabelFromKey(opt.label || opt.name || value);
      const roleKey = normalizeRoleKey(opt.roleKey || opt.name || opt.label || value);
      return `<option value="${value}" data-role-key="${roleKey}">${label}</option>`;
    }).join('');
}

function renderDepartmentOptions(options, placeholder) {
  const deptEl = document.getElementById('r-category');
  if (!deptEl) return;
  // Use a select if options available, else keep as text input
  if (deptEl.tagName === 'SELECT') {
    deptEl.innerHTML = `<option value="">${placeholder || 'Select category…'}</option>` +
      options.map(opt => {
        const value = opt.department_id || opt.value || '';
        const label = opt.name || opt.label || value;
        return `<option value="${value}">${label}</option>`;
      }).join('');
  }
}

function applyTenantToRegistration() {
  const tenant = window.TENANT || null;
  if (!tenant) return false;

  const orgNameInput = document.getElementById('r-institute');
  const orgTypeSelect = document.getElementById('r-org-type');
  const orgLabel = document.getElementById('r-org-name-label');
  const orgName = String(tenant.institution?.name || '').trim();
  const orgType = String(tenant.orgType || 'college').toLowerCase();

  if (orgNameInput) { orgNameInput.value = orgName; orgNameInput.readOnly = true; }
  if (orgTypeSelect) { orgTypeSelect.value = orgType; orgTypeSelect.disabled = true; }
  if (orgLabel) orgLabel.innerHTML = `Organization: ${orgName}${tenant.institution?.city ? ` (${tenant.institution.city})` : ''}`;

  const step2Title = document.getElementById('register-step-2-title');
  if (step2Title) step2Title.textContent = `${orgName || 'Organization'} · ${orgType.charAt(0).toUpperCase() + orgType.slice(1)}`;
  return true;
}

function updateRegisterFormByRole() {
  const roles = registerLookupState.roles?.length ? registerLookupState.roles : COLLEGE_ROLE_OPTIONS;
  const mapped = roles.map(r => ({
    value: r.role_id || r.value || '',
    label: r.name || r.label || r.role_id || r.value || '',
    roleKey: normalizeRoleKey(r.name || r.label || r.roleKey || r.role_id || r.value || '')
  })).filter(r => r.value);

  renderRoleOptions(mapped.length ? mapped : COLLEGE_ROLE_OPTIONS, 'Select your role…');
  refreshRegisterConditionalFields();
  applyTenantToRegistration();
}

function syncOrganizationName() {
  const orgEl = document.getElementById('r-institute');
  if (!orgEl) return;
  if (!orgEl.value) orgEl.value = '';
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
  if (!options.silent) {
    const panel = document.getElementById(`register-step-${next}`);
    const focusTarget = panel?.querySelector('input:not([disabled]):not([type="hidden"]), select:not([disabled]), button:not([disabled])');
    if (focusTarget) setTimeout(() => focusTarget.focus(), 100);
  }
}

function goRegisterStep(step) {
  const target = Math.min(Math.max(Number(step) || 1, 1), 3);
  if (target > registerFlowState.step) {
    for (let i = registerFlowState.step; i < target; i++) {
      if (!validateRegisterStep(i)) { setRegisterStep(i); return false; }
    }
  }
  setRegisterStep(target);
  return true;
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

    if (!name) { setFieldState('field-r-name', 'Name is required.'); valid = false; }
    else if (!isValidRegistrationName(name)) { setFieldState('field-r-name', 'Enter a valid full name.'); valid = false; }
    else setFieldState('field-r-name', '');

    if (!dob) { setFieldState('field-r-dob', 'Date of birth is required.'); valid = false; }
    else if (!isValidDob(dob)) { setFieldState('field-r-dob', 'Enter a valid date of birth.'); valid = false; }
    else setFieldState('field-r-dob', '');

    if (!email) { setFieldState('field-r-email', 'Email is required.'); valid = false; }
    else if (!isValidRegistrationEmail(email)) { setFieldState('field-r-email', 'Enter a valid email address.'); valid = false; }
    else setFieldState('field-r-email', '');

    if (!mobile) { setFieldState('field-r-mobile', 'Mobile number is required.'); valid = false; }
    else if (!isValidRegistrationMobile(mobile)) { setFieldState('field-r-mobile', 'Enter a valid mobile number.'); valid = false; }
    else setFieldState('field-r-mobile', '');

    if (!memberId) { setFieldState('field-r-emp-id', 'Student / Employee ID is required.'); valid = false; }
    else if (!/^\d+$/.test(memberId)) { setFieldState('field-r-emp-id', 'Use numeric digits only.'); valid = false; }
    else setFieldState('field-r-emp-id', '');
  }

  if (current === 2) {
    const role = getRegisterValue('r-role');
    const category = getRegisterValue('r-category');
    const subcategory = getRegisterValue('r-subcategory');
    const roleKey = getRegisterRoleKey();
    const studyLevel = getRegisterValue('r-study-level');
    const designation = getRegisterValue('r-designation');
    const tenantOrgName = String(window.TENANT?.institution?.name || tenantState.institution?.name || '').trim();
    const tenantOrgType = String(window.TENANT?.orgType || tenantState.orgType || '').trim();

    if (!tenantOrgName) { setFieldState('field-r-institute', 'Organization not loaded.'); valid = false; }
    else setFieldState('field-r-institute', '');
    if (!tenantOrgType) { setFieldState('field-r-org-type', 'Organization type not loaded.'); valid = false; }
    else setFieldState('field-r-org-type', '');
    if (!role) { setFieldState('field-r-role', 'Select a role.'); valid = false; }
    else setFieldState('field-r-role', '');
    if (!category) { setFieldState('field-r-category', 'Category is required.'); valid = false; }
    else if (!isValidDepartmentValue(category)) { setFieldState('field-r-category', 'Enter a valid category.'); valid = false; }
    else setFieldState('field-r-category', '');
    if (!subcategory) { setFieldState('field-r-subcategory', 'Subcategory is required.'); valid = false; }
    else setFieldState('field-r-subcategory', '');

    if (roleKey === 'student' && !studyLevel) { setFieldState('field-r-study-level', 'Select your semester.'); valid = false; }
    else setFieldState('field-r-study-level', '');
    if ((roleKey === 'teacher' || roleKey === 'employee') && !designation) { setFieldState('field-r-designation', 'Designation is required.'); valid = false; }
    else setFieldState('field-r-designation', '');
  }

  if (current === 3) {
    const pass = getRegisterValue('r-password');
    const confirmPass = getRegisterValue('r-confirm-password');
    if (!pass) { setFieldState('field-r-password', 'Password is required.'); valid = false; }
    else if (pass.length < 8) { setFieldState('field-r-password', 'Password must be at least 8 characters.'); valid = false; }
    else if (!/[A-Za-z]/.test(pass) || !/[0-9]/.test(pass)) { setFieldState('field-r-password', 'Password must include at least one letter and one number.'); valid = false; }
    else setFieldState('field-r-password', '');
    if (!confirmPass) { setFieldState('field-r-confirm-password', 'Confirm your password.'); valid = false; }
    else if (pass !== confirmPass) { setFieldState('field-r-confirm-password', 'Passwords do not match.'); valid = false; }
    else setFieldState('field-r-confirm-password', '');
  }
  return valid;
}

async function loadRegisterLookups() {
  syncOrganizationName();
  try {
    const [roleRes, deptRes, locRes] = await Promise.all([
      api({ action: 'getRoles' }),
      api({ action: 'getDepartments' }),
      api({ action: 'getLocations' })
    ]);

    const roles = (roleRes?.data || []).filter(r => r && (r.role_id || r.name));
    registerLookupState.roles = roles;
    if (tenantState) tenantState.roles = roles;

    const departments = (deptRes?.data || []).filter(d => d && d.department_id);
    registerLookupState.departments = departments;
    if (tenantState) tenantState.departments = departments;

    const locations = (locRes?.data || []).filter(l => l && l.attendance_location_id);
    registerLookupState.locations = locations;
    if (tenantState) tenantState.attendanceLocations = locations;

    updateRegisterFormByRole();
    if (departments.length) {
      const deptEl = document.getElementById('r-category');
      if (deptEl && deptEl.tagName === 'SELECT') renderDepartmentOptions(departments, 'Select category…');
    }
  } catch (e) {
    console.warn('Could not load register lookups:', e);
    renderRoleOptions(COLLEGE_ROLE_OPTIONS, 'Select your role…');
    refreshRegisterConditionalFields();
  }
}

// ── REGISTER — main function ──────────────────────────────────
// FIX: pubKeyCredParams includes BOTH -7 (ES256) and -257 (RS256)
// FIX: biometric_code saved directly during register call
async function handleRegister() { return handleRegisterV2(); }
async function handleBiometricRegister() { return handleRegisterV2(); }

async function registerBiometric(userId) {
  if (!window.PublicKeyCredential || !navigator.credentials?.create) {
    throw new Error('WebAuthn not supported on this browser/device');
  }
  const challenge = crypto.getRandomValues(new Uint8Array(32));
  const cred = await navigator.credentials.create({
    publicKey: {
      challenge,
      rp: { name: 'BioAttend', id: window.location.hostname },
      user: {
        id: new TextEncoder().encode(String(userId)),
        name: String(userId),
        displayName: 'BioAttend User'
      },
      // FIX: Include BOTH ES256 and RS256 — required by Chrome spec
      pubKeyCredParams: [
        { type: 'public-key', alg: -7 },   // ES256
        { type: 'public-key', alg: -257 }  // RS256
      ],
      authenticatorSelection: {
        authenticatorAttachment: 'platform',
        userVerification: 'required',
        requireResidentKey: false
      },
      timeout: 60000,
      attestation: 'none'
    }
  });

  if (!cred || !cred.rawId) throw new Error('Biometric registration failed — no credential returned');
  return { credentialId: bufferToBase64Url(cred.rawId) };
}

async function handleRegisterV2() {
  clearRegisterErrors();
  refreshRegisterConditionalFields();

  if (!validateRegisterStep(1)) { goRegisterStep(1); toast('Complete Step 1 first', 'error'); return; }
  if (!validateRegisterStep(2)) { goRegisterStep(2); toast('Complete Step 2 first', 'error'); return; }
  if (!validateRegisterStep(3)) { goRegisterStep(3); toast('Fix the security fields', 'error'); return; }

  const name = getRegisterValue('r-name');
  const email = getRegisterValue('r-email');
  const pass = document.getElementById('r-password').value;
  const dob = getRegisterValue('r-dob');
  const mobile = getRegisterValue('r-mobile');
  const category = getRegisterValue('r-category');
  const subcategory = getRegisterValue('r-subcategory');
  const role = document.getElementById('r-role').value;
  const inst = String(window.TENANT?.institution?.name || tenantState.institution?.name || 'Siddaganga Institute of Technology').trim();
  const orgType = String(window.TENANT?.orgType || tenantState.orgType || 'college').trim();
  const memberId = getRegisterValue('r-employee-id');
  const studyLevel = document.getElementById('r-study-level')?.value || '';
  const designation = getRegisterValue('r-designation');

  const btn = document.getElementById('btn-register');
  const bioHint = document.getElementById('bio-hint');

  if (bioHint) bioHint.textContent = 'Please use your fingerprint or Face ID when prompted…';
  setLoading('btn-register', true);

  try {
    // Step 1: Register biometric (triggers device prompt)
    if (bioHint) bioHint.textContent = '🔐 Scan your fingerprint or use Face ID…';
    const biometric = await registerBiometric(memberId || email);

    // Step 2: Get device ID
    const dId = await getDeviceId();

    // Step 3: Create account with biometric_code included
    if (bioHint) bioHint.textContent = '📡 Creating your account…';
    const d = await api({
      action: 'register',
      name, email,
      password: pass,
      dob, mobile,
      departmentId: category,
      subcategoryId: subcategory,
      roleId: role,
      instituteId: inst,
      orgType,
      studentEmployeeId: memberId,
      studyLevel,
      designation,
      biometricCode: biometric.credentialId,  // saved directly
      deviceId: dId,
      guid: tenantState.guid
    });

    if (d.success) {
      let registeredUid = d.userId;
      registerFlowState.accountCreated = true;
      if (bioHint) bioHint.textContent = '✓ Biometric registered and account created!';
      toast('✓ Account created successfully! You can now sign in.', 'success');
      // Switch to sign-in tab after 2s
      setTimeout(() => {
        switchMain('signin');
        const siEmail = document.getElementById('si-email');
        if (siEmail) siEmail.value = email;
      }, 2000);
    } else {
      if (bioHint) bioHint.textContent = 'When you tap create account, we will ask for fingerprint or Face ID before saving the account.';
      toast(d.message || 'Registration failed', 'error');
    }
  } catch (e) {
    if (bioHint) bioHint.textContent = 'When you tap create account, we will ask for fingerprint or Face ID before saving the account.';
    if (e.name === 'NotAllowedError') toast('Biometric cancelled — tap the button and scan your fingerprint/face', 'warn');
    else if (e.name === 'NotSupportedError') toast('Biometric not supported on this device. Try a different browser or device.', 'error');
    else if (e.name === 'SecurityError') toast('Security error — make sure you\'re on HTTPS or localhost', 'error');
    else if (e.name === 'InvalidStateError') toast('Biometric already registered for this ID. Try a different Student ID.', 'error');
    else toast('Error: ' + e.message, 'error');
  }
  setLoading('btn-register', false);
}

// ═══════════════════════════════════════════════════════════════
// TEACHER DASHBOARD
// FIX: Uses 'signIn' action (not 'loginUser')
// ═══════════════════════════════════════════════════════════════
let teacherData = null;
let sessionTimer = null;
let allStudents = [];
let historyData = [];

async function handleTeacherLogin() {
  const email = document.getElementById('t-email').value.trim();
  const pass = document.getElementById('t-password').value;
  if (!email || !pass) { toast('Enter email and password', 'error'); return; }
  setLoading('btn-t-login', true);
  try {
    const deviceId = await getDeviceId();
    // FIX: 'signIn' matches backend router, not 'loginUser'
    const d = await api({ action: 'signIn', email, password: pass, deviceId, guid: tenantState.guid });
    if (!d.success) { toast(d.message || 'Invalid credentials', 'error'); setLoading('btn-t-login', false); return; }

    const roleKey = normalizeRoleKey(d.roleKey || d.roleId || '');
    if (!isTeacherRole(roleKey) && !isAdminRole(roleKey)) {
      toast('Not a teacher or admin account. Students use the Sign In tab.', 'error');
      setLoading('btn-t-login', false);
      return;
    }
    teacherData = d;
    persistTeacherSession(d);
    document.getElementById('t-login-section').style.display = 'none';
    document.getElementById('t-dashboard').style.display = 'block';
    document.getElementById('t-welcome').textContent = 'Hello, ' + d.name;
    switchSub('session');
    checkActiveSess();
  } catch (e) { toast('Error: ' + e.message, 'error'); }
  setLoading('btn-t-login', false);
}

async function teacherLogout(silent = false) {
  teacherData = null;
  clearInterval(sessionTimer);
  liveSessionId = null;
  stopLivePolling();
  liveLastSyncTime = '';
  liveData = null;
  liveMapMarkers = {};
  if (liveMap) { try { liveMap.remove(); } catch (e) {} liveMap = null; }
  if (analyticsCharts.daily) { try { analyticsCharts.daily.destroy(); } catch (e) {} analyticsCharts.daily = null; }
  if (analyticsCharts.weekly) { try { analyticsCharts.weekly.destroy(); } catch (e) {} analyticsCharts.weekly = null; }
  if (!silent) {
    try { await api({ action: 'logout', guid: tenantState.guid }); } catch (e) {}
  }
  clearTeacherSession();
  const loginSection = document.getElementById('t-login-section');
  const dashboard = document.getElementById('t-dashboard');
  if (loginSection) loginSection.style.display = 'block';
  if (dashboard) dashboard.style.display = 'none';
  const emailEl = document.getElementById('t-email');
  const passEl = document.getElementById('t-password');
  if (emailEl) emailEl.value = '';
  if (passEl) passEl.value = '';
}

async function checkActiveSess() {
  try {
    const d = await api({ action: 'getActiveSession' });
    const el = document.getElementById('active-sess-display');
    const form = document.getElementById('open-sess-form');
    if (!el || !form) return;
    if (d.active) {
      let secs = d.secondsLeft;
      liveSessionId = d.session.session_id;
      el.innerHTML = `<div class="sess-card">
        <div class="sess-subj">🟢 ${d.session.subject} — LIVE</div>
        <div class="sess-meta">Closes in <span id="t-timer" style="font-weight:700;color:var(--success)">${fmtTime(secs)}</span></div>
        <button onclick="closeSess('${d.session.session_id}')" class="btn btn-danger" style="margin-top:9px;padding:8px">Stop Session</button>
      </div>`;
      form.style.display = 'none';
      clearInterval(sessionTimer);
      sessionTimer = setInterval(() => {
        secs--;
        const t = document.getElementById('t-timer');
        if (t) t.textContent = fmtTime(secs);
        if (secs <= 0) { clearInterval(sessionTimer); checkActiveSess(); }
      }, 1000);
    } else {
      liveSessionId = null;
      el.innerHTML = '';
      form.style.display = 'block';
    }
  } catch (e) { console.warn('checkActiveSess error:', e); }
}

async function openSession() {
  const subj = document.getElementById('t-subject').value.trim();
  const win = parseInt(document.getElementById('t-window').value);
  if (!subj) { toast('Enter subject name', 'error'); return; }
  setLoading('btn-open-sess', true);
  try {
    const d = await api({
      action: 'createSession',
      userId: teacherData.userId,
      teacherName: teacherData.name,
      roleId: teacherData.roleId,
      subject: subj,
      windowMinutes: win
    });
    if (d.success) {
      toast('✓ Session opened', 'success');
      document.getElementById('t-subject').value = '';
      liveLastSyncTime = '';
      checkActiveSess();
      if (document.getElementById('sp-live')?.classList.contains('active')) refreshLive(true);
    } else toast(d.message, 'error');
  } catch (e) { toast('Error: ' + e.message, 'error'); }
  setLoading('btn-open-sess', false);
}

async function closeSess(sid) {
  try {
    await api({ action: 'closeSession', sessionId: sid });
    toast('Session closed', 'success');
    clearInterval(sessionTimer);
    liveSessionId = null;
    liveLastSyncTime = '';
    checkActiveSess();
    stopLivePolling();
  } catch (e) { toast('Error', 'error'); }
}

// ── Live dashboard ────────────────────────────────────────────
function normalizeLivePayload(payload) {
  const attendance = payload?.liveAttendance || payload || {};
  return {
    sessionId: attendance.sessionId || payload?.sessionId || liveSessionId || '',
    session: attendance.session || payload?.session || null,
    totalIn: attendance.totalIn ?? (attendance.present || []).length,
    totalOut: attendance.totalOut ?? (attendance.absent || []).length,
    activeUsers: attendance.activeUsers || attendance.present || [],
    offlineUsers: attendance.offlineUsers || attendance.absent || [],
    recentlyExited: attendance.recentlyExited || [],
    locations: payload?.locations?.locations || payload?.locations || [],
    updatedAt: attendance.updatedAt || payload?.syncedAt || payload?.updatedAt || new Date().toISOString()
  };
}

async function refreshLive(force = false, internal = false) {
  if (liveRefreshInFlight) { liveRefreshQueued = true; return liveData; }
  liveRefreshInFlight = true;
  markLiveActivity();
  const listEl = document.getElementById('live-list');
  const infoEl = document.getElementById('live-info');
  const statEl = document.getElementById('live-stats');
  const toolEl = document.getElementById('live-toolbar');
  const refEl = document.getElementById('live-refresh');
  try {
    const chk = await api({ action: 'getActiveSession' });
    if (chk.active) liveSessionId = chk.session.session_id;
  } catch (e) {}
  if (!liveSessionId) {
    if (infoEl) { infoEl.className = 'no-session'; infoEl.style.display = 'block'; infoEl.textContent = 'No active session — open one from the Session tab'; }
    if (statEl) statEl.style.display = 'none';
    if (toolEl) toolEl.style.display = 'none';
    if (refEl) refEl.style.display = 'none';
    if (listEl) listEl.innerHTML = '';
    stopLivePolling();
    liveRefreshInFlight = false;
    return null;
  }
  try {
    const d = await api({ action: 'getDashboard', sessionId: liveSessionId });
    if (!d.success) throw new Error(d.message || 'Refresh failed');
    liveData = normalizeLivePayload(d);
    liveLastSyncTime = liveData.updatedAt;
    liveRetryDelay = 3000;
    if (infoEl) infoEl.style.display = 'none';
    if (statEl) statEl.style.display = 'grid';
    if (toolEl) toolEl.style.display = 'flex';
    if (refEl) refEl.style.display = 'block';
    const sp = document.getElementById('sp'), sa = document.getElementById('sa');
    const st = document.getElementById('st'), spc = document.getElementById('spc');
    if (sp) sp.textContent = liveData.activeUsers.length;
    if (sa) sa.textContent = liveData.offlineUsers.length;
    const total = liveData.activeUsers.length + liveData.offlineUsers.length;
    if (st) st.textContent = total;
    if (spc) spc.textContent = total ? Math.round(liveData.activeUsers.length / total * 100) + '%' : '0%';
    const updatedEl = document.getElementById('live-updated');
    if (updatedEl) updatedEl.textContent = 'Updated ' + new Date(liveData.updatedAt).toLocaleTimeString();
    renderLiveList();
    return liveData;
  } catch (e) {
    liveRetryDelay = Math.min(liveRetryDelay * 2, 30000);
    if (!internal) toast('Refresh error: ' + e.message, 'error');
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

function showLive(tab) {
  liveTab = tab;
  const segP = document.getElementById('seg-p'), segA = document.getElementById('seg-a'), segR = document.getElementById('seg-r');
  if (segP) segP.classList.toggle('active', tab === 'present');
  if (segA) segA.classList.toggle('active', tab === 'absent');
  if (segR) segR.classList.toggle('active', tab === 'recent');
  renderLiveList();
}

function renderLiveList() {
  if (!liveData) return;
  const el = document.getElementById('live-list');
  if (!el) return;
  let list = [];
  if (liveTab === 'present') list = liveData.activeUsers || [];
  else if (liveTab === 'absent') list = liveData.offlineUsers || [];
  else list = liveData.recentlyExited || [];
  if (!list.length) {
    el.innerHTML = `<div style="text-align:center;color:var(--muted);padding:16px;font-size:12px">${liveTab === 'present' ? 'No one present yet' : liveTab === 'absent' ? 'Everyone is present 🎉' : 'No recent exits'}</div>`;
    return;
  }
  el.innerHTML = list.map((s, i) => {
    const statusBadge = liveTab === 'present'
      ? `<span class="badge" style="background:#dcfce7;color:#166534">🟢 Present</span>`
      : liveTab === 'recent'
        ? `<span class="badge" style="background:#fef3c7;color:#92400e">🟡 Exited</span>`
        : `<span class="badge absent">🔴 Absent</span>`;
    return `<div class="att-item">
      <div style="flex:1">
        <div class="iname">${i + 1}. ${s.name}</div>
        <div class="imeta">${s.email} · ${s.category || s.department || '—'}${s.subcategory ? ' · ' + s.subcategory : ''} · ${s.entryTime || ''}${s.exitTime ? ' → exit ' + s.exitTime : ''}</div>
      </div>
      ${statusBadge}
    </div>`;
  }).join('');
}

function exportLive() {
  if (!liveData) return;
  const rows = [['full_name', 'email', 'category_id', 'subcategory_id', 'entry_time', 'exit_time', 'login_method', 'type_attendance']];
  (liveData.activeUsers || []).forEach(s => rows.push([s.name, s.email, s.category || s.department || '', s.subcategory || '', s.entryTime, s.exitTime || '', '', 'present']));
  (liveData.offlineUsers || []).forEach(s => rows.push([s.name, s.email, s.category || s.department || '', s.subcategory || '', '', '', '', 'absent']));
  dlCSV(rows, 'attendance_live.csv');
}

// ── History ───────────────────────────────────────────────────
async function loadHistory() {
  const date = document.getElementById('hist-date')?.value;
  const el = document.getElementById('hist-list');
  if (!el) return;
  el.innerHTML = '<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">Loading…</div>';
  try {
    const d = await api({ action: 'getSessions', userId: teacherData?.userId });
    let sessions = d.sessions || [];
    if (date) sessions = sessions.filter(s => s.date === date);
    historyData = sessions;
    if (!sessions.length) { el.innerHTML = '<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">No sessions found</div>'; const he = document.getElementById('hist-export'); if (he) he.style.display = 'none'; return; }
    const he = document.getElementById('hist-export');
    if (he) he.style.display = 'block';
    el.innerHTML = sessions.map(s => `
      <div class="att-item clickable" style="flex-direction:column;align-items:stretch" onclick="toggleSessDet('${s.sessionId}','${(s.subject || '').replace(/'/g, "\\'")}')">
        <div style="display:flex;justify-content:space-between;align-items:center">
          <div><div class="iname">${s.subject}</div><div class="imeta">${s.date} · ${s.startTime}–${s.endTime}</div></div>
          <div style="text-align:right"><span class="badge ${s.status === 'open' ? '' : 'closed'}">${s.presentCount || 0} present</span><div style="font-size:9.5px;margin-top:2px;color:var(--muted)">${s.status}</div></div>
        </div>
        <div id="sd-${s.sessionId}" style="display:none;margin-top:9px;border-top:1px solid var(--border);padding-top:9px"></div>
      </div>`).join('');
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

async function toggleSessDet(sid, subj) {
  const el = document.getElementById('sd-' + sid);
  if (!el) return;
  if (el.style.display === 'block') { el.style.display = 'none'; return; }
  el.style.display = 'block';
  el.innerHTML = '<div style="color:var(--muted);font-size:11px">Loading…</div>';
  try {
    const d = await api({ action: 'getDashboard', sessionId: sid });
    if (!d.success) { el.innerHTML = `<div style="color:var(--danger);font-size:11px">${d.message}</div>`; return; }
    el.innerHTML = `
      <div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px">
        <span style="color:var(--success)">✓ ${d.presentCount}</span>
        <span style="color:var(--danger)">✗ ${d.absentCount}</span>
        <span style="color:var(--muted)">${d.total} total · ${d.total ? Math.round(d.presentCount / d.total * 100) : 0}%</span>
      </div>
      ${(d.present || []).map(s => `<div style="font-size:11px;padding:3px 0;border-bottom:1px solid var(--border);color:var(--text)">${s.name} <span style="color:var(--muted)">${s.entryTime}${s.exitTime ? ' → ' + s.exitTime : ''} · ${s.method}</span></div>`).join('')}
      ${!d.present?.length ? '<div style="font-size:11px;color:var(--muted)">No students marked</div>' : ''}
      <button onclick="event.stopPropagation();exportSession('${sid}','${subj}')" class="export-btn" style="float:none;margin-top:9px">↓ Export CSV</button>`;
  } catch (e) { el.innerHTML = '<div style="color:var(--danger);font-size:11px">Error loading</div>'; }
}

async function exportSession(sid, subj) {
  try {
    const d = await api({ action: 'exportAttendance', sessionId: sid });
    if (d.success) dlCSV(null, 'att_' + (subj || 'export').replace(/\s+/g, '_') + '.csv', d.csv);
    else toast(d.message, 'error');
  } catch (e) { toast('Export failed', 'error'); }
}

async function exportHistory() {
  if (!historyData.length) return;
  try {
    const d = await api({ action: 'exportAttendance', date: historyData[0]?.date });
    if (d.success) dlCSV(null, 'att_' + (historyData[0]?.date || 'all') + '.csv', d.csv);
    else toast(d.message, 'error');
  } catch (e) { toast('Export failed', 'error'); }
}

// ── Students ──────────────────────────────────────────────────
async function loadStudents() {
  const el = document.getElementById('stud-list');
  if (!el) return;
  el.innerHTML = '<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">Loading…</div>';
  try {
    const d = await api({ action: 'getStudents' });
    allStudents = d.students || [];
    const cnt = document.getElementById('stud-count');
    if (cnt) cnt.textContent = allStudents.length + ' students';
    renderStudents(allStudents);
  } catch (e) { toast('Error', 'error'); }
}

function filterStudents() {
  const q = document.getElementById('student-search')?.value.toLowerCase() || '';
  const f = allStudents.filter(s => (s.name || '').toLowerCase().includes(q) || (s.email || '').toLowerCase().includes(q) || (s.category || s.department || '').toLowerCase().includes(q) || (s.subcategory || '').toLowerCase().includes(q));
  const cnt = document.getElementById('stud-count');
  if (cnt) cnt.textContent = f.length + ' of ' + allStudents.length + ' students';
  renderStudents(f);
}

function renderStudents(list) {
  const el = document.getElementById('stud-list');
  if (!el) return;
  if (!list.length) { el.innerHTML = '<div style="text-align:center;color:var(--muted);font-size:12px;padding:14px">No students found</div>'; return; }
  el.innerHTML = list.map((s, i) => `
    <div class="att-item">
      <div><div class="iname">${i + 1}. ${s.name}</div><div class="imeta">${s.email} · category: ${s.category || s.department || '—'}${s.subcategory ? ' · ' + s.subcategory : ''}</div></div>
      <div style="text-align:right;font-size:10px;line-height:1.8">
        <span style="color:${s.hasBio ? 'var(--success)' : 'var(--muted)'}">${s.hasBio ? '🔐 bio' : '🔐 none'}</span><br>
        <span style="color:${s.hasDevice ? 'var(--success)' : 'var(--muted)'}">${s.hasDevice ? '📱 bound' : '📱 none'}</span>
      </div>
    </div>`).join('');
}

// ── Analytics (stub — extend as needed) ──────────────────────
async function loadAnalytics(force = false) {
  const dateEl = document.getElementById('analytics-date');
  if (dateEl && !dateEl.value) dateEl.value = new Date().toISOString().slice(0, 10);
  // Basic implementation — extend with getDailyStats when backend supports it
  try {
    const d = await api({ action: 'getDashboard', sessionId: liveSessionId || 'today' });
    const sp = document.getElementById('an-present'), sa = document.getElementById('an-absent');
    const sl = document.getElementById('an-late'), sr = document.getElementById('an-rate');
    if (d.success) {
      if (sp) sp.textContent = d.presentCount || 0;
      if (sa) sa.textContent = d.absentCount || 0;
      if (sl) sl.textContent = '—';
      if (sr) sr.textContent = d.total ? Math.round(d.presentCount / d.total * 100) + '%' : '—';
    }
  } catch (e) {}
}

// ── Admin ─────────────────────────────────────────────────────
async function addDepartment() {
  const id = document.getElementById('ad-id')?.value.trim();
  const name = document.getElementById('ad-name')?.value.trim();
  const incharge = document.getElementById('ad-incharge')?.value.trim();
  const email = document.getElementById('ad-email')?.value.trim();
  if (!id || !name) { toast('category_id and name are required', 'error'); return; }
  try {
    const d = await api({ action: 'addDepartment', departmentId: id, name, inCharge: incharge, email });
    if (d.success) {
      toast('✓ Category added', 'success');
      const eid = document.getElementById('ad-id'), ename = document.getElementById('ad-name');
      if (eid) eid.value = ''; if (ename) ename.value = '';
      loadDepts();
    } else toast(d.message, 'error');
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

async function loadDepts() {
  const el = document.getElementById('dept-list');
  if (!el) return;
  try {
    const d = await api({ action: 'getDepartments' });
    const rows = d.data || [];
    el.innerHTML = rows.length
      ? rows.map(r => `<div class="att-item"><div><div class="iname">${r.department_id} — ${r.name}</div><div class="imeta">in_charge: ${r.in_charge || '—'} · ${r.email || '—'}</div></div></div>`).join('')
      : '<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">No categories yet</div>';
  } catch (e) {}
}

async function addLocation() {
  const name = document.getElementById('al-name')?.value.trim();
  const address = document.getElementById('al-address')?.value.trim() || '';
  const lat = document.getElementById('al-lat')?.value.trim();
  const lng = document.getElementById('al-lng')?.value.trim();
  const radius = document.getElementById('al-radius')?.value.trim() || '200';
  const reuseLocationId = document.getElementById('al-reuse')?.value.trim() || '';
  if (!name || !lat || !lng) { toast('Location name, latitude and longitude are required', 'error'); return; }
  if (reuseLocationId && !/^\d+$/.test(reuseLocationId)) { toast('Location ID must be numeric', 'error'); return; }
  try {
    const payload = {
      action: 'addAttendanceLocation',
      name,
      address,
      latitude: parseFloat(lat),
      longitude: parseFloat(lng),
      geofenceRadius: parseInt(radius) || 200,
      reuseLocationId
    };
    const d = await api(payload);
    if (d.success) {
      toast(d.reused ? 'Existing location reused' : 'Location added', 'success');
      const fields = ['al-name','al-address','al-lat','al-lng','al-radius','al-reuse'];
      fields.forEach(id => { const el = document.getElementById(id); if (el) el.value = id === 'al-radius' ? '200' : ''; });
      await loadLocs();
    } else toast(d.message, 'error');
    if (d.duplicateWarning && d.nearbyLocations && d.nearbyLocations.length) {
      const reuse = window.confirm(`A location already exists within 20 meters: ${d.nearbyLocations[0].name || d.nearbyLocations[0].attendance_location_id}. Reuse it?`);
      const confirmPayload = reuse
        ? { ...payload, reuseLocationId: d.nearbyLocations[0].attendance_location_id }
        : { ...payload, confirmDuplicate: true };
      const follow = await api(confirmPayload);
      if (follow.success) { toast('Location saved', 'success'); await loadLocs(); }
      else toast(follow.message, 'error');
    }
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

async function loadLocs() {
  const el = document.getElementById('loc-list');
  if (!el) return;
  try {
    const d = await api({ action: 'getLocations' });
    const rows = d.data || [];
    el.innerHTML = rows.length
      ? rows.map(r => `<div class="att-item"><div><div class="iname">${r.attendance_location_id} — ${r.name}</div><div class="imeta">${r.address || '-'} · lat: ${r.latitude} · lng: ${r.longitude} · radius: ${r.geofence_radius || 200}m</div></div></div>`).join('')
      : '<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">No locations yet</div>';
  } catch (e) {}
}

async function addUserLocMap() {
  const category = document.getElementById('ulm-category')?.value.trim() || document.getElementById('admin-map-category')?.value.trim();
  const subcategory = document.getElementById('ulm-subcategory')?.value.trim() || document.getElementById('admin-map-subcategory')?.value.trim();
  const lid = document.getElementById('ulm-lid')?.value.trim() || document.getElementById('admin-map-location')?.value.trim();
  const dist = document.getElementById('ulm-dist')?.value.trim() || document.getElementById('admin-map-distance')?.value.trim();
  if (!category || !subcategory || !lid) { toast('category, subcategory and attendance location are required', 'error'); return; }
  if (!/^\d+$/.test(lid)) { toast('attendance location id must be numeric', 'error'); return; }
  try {
    const d = await api({ action: 'addCategoryLocationMap', categoryId: category, subcategoryId: subcategory, locationId: lid, allowedDistance: parseInt(dist) || 200 });
    if (d.success) { toast('✓ Mapping added', 'success'); }
    else toast(d.message, 'error');
  } catch (e) { toast('Error: ' + e.message, 'error'); }
}

// ── CSV download ──────────────────────────────────────────────
function dlCSV(rows, filename, csvStr) {
  const c = csvStr || rows.map(r => r.map(x => '"' + String(x || '').replace(/"/g, '""') + '"').join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([c], { type: 'text/csv' }));
  a.download = filename;
  a.click();
}

// ── Init ──────────────────────────────────────────────────────
// Restore session if student already marked today
try {
  const uid = sessionStorage.getItem('ba_uid');
  if (uid) {
    markedUserId = uid;
    showAttendanceCard({
      name: sessionStorage.getItem('ba_name'),
      date: sessionStorage.getItem('ba_date'),
      time: sessionStorage.getItem('ba_time'),
      location: sessionStorage.getItem('ba_loc'),
      method: sessionStorage.getItem('ba_meth'),
      distanceFromCentre: sessionStorage.getItem('ba_dist')
    }, uid);
  }
} catch (e) {}

// Online/offline bar
const offBar = document.getElementById('offline-bar');
function chkOnline() { if (offBar) offBar.classList.toggle('show', !navigator.onLine); }
window.addEventListener('online', chkOnline);
window.addEventListener('offline', chkOnline);
chkOnline();

// Set DOB max
const rDob = document.getElementById('r-dob');
if (rDob) rDob.max = new Date().toISOString().slice(0, 10);

// Restore teacher session
try {
  const savedTeacher = localStorage.getItem('ba_teacher_session');
  if (savedTeacher) {
    teacherData = JSON.parse(savedTeacher);
    if (teacherData?.authToken) scheduleAuthExpiry();
  }
} catch (e) {}

// Boot tenant (never fails — always uses fallback)
(async () => {
  const ok = await bootTenant();
  if (ok !== false) applyTenantToRegistration();
})();
