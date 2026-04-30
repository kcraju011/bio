
const DEFAULT_TENANT_API = 'https://script.google.com/macros/s/AKfycbyfgsjU607novJJlhwZfMdlSreCGP7OLaaj6ztikWQb4VawkisqPGLwdqDkDuqYfjQlZw/exec';

const NERVE_URL = 'https://script.google.com/macros/s/AKfycbx4Ef8qNz71xgGYE7jGxV3C7yO29q97zLnGPPKLlCp_0A-HdyIHltxat9hYzzWs37u5hw/exec';

const TENANT_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
const EXPECTED_TENANTS = {
  '1': {
    name: 'SIT',
    apiUrl: 'https://script.google.com/macros/s/AKfycbyfgsjU607novJJlhwZfMdlSreCGP7OLaaj6ztikWQb4VawkisqPGLwdqDkDuqYfjQlZw/exec'
  },         
  '2': {
    name: 'SSIT',
    apiUrl: 'https://script.google.com/macros/s/AKfycbzQ1Abzar9Ydny0mZyw4JMMOrxs868Flpdb7vHdHhltbYHdRIVQYl-aTsUQ6uP0zei_/exec'
  }
};


// â”€â”€ FALLBACK TENANTS (OFFLINE MODE) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const FALLBACK_GUIDS = new Set(['1', '2']);

const FALLBACK_TENANTS = {

  '1': {
    success: true,
    guid: '1',
    orgType: 'college',
    institution: {
      name: 'SIT',
      city: 'Tumakuru',
      logoUrl: 'https://web.sit.ac.in/wp-content/uploads/2025/03/SIT-Logo-1.png',
      website: 'https://web.sit.ac.in/',
      address: 'Siddaganga Institute of Technology'
    },
    application: {
      id: 101,
      name: 'Attendance monitoring',
      description: 'Biometric attendance for students, teachers and employees'
    },
    apiUrl: 'https://script.google.com/macros/s/AKfycbyfgsjU607novJJlhwZfMdlSreCGP7OLaaj6ztikWQb4VawkisqPGLwdqDkDuqYfjQlZw/exec'
  },

  '2': {
    success: true,
    guid: '2',
    orgType: 'college',
    institution: {
      name: 'SSIT',
      city: 'Tumakuru',
      logoUrl: 'https://ssit.edu.in/img/ssit-logo.png',
      website: 'https://ssit.edu.in/',
      address: 'Sri Siddhartha Institute of Technology'
    },
    application: {
      id: 102,
      name: 'Attendance monitoring',
      description: 'Biometric attendance for students, teachers and employees'
    },
    apiUrl: 'https://script.google.com/macros/s/AKfycbzQ1Abzar9Ydny0mZyw4JMMOrxs868Flpdb7vHdHhltbYHdRIVQYl-aTsUQ6uP0zei_/exec'
  },

};

// â”€â”€ State â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
let TENANT_API    = DEFAULT_TENANT_API;
let tenantState   = {
  guid: '',
  orgType: '',
  institution: {},
  application: {},
  roles: [],
  departments: [],
  attendanceLocations: []
};

// â”€â”€ API â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function readGuidFromUrl() {
  const params = new URLSearchParams(window.location.search);
  return String(params.get('q') || '').trim();
}

function tenantCacheKey(guid) {
  return `tenant_${guid}`;
}

function resolveTenantApiUrl() {
  const guid = String(tenantState.guid || readGuidFromUrl() || '').trim();
  if (guid && EXPECTED_TENANTS[guid]?.apiUrl) return EXPECTED_TENANTS[guid].apiUrl;
  if (TENANT_API) return TENANT_API;
  return DEFAULT_TENANT_API;
}

function getTenantApiCandidates(extraUrl) {
  const seen = new Set();
  const urls = [];
  [extraUrl, resolveTenantApiUrl(), TENANT_API, DEFAULT_TENANT_API, EXPECTED_TENANTS[String(tenantState.guid || readGuidFromUrl() || '').trim()]?.apiUrl]
    .forEach(url => {
      const normalized = String(url || '').trim();
      if (!normalized || seen.has(normalized)) return;
      seen.add(normalized);
      urls.push(normalized);
    });
  return urls;
}

async function jsonpRequestWithFallback(urls, timeoutMs = 45000) {
  let lastError = null;
  for (const url of urls) {
    try {
      return await jsonpRequest(url, timeoutMs);
    } catch (err) {
      lastError = err;
      console.warn('[tenant] request failed, trying next url:', url, err);
    }
  }
  throw lastError || new Error('Failed to load tenant response');
}

function credentialIdToUint8Array(value) {
  const raw = String(value || '').trim();
  if (!raw) return new Uint8Array();
  const normalized = raw.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized + '='.repeat((4 - (normalized.length % 4 || 4)) % 4);
  return Uint8Array.from(atob(padded), c => c.charCodeAt(0));
}

function setTenantLoading(on, text) {
  const box = document.getElementById('tenant-loader');
  const label = document.getElementById('tenant-loader-text');
  if (!box) return;
  if (label && text) label.textContent = text;
  box.style.display = on ? 'flex' : 'none';
}

function jsonpRequest(url, timeoutMs = 45000) {
  return new Promise((resolve, reject) => {
    const cb = '__ba_jsonp_' + Math.random().toString(36).slice(2) + Date.now().toString(36);
    const script = document.createElement('script');
    const sep = url.includes('?') ? '&' : '?';
    let timer = null;

    function cleanup(options = {}) {
      const preserveCallback = !!options.preserveCallback;
      if (timer) clearTimeout(timer);
      if (script.parentNode) script.parentNode.removeChild(script);
      if (preserveCallback) {
        window[cb] = () => {};
      } else {
        try { delete window[cb]; } catch (e) { window[cb] = undefined; }
      }
    }

    window[cb] = data => {
      cleanup();
      resolve(data);
    };

    script.async = true;
    script.onerror = () => {
      cleanup({ preserveCallback: true });
      reject(new Error('Failed to load tenant response'));
    };
    script.src = `${url}${sep}callback=${encodeURIComponent(cb)}`;
    document.head.appendChild(script);

    timer = setTimeout(() => {
      cleanup({ preserveCallback: true });
      reject(new Error('Tenant request timed out'));
    }, timeoutMs);
  });
}

function getAuthToken() {
  if (window.__ba_auth_token) return window.__ba_auth_token;
  if (teacherData?.authToken) return teacherData.authToken;
  try {
    return localStorage.getItem('ba_auth_token') || '';
  } catch (e) {
    return '';
  }
}

function getRefreshToken() {
  if (window.__ba_refresh_token) return window.__ba_refresh_token;
  if (teacherData?.refreshToken) return teacherData.refreshToken;
  try {
    return localStorage.getItem('ba_refresh_token') || '';
  } catch (e) {
    return '';
  }
}

function getAuthExpiresAt() {
  try {
    return localStorage.getItem('ba_auth_expires_at') || '';
  } catch (e) {
    return '';
  }
}

function isAuthExpired() {
  const expiresAt = getAuthExpiresAt();
  if (!expiresAt) return false;
  const ms = Date.parse(expiresAt);
  return !Number.isNaN(ms) && Date.now() >= ms;
}

function scheduleAuthExpiry() {
  try {
    if (window.__ba_auth_expiry_timer) {
      clearTimeout(window.__ba_auth_expiry_timer);
      window.__ba_auth_expiry_timer = null;
    }
    const expiresAt = getAuthExpiresAt();
    if (!expiresAt) return;
    const ms = Date.parse(expiresAt);
    if (Number.isNaN(ms)) return;
    const delay = ms - Date.now();
    if (delay <= 0) {
      handleExpiredAuth();
      return;
    }
    window.__ba_auth_expiry_timer = setTimeout(() => {
      handleExpiredAuth();
    }, delay);
  } catch (e) {}
}

function handleExpiredAuth() {
  clearTeacherSession();
  try {
    if (typeof signedInUser !== 'undefined') signedInUser = null;
    if (typeof markedUserId !== 'undefined') markedUserId = null;
    if (typeof stopTracking === 'function') stopTracking();
  } catch (e) {}
  if (typeof teacherLogout === 'function') {
    try { teacherLogout(true); } catch (e) {}
  }
  if (typeof restoreSignInForm === 'function') {
    try { restoreSignInForm(); } catch (e) {}
  }
}

function persistTeacherSession(data) {
  if (!data?.authToken) return;
  try {
    window.__ba_auth_token = data.authToken;
    window.__ba_refresh_token = data.refreshToken || '';
    localStorage.setItem('ba_auth_token', data.authToken);
    if (data.refreshToken) localStorage.setItem('ba_refresh_token', data.refreshToken);
    if (data.authExpiresAt) localStorage.setItem('ba_auth_expires_at', data.authExpiresAt);
    if (data.refreshExpiresAt) localStorage.setItem('ba_refresh_expires_at', data.refreshExpiresAt);
    if (data.userId) localStorage.setItem('ba_auth_user_id', String(data.userId));
    if (data.roleId) localStorage.setItem('ba_auth_role', String(data.roleId));
    if (data.guid) localStorage.setItem('ba_auth_guid', String(data.guid));
    if (data.deviceId) localStorage.setItem('ba_auth_device_id', String(data.deviceId));
    localStorage.setItem('ba_teacher_session', JSON.stringify(data));
    scheduleAuthExpiry();
  } catch (e) {}
}

function clearTeacherSession() {
  try {
    window.__ba_auth_token = '';
    window.__ba_refresh_token = '';
    localStorage.removeItem('ba_auth_token');
    localStorage.removeItem('ba_refresh_token');
    localStorage.removeItem('ba_auth_expires_at');
    localStorage.removeItem('ba_refresh_expires_at');
    localStorage.removeItem('ba_auth_user_id');
    localStorage.removeItem('ba_auth_role');
    localStorage.removeItem('ba_auth_guid');
    localStorage.removeItem('ba_auth_device_id');
    localStorage.removeItem('ba_teacher_session');
    if (window.__ba_auth_expiry_timer) {
      clearTimeout(window.__ba_auth_expiry_timer);
      window.__ba_auth_expiry_timer = null;
    }
    if (typeof teacherData !== 'undefined') teacherData = null;
    if (typeof signedInUser !== 'undefined') signedInUser = null;
  } catch (e) {}
}

function authExpirySoon() {
  const expiresAt = getAuthExpiresAt();
  if (!expiresAt) return false;
  const ms = Date.parse(expiresAt);
  if (Number.isNaN(ms)) return false;
  return (ms - Date.now()) < 60 * 1000;
}

async function refreshAccessToken() {
  const refreshToken = getRefreshToken();
  if (!refreshToken) return null;
  const payload = {
    action: 'refreshToken',
    refreshToken,
    guid: tenantState.guid
  };
  const query = '?data=' + encodeURIComponent(JSON.stringify(payload));
  const res = await jsonpRequestWithFallback(getTenantApiCandidates().map(url => url + query));
  if (res && res.success && res.authToken) {
    persistTeacherSession(res);
    if (typeof teacherData !== 'undefined' && teacherData) teacherData = { ...teacherData, ...res };
    if (typeof signedInUser !== 'undefined' && signedInUser) signedInUser = { ...signedInUser, ...res };
    return res;
  }
  return null;
}

async function api(payload) {
  const apiUrl = resolveTenantApiUrl();
  if (!apiUrl) throw new Error('Tenant API is not initialized');
  if (isAuthExpired()) {
    const refreshed = await refreshAccessToken().catch(() => null);
    if (!refreshed) {
      handleExpiredAuth();
      throw new Error('Session expired');
    }
  } else if (authExpirySoon()) {
    await refreshAccessToken().catch(() => null);
  }
  const requestPayload = { ...(payload || {}) };
  if (!requestPayload.guid && tenantState.guid) requestPayload.guid = tenantState.guid;
  if (!requestPayload.authToken) {
    const authToken = getAuthToken();
    if (authToken) requestPayload.authToken = authToken;
  }
  const query = '?data=' + encodeURIComponent(JSON.stringify(requestPayload));
  console.log('TENNANT API request sent');
  let response = await jsonpRequestWithFallback(getTenantApiCandidates(apiUrl).map(url => url + query));
  if (response && response.success === false) {
    const message = String(response.message || '').toLowerCase();
    const code = String(response.code || '').toUpperCase();
    if (code === 'TOKEN_EXPIRED' || message.includes('session expired')) {
      const refreshed = await refreshAccessToken().catch(() => null);
      if (refreshed) {
        requestPayload.authToken = getAuthToken();
        const retryQuery = '?data=' + encodeURIComponent(JSON.stringify(requestPayload));
        response = await jsonpRequestWithFallback(getTenantApiCandidates(apiUrl).map(url => url + retryQuery));
      } else {
        handleExpiredAuth();
      }
    } else if (message.includes('unauthorized') || message.includes('authentication required')) {
      handleExpiredAuth();
    }
  }
  return response;
}

// â”€â”€ Toast â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function toast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'show ' + type;
  clearTimeout(t._t); t._t = setTimeout(() => t.className = '', 3500);
}

function setLoading(id, on) {
  const b = document.getElementById(id); if (!b) return;
  if (on) { b._h = b.innerHTML; b.innerHTML = '<span class="spin"></span> Please waitâ€¦'; b.disabled = true; }
  else { b.innerHTML = b._h || b.innerHTML; b.disabled = false; }
}

function togglePw(id) { const i = document.getElementById(id); i.type = i.type === 'password' ? 'text' : 'password'; }
function fmtTime(s) { const m = Math.floor(Math.max(0,s)/60), x = Math.max(0,s)%60; return String(m).padStart(2,'0')+':'+String(x).padStart(2,'0'); }

function getFallbackTenantProfile(guid) {
  const normalized = String(guid || '').trim();
  return FALLBACK_GUIDS.has(normalized) ? FALLBACK_TENANTS[normalized] || null : null;
}

function normalizeTenantName(value) {
  return String(value || '').trim().toLowerCase();
}

function isExpectedTenantProfile(guid, profile) {
  const expected = EXPECTED_TENANTS[String(guid || '').trim()];
  if (!profile) return false;
  if (!expected) return true;
  const profileName = normalizeTenantName(profile?.institution?.name || profile?.name || '');
  const expectedName = normalizeTenantName(expected.name);
  const profileApi = String(profile?.apiUrl || '').trim();
  const expectedApi = String(expected.apiUrl || '').trim();
  if (profileName && profileName !== expectedName) return false;
  if (profileApi && profileApi !== expectedApi) return false;
  return true;
}

function getDefaultTenantProfile(guid) {
  const fallback = getFallbackTenantProfile(guid);
  if (fallback) return fallback;
  return null;
}

function readCachedTenant(guid) {
  try {
    const raw = localStorage.getItem(tenantCacheKey(guid));
    if (!raw) return null;
    const cached = JSON.parse(raw);
    if (!cached?.guid || String(cached.guid) !== String(guid)) return null;
    if (!isExpectedTenantProfile(guid, cached.data)) {
      localStorage.removeItem(tenantCacheKey(guid));
      return null;
    }
    if (!cached?.expiresAt || Date.now() > cached.expiresAt) {
      localStorage.removeItem(tenantCacheKey(guid));
      return null;
    }
    return cached.data || null;
  } catch (e) {
    return null;
  }
}

function cacheTenantProfile(guid, data) {
  try {
    localStorage.setItem(tenantCacheKey(guid), JSON.stringify({
      guid: String(guid),
      expiresAt: Date.now() + TENANT_CACHE_TTL_MS,
      data
    }));
  } catch (e) {}
}

async function fetchTenantProfile(guid) {
  const normalizedGuid = String(guid || '').trim();
  if (!normalizedGuid) return null;
  if (!NERVE_URL || NERVE_URL.includes('/dev')) {
    return getDefaultTenantProfile(guid);
  }
  if (NERVE_URL) {
    try {
      const joiner = NERVE_URL.includes('?') ? '&' : '?';
      console.log('[tenant] fetching Nerve profile for GUID:', guid);
      const payload = await jsonpRequest(`${NERVE_URL}${joiner}getApplicationFromGuid=${encodeURIComponent(guid)}`);
      console.log('[tenant] Nerve response:', payload);
      if (payload?.success && String(payload.guid || guid) === String(guid) && payload?.apiUrl && isExpectedTenantProfile(guid, payload)) return payload;
    } catch (e) {
      console.warn('[tenant] Nerve lookup failed:', e);
    }
  }
  return getDefaultTenantProfile(guid);
}

function applyTenantBranding(profile) {
  const institution = profile?.institution || {};
  const application = profile?.application || {};
  tenantState.guid = String(profile?.guid || tenantState.guid || readGuidFromUrl());
  tenantState.orgType = String(profile?.orgType || institution?.orgType || '').toLowerCase();
  tenantState.institution = institution;
  tenantState.application = application;
  TENANT_API = profile?.apiUrl || EXPECTED_TENANTS[String(profile?.guid || tenantState.guid || '').trim()]?.apiUrl || DEFAULT_TENANT_API;
  window.TENANT_API = TENANT_API;
  window.TENANT = profile || null;

  const orgName = institution.name || 'BioAttend';
  const subtitle = application.description || 'Biometric Attendance System';
  const city = institution.city || tenantState.orgType || 'Organization';
  const logo = institution.logoUrl || '';

  document.title = `${orgName} | BioAttend`;
  const brandEl = document.getElementById('org-name');
  const tagEl = document.getElementById('org-tagline');
  const cityEl = document.getElementById('org-city-text');
  const teacherSubtitle = document.getElementById('teacher-dashboard-subtitle');
  const logoEl = document.getElementById('org-logo');
  const fallbackEl = document.getElementById('org-logo-fallback');

  if (brandEl) brandEl.textContent = orgName;
  if (tagEl) tagEl.textContent = subtitle;
  if (cityEl) cityEl.textContent = city;
  if (teacherSubtitle) teacherSubtitle.textContent = `Teacher Dashboard · ${orgName}`;
  
  if (logoEl && fallbackEl) {
    logoEl.onload = () => {
      logoEl.style.display = 'block';
      fallbackEl.style.display = 'none';
    };
    logoEl.onerror = () => {
      logoEl.removeAttribute('src');
      logoEl.style.display = 'none';
      fallbackEl.style.display = 'block';
    };
    if (logo) {
      logoEl.referrerPolicy = 'no-referrer';
      logoEl.loading = 'eager';
      logoEl.src = logo;
      logoEl.style.display = 'block';
      fallbackEl.style.display = 'none';
    } else {
      logoEl.removeAttribute('src');
      logoEl.style.display = 'none';
      fallbackEl.style.display = 'block';
    }
  }
}

async function bootTenant() {
  const guid = readGuidFromUrl();
  tenantState.guid = guid;
  setTenantLoading(true, 'Loading organization...');
  let invalidLink = false;
  try {
    console.info('[tenant] boot start', { guid });
    if (!guid) {
      invalidLink = true;
      throw new Error('Invalid tenant link');
    }
    try { localStorage.removeItem('tenant'); } catch (e) {}
    const tenantProfile = await fetchTenantProfile(guid);
    if (!tenantProfile?.apiUrl && !tenantProfile?.institution) throw new Error('Tenant profile is incomplete');
    applyTenantBranding({ ...tenantProfile, guid });
    cacheTenantProfile(guid, { ...tenantProfile, guid });
    await loadRegisterLookups();
  } catch (e) {
    console.warn('[tenant] boot fallback', e);
    if (invalidLink) {
      const box = document.getElementById('tenant-loader');
      const label = document.getElementById('tenant-loader-text');
      if (label) label.textContent = 'Invalid tenant link';
      if (box) box.innerHTML = '<strong>Invalid link</strong><span>Please use a valid tenant URL with a tenant GUID.</span>';
      return;
    }
    const cachedProfile = readCachedTenant(guid);
    const fallbackProfile = cachedProfile || getFallbackTenantProfile(guid);
    if (!fallbackProfile) {
      setTenantLoading(true, 'Invalid tenant link');
      return;
    }
    applyTenantBranding({ ...fallbackProfile, guid });
    toast('Using fallback tenant configuration', 'warn');
    await loadRegisterLookups();
  } finally {
    if (!invalidLink) setTenantLoading(false);
  }
}

// â”€â”€ Tab navigation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

