const AUTH_STORAGE_KEY = 'kpi-auth-v1';

const AUTH_STATE = {
  user: null
};

export function getAuthState() {
  return AUTH_STATE;
}

export function setAuthUser(user) {
  AUTH_STATE.user = user || null;
}

export function canUploadData() {
  const roles = AUTH_STATE.user?.roles || [];
  return roles.includes('admin') || roles.includes('analyst');
}

export function saveAuthSession(session) {
  localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(session || {}));
}

export function readAuthSession() {
  try {
    const raw = localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (_err) {
    return null;
  }
}

export function clearAuthSession() {
  localStorage.removeItem(AUTH_STORAGE_KEY);
}

