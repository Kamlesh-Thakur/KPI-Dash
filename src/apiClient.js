const API_BASE = import.meta.env.VITE_API_BASE_URL || 'http://localhost:8787/api';

let accessToken = '';
let refreshToken = '';

function setAuthHeaders(headers = {}) {
  if (accessToken) headers.Authorization = `Bearer ${accessToken}`;
  return headers;
}

export function setTokens(nextAccess, nextRefresh) {
  accessToken = nextAccess || '';
  refreshToken = nextRefresh || '';
}

export function getTokens() {
  return { accessToken, refreshToken };
}

export async function apiFetch(path, options = {}) {
  const headers = setAuthHeaders({
    'Content-Type': 'application/json',
    ...(options.headers || {})
  });
  const res = await fetch(`${API_BASE}${path}`, { ...options, headers });
  if (!res.ok) {
    const payload = await res.json().catch(() => ({}));
    throw new Error(payload.error || `API error ${res.status}`);
  }
  return res.status === 204 ? null : res.json();
}

export async function login(email, password) {
  const payload = await apiFetch('/auth/login', {
    method: 'POST',
    body: JSON.stringify({ email, password })
  });
  setTokens(payload.accessToken, payload.refreshToken);
  return payload;
}

export async function refreshAccessToken() {
  if (!refreshToken) return null;
  const payload = await apiFetch('/auth/refresh', {
    method: 'POST',
    body: JSON.stringify({ refreshToken }),
    headers: { Authorization: '' }
  });
  setTokens(payload.accessToken, payload.refreshToken);
  return payload;
}

export async function getMe() {
  return apiFetch('/auth/me');
}

export async function loadApiDataset() {
  const [raw, incident] = await Promise.all([
    apiFetch('/kpi/rows?dataset=raw_data'),
    apiFetch('/kpi/rows?dataset=incident')
  ]);
  return {
    rawData: raw.rows || [],
    incidentData: incident.rows || [],
    branchEfficiency: [],
    branchMapping: [],
    dropdowns: [],
    overallKpi: null
  };
}

export async function uploadWorkbookFile(file) {
  const form = new FormData();
  form.append('file', file);
  const headers = setAuthHeaders({});
  const res = await fetch(`${API_BASE}/uploads/workbook`, {
    method: 'POST',
    headers,
    body: form
  });
  if (!res.ok) {
    const payload = await res.json().catch(() => ({}));
    throw new Error(payload.error || `Upload failed ${res.status}`);
  }
  return res.json();
}

