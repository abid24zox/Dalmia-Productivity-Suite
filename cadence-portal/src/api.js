// Thin HTTP client to the shared Cadence service. Every portal read/write goes
// through here, so the portal and the Teams agent operate on ONE dataset.
// Empty = same-origin: requests go to /api/* on whatever host served the app
// (localhost or an ngrok tunnel) and Vite's dev proxy forwards them to the
// Cadence service. Set VITE_CADENCE_API_URL only to point at a remote service.
const BASE = import.meta.env.VITE_CADENCE_API_URL || '';
const KEY = import.meta.env.VITE_CADENCE_API_KEY || '';

async function call(path, init = {}) {
  // 'ngrok-skip-browser-warning' stops ngrok's free-tier interstitial from ever
  // intercepting an API fetch when the app is shared through a tunnel.
  const headers = { 'content-type': 'application/json', 'ngrok-skip-browser-warning': '1', ...(init.headers || {}) };
  if (KEY) headers['x-api-key'] = KEY;
  const res = await fetch(`${BASE}${path}`, { ...init, headers });
  const text = await res.text();
  let body = {};
  try { body = text ? JSON.parse(text) : {}; } catch { body = { raw: text }; }
  if (!res.ok) { const err = new Error(body.error || `Cadence API ${res.status} on ${path}`); err.status = res.status; err.notConfigured = !!body.notConfigured; throw err; }
  return body;
}

const post = (path, body) => call(path, { method: 'POST', body: JSON.stringify(body || {}) });
const patch = (path, body) => call(path, { method: 'PATCH', body: JSON.stringify(body || {}) });
const del = (path) => call(path, { method: 'DELETE' });

export const api = {
  base: BASE,
  health: () => call('/api/health'),
  login: (username, password) => post('/api/login', { username, password }).then((r) => r.user),
  snapshot: () => call('/api/snapshot'),

  // writes — each returns { snapshot, ... }
  addWorks: (works) => post('/api/works', { works }),
  patchWork: (id, p) => patch(`/api/works/${id}`, p),
  deleteWork: (id) => del(`/api/works/${id}`),
  addActs: (activities) => post('/api/activities', { activities }),
  patchAct: (id, p) => patch(`/api/activities/${id}`, p),
  deleteAct: (id) => del(`/api/activities/${id}`),
  addCr: (cr) => post('/api/crs', cr),
  patchCr: (id, p) => patch(`/api/crs/${id}`, p),
  decideApproval: (id, body) => post(`/api/approvals/${id}/decide`, body),
  addTeam: (team) => post('/api/teams', team),
  patchTeam: (id, p) => patch(`/api/teams/${id}`, p),
  addRemark: (remark) => post('/api/remarks', remark),
  markRemarksRead: (userId) => post('/api/remarks/read', { userId }),

  // AI (shared Foundry model) + documents + speech
  aiComplete: (system, user) => post('/api/ai/complete', { system, user }).then((r) => r.text),
  aiExtract: (data, name) => post('/api/ai/extract', { data, name }).then((r) => r.text),
  // Fetch a OneDrive/SharePoint share link server-side → { text, name }.
  aiFetchUrl: (url) => post('/api/ai/fetch-url', { url }),
  // Proxy a pre-authenticated download URL → { name, dataB64 } (CORS fallback).
  aiFetchDownload: (url, name) => post('/api/ai/fetch-download', { url, name }),
  aiTranscribe: (audio, mimetype) => post('/api/ai/transcribe', { audio, mimetype }).then((r) => r.text),
};
