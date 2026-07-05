// Thin HTTP client to the shared Cadence service. All agent tools go through
// here, so the bot never touches state directly — Teams actions become Cadence
// API calls, which the portal then reads.
// In the merged single-App-Service deploy the API lives in THIS process, so
// default to our own port (App Service sets PORT). Standalone bot dev overrides
// with CADENCE_API_URL pointing at the separately-running service.
const BASE = process.env.CADENCE_API_URL || `http://localhost:${process.env.PORT || 4000}`;
const KEY = process.env.CADENCE_API_KEY || '';

async function call(path: string, init: RequestInit = {}): Promise<any> {
  const headers: Record<string, string> = { 'content-type': 'application/json', ...(init.headers as any) };
  if (KEY) headers['x-api-key'] = KEY;
  const res = await fetch(`${BASE}${path}`, { ...init, headers });
  const text = await res.text();
  let body: any = {};
  try { body = text ? JSON.parse(text) : {}; } catch { body = { raw: text }; }
  if (!res.ok) throw new Error(body.error || `Cadence API ${res.status} on ${path}`);
  return body;
}

export const cadence = {
  health: () => call('/api/health'),
  users: () => call('/api/users'),
  teams: () => call('/api/teams'),
  portfolio: (userId: string) => call(`/api/portfolio?userId=${encodeURIComponent(userId)}`),
  initiative: (idOrTitle: string) => call(`/api/initiatives/${encodeURIComponent(idOrTitle)}`),
  createInitiative: (payload: any) => call('/api/initiatives', { method: 'POST', body: JSON.stringify(payload) }),
  scheduleActivity: (id: string, body: any) => call(`/api/activities/${id}/schedule`, { method: 'POST', body: JSON.stringify(body) }),
  reassignActivity: (id: string, assigneeId: string) => call(`/api/activities/${id}/reassign`, { method: 'POST', body: JSON.stringify({ assigneeId }) }),
  attention: (userId: string) => call(`/api/attention?userId=${encodeURIComponent(userId)}`),
  capacity: (userId: string) => call(`/api/capacity?userId=${encodeURIComponent(userId)}`),
  approvals: (userId: string) => call(`/api/approvals?userId=${encodeURIComponent(userId)}`),
  decideApproval: (id: string, body: any) => call(`/api/approvals/${id}/decide`, { method: 'POST', body: JSON.stringify(body) }),
  deliverables: (work: string) => call(`/api/deliverables?work=${encodeURIComponent(work)}`),
  attachDeliverable: (workId: string, body: any) => call(`/api/works/${workId}/deliverables/attach`, { method: 'POST', body: JSON.stringify(body) }),
  completeWork: (workId: string) => call(`/api/works/${workId}/complete`, { method: 'POST', body: '{}' }),
  resolve: (kind: 'user' | 'team' | 'initiative' | 'activity' | 'work', q: string) =>
    call(`/api/resolve?kind=${kind}&q=${encodeURIComponent(q)}`).then((r) => r.match as { id: string; name: string } | null),
};

// Convenience resolvers that throw a friendly error the model can relay.
export async function needTeam(q: string) { const m = await cadence.resolve('team', q); if (!m) throw new Error(`No team matches "${q}". Ask the user which team.`); return m; }
export async function needUser(q: string) { const m = await cadence.resolve('user', q); if (!m) throw new Error(`No person matches "${q}".`); return m; }
export async function needActivity(q: string) { const m = await cadence.resolve('activity', q); if (!m) throw new Error(`No activity matches "${q}".`); return m; }
export async function needWork(q: string) { const m = await cadence.resolve('work', q); if (!m) throw new Error(`No work matches "${q}". Ask the user which work package.`); return m; }
