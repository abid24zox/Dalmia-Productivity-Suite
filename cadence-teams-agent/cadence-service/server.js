// Cadence shared service — the REST contract that BOTH the Teams agent and the
// portal read from and write to. Single source of truth. Start: `npm start`
// (PORT 4000).
require('dotenv/config');
const express = require('express');
const cors = require('cors');
const store = require('./store');
const ai = require('./ai');

const app = express();
app.use(cors());
app.use(express.json({ limit: '30mb' })); // room for base64 docs / audio

// light request log
app.use((req, _res, next) => { console.log(`${req.method} ${req.url}`); next(); });

// Optional shared-secret guard (set CADENCE_API_KEY to enable). Login, health,
// and AI stay open so the browser portal can reach them without embedding a key.
app.use((req, res, next) => {
  const key = process.env.CADENCE_API_KEY;
  if (!key) return next();
  if (['/api/health', '/api/login'].includes(req.path) || req.path.startsWith('/api/ai')) return next();
  if (req.get('x-api-key') === key) return next();
  return res.status(401).json({ error: 'unauthorized' });
});

const ok = (res, data) => res.json(data);
const bad = (res, msg, code = 400) => res.status(code).json({ error: msg });
// Every write returns the fresh full snapshot so the portal reconciles to truth.
const withSnap = (res, extra = {}) => res.json({ ...extra, snapshot: store.snapshot() });

app.get('/api/health', (_req, res) => ok(res, { status: 'ok', foundry: ai.hasFoundry, deepgram: !!process.env.DEEPGRAM_API_KEY, time: new Date().toISOString() }));

/* ---------- auth ---------- */
app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  const u = store.login(username, password);
  if (!u) return bad(res, 'Wrong username or password.', 401);
  ok(res, { user: u });
});

/* ---------- raw snapshot (portal) ---------- */
app.get('/api/snapshot', (_req, res) => ok(res, store.snapshot()));

/* ---------- directory ---------- */
app.get('/api/users', (_req, res) => ok(res, { users: store.snapshot().users }));
app.get('/api/teams', (_req, res) => ok(res, { teams: store.db.teams.map((t) => ({ ...t, members: t.memberIds.map(store.uName) })) }));
app.post('/api/teams', (req, res) => {
  const { name, memberIds, id } = req.body || {};
  if (!name || !Array.isArray(memberIds) || !memberIds.length) return bad(res, 'name and memberIds[] required');
  const team = id ? store.addTeam({ id, name, memberIds }) : store.createTeam(name, memberIds);
  withSnap(res, { team });
});
app.patch('/api/teams/:id', (req, res) => { const t = store.patchTeam(req.params.id, req.body || {}); if (!t) return bad(res, 'team not found', 404); withSnap(res, { team: t }); });

/* ---------- scoped reads (bot) ---------- */
app.get('/api/portfolio', (req, res) => ok(res, store.getPortfolio(req.query.userId)));
app.get('/api/attention', (req, res) => ok(res, store.getAttention(req.query.userId)));
app.get('/api/capacity', (req, res) => ok(res, store.getCapacity(req.query.userId)));
app.get('/api/approvals', (req, res) => ok(res, store.getApprovals(req.query.userId)));

app.get('/api/resolve', (req, res) => {
  const { kind, q } = req.query;
  let e = null;
  if (kind === 'user') e = store.resolveUser(q);
  else if (kind === 'team') e = store.resolveTeam(q);
  else if (kind === 'initiative') e = store.resolveInitiative(q);
  else if (kind === 'activity') e = store.resolveActivity(q);
  ok(res, { match: e ? { id: e.id, name: e.name || e.title } : null });
});

app.get('/api/initiatives', (req, res) => ok(res, { initiatives: store.getPortfolio(req.query.userId).initiatives }));
app.get('/api/initiatives/:id', (req, res) => {
  const top = store.resolveInitiative(req.params.id);
  if (!top) return bad(res, 'initiative not found', 404);
  ok(res, store.detailInitiative(top));
});
app.post('/api/initiatives', (req, res) => {
  const { ownerId, title, type, objective, deadline, teamId, parentId, subworks } = req.body || {};
  if (!title) return bad(res, 'title required');
  const initiative = store.createInitiative({ ownerId, title, type, objective, deadline, teamId, parentId, subworks: subworks || [] });
  withSnap(res, { initiative });
});

/* ---------- generic work / activity writes (portal primitives) ---------- */
app.post('/api/works', (req, res) => { const works = store.addWorks(req.body?.works || req.body); withSnap(res, { works }); });
app.patch('/api/works/:id', (req, res) => { const w = store.patchWork(req.params.id, req.body || {}); if (!w) return bad(res, 'work not found', 404); withSnap(res, { work: w }); });
app.delete('/api/works/:id', (req, res) => { withSnap(res, store.deleteWork(req.params.id)); });

app.post('/api/activities', (req, res) => { const acts = store.addActs(req.body?.activities || req.body); withSnap(res, { activities: acts }); });
app.patch('/api/activities/:id', (req, res) => { const a = store.patchAct(req.params.id, req.body || {}); if (!a) return bad(res, 'activity not found', 404); withSnap(res, { activity: a }); });
app.delete('/api/activities/:id', (req, res) => { const okd = store.deleteAct(req.params.id); if (!okd) return bad(res, 'activity not found', 404); withSnap(res, { deleted: req.params.id }); });

// bot-friendly aliases
app.post('/api/activities/:id/schedule', (req, res) => { const a = store.scheduleActivity(req.params.id, req.body || {}); if (!a) return bad(res, 'activity not found', 404); withSnap(res, { activity: a }); });
app.post('/api/activities/:id/reassign', (req, res) => { const { assigneeId } = req.body || {}; if (!assigneeId) return bad(res, 'assigneeId required'); const a = store.reassignActivity(req.params.id, assigneeId); if (!a) return bad(res, 'activity not found', 404); withSnap(res, { activity: a }); });

/* ---------- change requests / approvals ---------- */
app.post('/api/crs', (req, res) => { const cr = store.addCr(req.body || {}); withSnap(res, { cr }); });
app.patch('/api/crs/:id', (req, res) => { const c = store.patchCr(req.params.id, req.body || {}); if (!c) return bad(res, 'change request not found', 404); withSnap(res, { cr: c }); });
app.post('/api/approvals/:id/decide', (req, res) => { const out = store.decideApproval(req.params.id, req.body || {}); if (!out) return bad(res, 'change request not found', 404); withSnap(res, out); });

/* ---------- remarks / nudges (portal) ---------- */
app.post('/api/remarks', (req, res) => { const r = store.addRemark(req.body || {}); withSnap(res, { remark: r }); });
app.post('/api/remarks/read', (req, res) => { const { userId } = req.body || {}; store.markRemarksRead(userId); withSnap(res, {}); });

/* ---------- AI (shared Foundry model) ---------- */
const aiErr = (res, e) => { if (e && e.notConfigured) return res.status(501).json({ error: e.message, notConfigured: true }); return res.status(500).json({ error: e.message || String(e) }); };

app.post('/api/ai/complete', async (req, res) => {
  try { const text = await ai.complete(req.body?.system, req.body?.user); ok(res, { text }); }
  catch (e) { aiErr(res, e); }
});
app.post('/api/ai/extract', async (req, res) => {
  try { const text = await ai.extract(req.body?.data, req.body?.name); ok(res, { text, name: req.body?.name || 'document' }); }
  catch (e) { aiErr(res, e); }
});
app.post('/api/ai/transcribe', async (req, res) => {
  try { const text = await ai.transcribe(req.body?.audio, req.body?.mimetype); ok(res, { text }); }
  catch (e) { aiErr(res, e); }
});

const PORT = process.env.PORT || 4000;
if (require.main === module) {
  app.listen(PORT, () => console.log(`Cadence service on http://localhost:${PORT}  (Foundry: ${ai.hasFoundry ? 'on' : 'off'})`));
}
module.exports = app;
