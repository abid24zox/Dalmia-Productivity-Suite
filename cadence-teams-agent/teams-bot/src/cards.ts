// Adaptive Card builders — the agent replies with these instead of walls of
// text. Cards are plain JSON (schema 1.4) wrapped by CardFactory in bot.ts.
const PORTAL = process.env.PORTAL_URL || 'https://cadence.contoso.com';
const ragColor = (rag: string) => (rag === 'red' ? 'Attention' : rag === 'amber' ? 'Warning' : 'Good');

const card = (body: any[], actions: any[] = []) => ({
  type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', body, ...(actions.length ? { actions } : {}),
});
const title = (t: string, sub?: string) => ({ type: 'Container', items: [{ type: 'TextBlock', text: t, weight: 'Bolder', size: 'Medium', wrap: true }, ...(sub ? [{ type: 'TextBlock', text: sub, isSubtle: true, spacing: 'None', wrap: true }] : [])] });

export function portfolioCard(p: any) {
  const t = p.tiles;
  const facts = [
    { title: 'Scope', value: t.scope }, { title: 'Initiatives', value: `${t.initiatives}` },
    { title: 'On track / at risk', value: `${t.onTrack} / ${t.atRisk}` },
    { title: 'Avg result', value: t.avgResult == null ? '—' : `${t.avgResult}%` },
    { title: 'Overdue activities', value: `${t.overdueActivities}` }, { title: 'Approvals pending', value: `${t.approvalsPending}` },
  ];
  const rows = p.initiatives.map((i: any) => ({
    type: 'Container', spacing: 'Small', separator: true, items: [
      { type: 'ColumnSet', columns: [
        { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: i.title, weight: 'Bolder', wrap: true }, { type: 'TextBlock', text: `${i.ownerName} · ${i.fn} · ${i.sufficiency}`, isSubtle: true, spacing: 'None', wrap: true }] },
        { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: i.resultPct == null ? '—' : `${i.resultPct}%`, color: ragColor(i.rag), weight: 'Bolder' }, { type: 'TextBlock', text: 'result', isSubtle: true, spacing: 'None', size: 'Small' }] },
      ] },
      { type: 'TextBlock', text: `Planning ${i.planning}% · Execution ${i.execution}% · ${i.stats.done}/${i.stats.total} done${i.stats.overdue ? ` · ${i.stats.overdue} overdue` : ''}${i.stuck ? ` · stuck: ${i.stuck}` : ''}`, size: 'Small', isSubtle: true, wrap: true },
    ],
  }));
  return card([title('Cadence — portfolio', t.scope), { type: 'FactSet', facts }, ...rows], [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

export function initiativeCard(i: any) {
  const gapTxt = i.result ? `${i.result.metric}: ${i.result.current} of ${i.result.target} ${i.result.unit} · gap ${i.gap} ${i.result.unit}` : 'No result metric';
  const delivTxt = (s: any) => {
    const d = s.deliverables; if (!d || !d.total) return '';
    return ` · deliverables ${d.done}/${d.total}${d.avgScore != null ? ` · avg ${d.avgScore}` : ''}`;
  };
  const subs = (i.subworks || []).map((s: any) => ({
    type: 'Container', spacing: 'Small', separator: true, items: [
      { type: 'TextBlock', text: `${s.title}  —  P ${s.planning}% · E ${s.execution}%${s.completedAt ? ' · ✔ complete' : ''}${delivTxt(s)}`, weight: 'Bolder', wrap: true },
      ...s.activities.map((a: any) => ({ type: 'TextBlock', size: 'Small', wrap: true, color: a.overdue ? 'Attention' : a.blocked ? 'Warning' : 'Default', text: `• ${a.title} — ${a.assigneeName || 'Unassigned'} · ${a.startDate ? a.startDate + ' → ' : ''}${a.date || 'no date'} · ${a.actualHrs || 0}/${a.plannedHrs}h · ${a.status}${a.blocked ? ' · blocked' : ''}` })),
    ],
  }));
  return card([
    title(i.title, `${i.ownerName} · ${i.fn}${i.teamName ? ' · ' + i.teamName : ''} · ${i.scope || ''}`),
    { type: 'FactSet', facts: [
      { title: 'Result', value: i.resultPct == null ? '—' : `${i.resultPct}%` }, { title: 'Gap to target', value: i.gap == null ? '—' : `${i.gap} ${i.result?.unit || ''}` },
      { title: 'Planning / Execution', value: `${i.planning}% / ${i.execution}%` }, { title: 'Status', value: `${i.rag.toUpperCase()} · ${i.sufficiency}` },
      { title: 'Timeline', value: `${i.startDate || '?'} → ${i.endDate || i.deadline || '?'}` }, { title: 'Effort (actual/est)', value: i.effort ? `${i.effort.actual}/${i.effort.planned} h` : '—' },
    ] },
    { type: 'TextBlock', text: gapTxt, isSubtle: true, wrap: true, spacing: 'Small' },
    ...subs,
  ], [{ type: 'Action.OpenUrl', title: 'Open in portal', url: `${PORTAL}` }]);
}

export function createdCard(i: any) {
  const nAct = (i.subworks || []).reduce((s: number, sw: any) => s + sw.activities.length, 0);
  return card([
    title('Initiative created & assigned', i.title),
    { type: 'FactSet', facts: [
      { title: 'Type', value: i.type }, { title: 'Team', value: i.teamName || '—' }, { title: 'Scope', value: i.scope || '—' },
      { title: 'Deadline', value: i.deadline || '—' }, { title: 'Works', value: `${(i.works || i.subworks || []).length}` }, { title: 'Activities assigned', value: `${nAct}` },
    ] },
    ...(i.subworks || []).map((s: any) => ({ type: 'TextBlock', size: 'Small', wrap: true, text: `**${s.title}** — ${s.activities.map((a: any) => `${a.title} (${a.assigneeName || 'unassigned'}, ${a.date || 'tbd'})`).join('; ')}` })),
  ], [{ type: 'Action.OpenUrl', title: 'View in portal', url: PORTAL }]);
}

export function attentionCard(items: any[]) {
  if (!items.length) return card([title('Needs attention'), { type: 'TextBlock', text: 'All clear — nothing overdue, blocked, or stuck in your scope.', wrap: true }]);
  return card([title('Needs attention', `${items.length} item(s)`), ...items.map((it) => ({ type: 'TextBlock', wrap: true, color: it.kind === 'overdue' || it.kind === 'blocked' ? 'Attention' : 'Warning', text: `• [${it.kind}] ${it.title} — ${it.detail}${it.assignee ? ` (${it.assignee})` : ''}` }))]);
}

export function capacityCard(rows: any[]) {
  return card([title('Team capacity', 'Open (not-yet-done) work per person'), { type: 'FactSet', facts: rows.map((r) => ({ title: `${r.name} · ${r.fn}`, value: `${r.openHours}h` })) }]);
}

export function memberStatusCard(r: any) {
  const u = r.user, a = r.activities, d = r.deliverables;
  const facts = [
    { title: 'Role', value: `${u.title} · ${u.fn}` },
    { title: 'Execution', value: `${a.execPct}% · ${a.done}/${a.total} activities done` },
    { title: 'Needs attention', value: `${a.overdue} overdue · ${a.blocked} blocked` },
    { title: 'Deliverables', value: d.total ? `${d.delivered}/${d.total}${d.avgScore != null ? ` · avg ${d.avgScore}/100` : ''}` : '—' },
  ];
  const owned = (r.owned || []).map((w: any) => ({
    type: 'TextBlock', size: 'Small', wrap: true, color: ragColor(w.rag),
    text: `• ${w.title} — ${w.level}${w.resultPct != null ? ` · ${w.resultPct}% result` : ''} · P ${w.planning}% / E ${w.execution}%`,
  }));
  return card([
    title(u.name, `${u.title} · ${u.fn}`),
    { type: 'FactSet', facts },
    ...(owned.length ? [{ type: 'TextBlock', text: 'Owns', weight: 'Bolder', spacing: 'Medium', wrap: true }, ...owned]
                     : [{ type: 'TextBlock', text: 'Owns no initiatives or works directly.', isSubtle: true, wrap: true, spacing: 'Small' }]),
  ], [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

export function approvalsCard(pending: any[]) {
  if (!pending.length) return card([title('Approvals'), { type: 'TextBlock', text: 'No pending approvals in your scope.', wrap: true }]);
  const body: any[] = [title('Pending approvals', `${pending.length}`)];
  for (const c of pending) {
    body.push({ type: 'Container', separator: true, spacing: 'Small', items: [
      { type: 'TextBlock', text: `${c.proposer} — ${c.kindLabel}`, weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: `${c.initiative}`, isSubtle: true, size: 'Small', spacing: 'None', wrap: true },
      { type: 'TextBlock', text: c.desc, wrap: true, size: 'Small' },
      { type: 'Input.Text', id: `remark_${c.id}`, placeholder: 'Optional remark (can spin off a follow-up)…', isMultiline: true },
      { type: 'Input.Toggle', id: `spinoff_${c.id}`, title: 'Create a follow-up initiative from my remark', value: 'false' },
      { type: 'ActionSet', actions: [
        { type: 'Action.Submit', title: 'Approve', style: 'positive', data: { action: 'decide_approval', id: c.id, approve: true } },
        { type: 'Action.Submit', title: 'Reject', data: { action: 'decide_approval', id: c.id, approve: false } },
      ] },
    ] });
  }
  return card(body);
}

const KIND_ICON: Record<string, string> = { document: '📄', spreadsheet: '📊', email: '✉️', slides: '📽️', other: '📎' };

export function deliverablesCard(d: any) {
  if (!d) return textCard('Deliverables', 'No matching work.');
  const head = title(`${d.workTitle} — deliverables`, `${d.done}/${d.total} delivered${d.avgScore != null ? ` · avg ${d.avgScore}/100` : ''}${d.completedAt ? ' · ✔ complete' : ''}`);
  if (!d.items.length) return card([head, { type: 'TextBlock', text: 'No deliverables defined on this work yet.', isSubtle: true, wrap: true }]);
  const rows = d.items.map((it: any) => ({
    type: 'ColumnSet', spacing: 'Small', separator: true, columns: [
      { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: it.done ? '✅' : '⬜', wrap: false }] },
      { type: 'Column', width: 'stretch', items: [
        { type: 'TextBlock', text: `${KIND_ICON[it.kind] || '📎'} ${it.label}`, wrap: true },
        ...(it.file || it.verdict ? [{ type: 'TextBlock', text: `${it.file ? it.file : ''}${it.file && it.verdict ? ' · ' : ''}${it.verdict || ''}`, isSubtle: true, size: 'Small', spacing: 'None', wrap: true }] : []),
      ] },
      { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: it.score != null ? `${it.score}/100` : '—', color: it.score != null ? 'Good' : 'Default', horizontalAlignment: 'Right' }] },
    ],
  }));
  return card([head, ...rows], [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

export function deliverableLoggedCard(work: any, item: any, scored: boolean) {
  const facts = [
    { title: 'Work', value: work.title }, { title: 'Deliverable', value: item.label },
    { title: 'File', value: item.file ? item.file.name : '—' }, { title: 'Status', value: 'Delivered ✅' },
  ];
  if (scored && typeof item.score === 'number') facts.push({ title: 'Score', value: `${item.score}/100 — ${item.verdict || ''}` });
  return card([
    title('Deliverable logged', `${KIND_ICON[item.kind] || '📎'} ${item.label}`),
    { type: 'FactSet', facts },
    ...(scored && item.feedback ? [{ type: 'TextBlock', text: item.feedback, wrap: true, isSubtle: true, spacing: 'Small' }] : []),
  ], [{ type: 'Action.OpenUrl', title: 'Open in portal', url: PORTAL }]);
}

export function textCard(heading: string, msg: string) {
  return card([title(heading), { type: 'TextBlock', text: msg, wrap: true }]);
}
