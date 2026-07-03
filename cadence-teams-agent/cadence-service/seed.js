// Seed data — the SINGLE source of truth shared by the Cadence portal (reads +
// writes) and the Teams agent (reads + writes). This mirrors the portal
// prototype's object model EXACTLY: a 4-level tree
//   objective -> initiative -> work -> activity
// with an explicit `level` on every work node; activities attach directly to a
// work via `workId`. Dates are relative to "today" so the demo always has
// overdue / upcoming items.
const MSD = 86400000;
const sod = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const TODAY = sod(new Date());
const addDays = (d, n) => sod(new Date(sod(d).getTime() + n * MSD));
// Local-date ISO — not toISOString(), which shifts to the previous day in +offset zones.
const iso = (d) => { const x = sod(d); return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`; };
const D = (n) => iso(addDays(TODAY, n));

// Portal login credentials live here now (server-side). Members without a
// username/password can still be assigned work; they just don't sign in.
const USERS = [
  { id: 'u_vik', username: 'vikram', password: 'md@2026', email: 'vikram@contoso.com', name: 'Vikram Rao', title: 'MD', level: 'md', fn: 'Office of MD', reports_to: null },
  { id: 'u_mee', username: 'meera', password: 'vp@2026', email: 'meera@contoso.com', name: 'Meera Iyer', title: 'VP, Supply Chain', level: 'vp', fn: 'Supply Chain', reports_to: 'u_vik' },
  { id: 'u_pri', username: 'priya', password: 'vp@2026', email: 'priya@contoso.com', name: 'Priya Nair', title: 'VP, IT', level: 'vp', fn: 'IT', reports_to: 'u_vik' },
  { id: 'u_roh', username: 'rohit', password: 'team@2026', email: 'rohit@contoso.com', name: 'Rohit Sharma', title: 'Executive, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_neh', username: 'neha', password: 'team@2026', email: 'neha@contoso.com', name: 'Neha Gupta', title: 'Executive, SC', level: 'member', fn: 'Supply Chain', reports_to: 'u_mee' },
  { id: 'u_arj', email: 'arjun@contoso.com', name: 'Arjun Mehta', title: 'Executive, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_kav', email: 'kavya@contoso.com', name: 'Kavya Reddy', title: 'Analyst, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_sam', email: 'sameer@contoso.com', name: 'Sameer Khan', title: 'Executive, SC', level: 'member', fn: 'Supply Chain', reports_to: 'u_mee' },
  { id: 'u_div', email: 'divya@contoso.com', name: 'Divya Rao', title: 'Analyst, SC', level: 'member', fn: 'Supply Chain', reports_to: 'u_mee' },
  { id: 'u_ana', email: 'anita@contoso.com', name: 'Anita Desai', title: 'Manager, Finance', level: 'member', fn: 'Finance', reports_to: 'u_vik' },
  { id: 'u_raj', email: 'raj@contoso.com', name: 'Raj Malhotra', title: 'Manager, HR', level: 'member', fn: 'HR', reports_to: 'u_vik' },
];

const TEAMS = [
  { id: 't_it', name: 'IT Ops', memberIds: ['u_pri', 'u_roh', 'u_arj', 'u_kav'] },
  { id: 't_sc', name: 'Supply Chain', memberIds: ['u_mee', 'u_neh', 'u_sam', 'u_div'] },
  { id: 't_fin', name: 'Finance & Controls', memberIds: ['u_ana', 'u_raj'] },
  { id: 't_dig', name: 'Digital Rollout (cross-fn)', memberIds: ['u_roh', 'u_sam', 'u_ana'] },
];

const WORKS = [
  // Objectives (level 1) — MD-owned. Each carries a deadline so every level shows
  // a due date and "days left" up the tree.
  { id: 'o1', parentId: null, level: 'objective', title: 'Modernise field operations', type: 'general', ownerId: 'u_vik', deadline: D(40) },
  { id: 'o2', parentId: null, level: 'objective', title: 'Cut supply-chain cost', type: 'general', ownerId: 'u_vik', deadline: D(55) },
  // Initiatives (level 2) — VP-owned, carry the result metric + a deadline
  { id: 'w1', parentId: 'o1', level: 'initiative', title: 'Replace retired laptops (Oct-24 & older)', type: 'procurement', ownerId: 'u_pri', teamId: 't_it', scope: 'group', deadline: D(18), result: { metric: 'Laptops issued', unit: 'count', baseline: 0, target: 120, current: 38 } },
  { id: 'w1x', parentId: 'o1', level: 'initiative', title: 'Roll out MFA to field staff', type: 'compliance', ownerId: 'u_pri', teamId: 't_it', scope: 'group', deadline: D(30), result: { metric: 'Coverage', unit: '%', baseline: 0, target: 100, current: 20 } },
  { id: 'w2', parentId: 'o2', level: 'initiative', title: 'Reduce logistics cost per bag by ₹0.50', type: 'cost', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', deadline: D(45), result: { metric: 'Cost saved', unit: '₹/bag', baseline: 0, target: 0.5, current: 0.18 } },
  // Works (level 3) — hold the activities directly; each has its own deadline
  { id: 'w1W1', parentId: 'w1', level: 'work', title: 'Demand & selection', type: 'procurement', ownerId: 'u_roh', deadline: D(3) },
  { id: 'w1W2', parentId: 'w1', level: 'work', title: 'Procurement', type: 'procurement', ownerId: 'u_pri', deadline: D(10) },
  { id: 'w1W3', parentId: 'w1', level: 'work', title: 'Rollout', type: 'procurement', ownerId: 'u_roh', deadline: D(16) },
  { id: 'w1xW1', parentId: 'w1x', level: 'work', title: 'Assess & pilot', type: 'compliance', ownerId: 'u_roh', deadline: D(12) },
  { id: 'w2W1', parentId: 'w2', level: 'work', title: 'Freight optimisation', type: 'cost', ownerId: 'u_neh', deadline: D(8) },
  { id: 'w2W2', parentId: 'w2', level: 'work', title: 'Commercials', type: 'cost', ownerId: 'u_mee', deadline: D(40) },
  // Operations objective — recurring plant routines (cement)
  { id: 'o3', parentId: null, level: 'objective', title: 'Run reliable, compliant plant operations', type: 'general', ownerId: 'u_vik', deadline: D(80) },
  { id: 'w3', parentId: 'o3', level: 'initiative', title: 'Plant routines & statutory upkeep', type: 'compliance', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', deadline: D(25), result: { metric: 'Plant uptime', unit: '%', baseline: 0, target: 95, current: 88 } },
  { id: 'w3W1', parentId: 'w3', level: 'work', title: 'Monthly payroll run', type: 'general', ownerId: 'u_ana', deadline: D(3), recurring: { cadence: 'monthly' } },
  { id: 'w3W2', parentId: 'w3', level: 'work', title: 'Kiln & raw-mill preventive maintenance', type: 'general', ownerId: 'u_sam', deadline: D(8), recurring: { cadence: 'monthly' } },
  { id: 'w3W3', parentId: 'w3', level: 'work', title: 'Statutory & pollution-control compliance', type: 'compliance', ownerId: 'u_div', deadline: D(-1), recurring: { cadence: 'quarterly' } },
];

// Activities attach directly to a work (level 3) via workId.
const ACTIVITIES = [
  { id: 'a1', workId: 'w1W1', title: 'Draft survey form', assigneeId: 'u_roh', date: D(-2), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'a2', workId: 'w1W1', title: 'Send survey to all staff', assigneeId: 'u_roh', date: D(-1), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'self' },
  { id: 'a3', workId: 'w1W1', title: 'Collate responses', assigneeId: 'u_roh', date: D(-1), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a3b', workId: 'w1W1', title: 'Prepare survey summary', assigneeId: 'u_roh', date: D(0), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self' },
  { id: 'a4', workId: 'w1W1', title: 'Gather target specs', assigneeId: 'u_roh', date: D(-1), status: 'executed', plannedHrs: 2, actualHrs: 3, actType: 'self' },
  { id: 'a5', workId: 'w1W1', title: 'Get 3 vendor quotes', assigneeId: 'u_roh', date: D(2), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'call' },
  { id: 'a6', workId: 'w1W2', title: 'Prepare BOM + cost sheet', assigneeId: null, date: null, status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a7', workId: 'w1W2', title: 'HOD sign-offs (IT, Fin, Ops)', assigneeId: null, date: null, status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'meeting' },
  { id: 'a8', workId: 'w1W2', title: 'Raise purchase requisition', assigneeId: 'u_roh', date: D(1), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self' },
  { id: 'a9', workId: 'w1W2', title: 'Issue PO to vendor', assigneeId: 'u_roh', date: D(3), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self', blocked: true },
  { id: 'ae1', workId: 'w1W3', title: 'Image devices', assigneeId: 'u_roh', date: D(4), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'ae2', workId: 'w1W3', title: 'Hand over to users', assigneeId: 'u_roh', date: D(6), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'site' },
  { id: 'ax1', workId: 'w1xW1', title: 'Inventory apps & auth methods', assigneeId: 'u_roh', date: D(0), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'ax2', workId: 'w1xW1', title: 'Map risky logins', assigneeId: 'u_roh', date: D(2), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a10', workId: 'w2W1', title: 'Map current lane costs', assigneeId: 'u_neh', date: D(-3), status: 'executed', plannedHrs: 3, actualHrs: 3, actType: 'self' },
  { id: 'a11', workId: 'w2W1', title: 'Model route consolidation', assigneeId: 'u_neh', date: D(0), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'self' },
  { id: 'a12', workId: 'w2W1', title: 'Pull load-factor data by plant', assigneeId: 'u_neh', date: D(-1), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'a13', workId: 'w2W2', title: 'Draft renegotiation terms', assigneeId: 'u_mee', date: D(1), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'ap1', workId: 'w3W1', title: 'Validate attendance & inputs', assigneeId: 'u_ana', date: D(-2), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'ap2', workId: 'w3W1', title: 'Process & disburse salaries', assigneeId: 'u_ana', date: D(1), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'am1', workId: 'w3W2', title: 'Inspect refractory & mill bearings', assigneeId: 'u_sam', date: D(0), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'site' },
  { id: 'am2', workId: 'w3W2', title: 'Lubrication & alignment', assigneeId: 'u_sam', date: D(3), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'site' },
  { id: 'ac1', workId: 'w3W3', title: 'Compile CPCB emission data', assigneeId: 'u_div', date: D(-2), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'ac2', workId: 'w3W3', title: 'File returns & consent renewal', assigneeId: 'u_div', date: D(4), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
];

const CRS = [
  { id: 'cr1', workId: 'w1', subworkId: 'w1W1', proposerId: 'u_roh', kind: 'add_activity', desc: 'Found 3 depots missed in the first survey — need a quick re-run.', payload: { title: 'Re-run survey for 3 missed depots', hrs: 2, type: 'self' }, status: 'pending' },
];

// Portal-only cross-user nudges. The bot ignores these, but they are persisted
// so multiple portal sessions stay consistent.
const REMARKS = [];

module.exports = { USERS, TEAMS, WORKS, ACTIVITIES, CRS, REMARKS, MSD, sod, addDays, iso, TODAY };
