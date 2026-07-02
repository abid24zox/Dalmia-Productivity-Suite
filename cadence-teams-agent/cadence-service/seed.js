// Seed data — the SINGLE source of truth shared by the Cadence portal (reads +
// writes) and the Teams agent (reads + writes). This mirrors the portal
// prototype's object model EXACTLY: a 5-level tree
//   objective -> initiative -> work -> subwork -> activity
// with an explicit `level` on every work node. Dates are relative to "today" so
// the demo always has overdue / upcoming items.
const MSD = 86400000;
const sod = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const TODAY = sod(new Date());
const addDays = (d, n) => sod(new Date(sod(d).getTime() + n * MSD));
const iso = (d) => sod(d).toISOString().slice(0, 10);
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
  // Objectives (level 1) — MD-owned
  { id: 'o1', parentId: null, level: 'objective', title: 'Modernise field operations', type: 'general', ownerId: 'u_vik' },
  { id: 'o2', parentId: null, level: 'objective', title: 'Cut supply-chain cost', type: 'general', ownerId: 'u_vik' },
  // Initiatives (level 2) — VP-owned, carry the result metric
  { id: 'w1', parentId: 'o1', level: 'initiative', title: 'Replace retired laptops (Oct-24 & older)', type: 'procurement', ownerId: 'u_pri', teamId: 't_it', scope: 'group', result: { metric: 'Laptops issued', unit: 'count', baseline: 0, target: 120, current: 38 } },
  { id: 'w1x', parentId: 'o1', level: 'initiative', title: 'Roll out MFA to field staff', type: 'compliance', ownerId: 'u_pri', teamId: 't_it', scope: 'group', result: { metric: 'Coverage', unit: '%', baseline: 0, target: 100, current: 20 } },
  { id: 'w2', parentId: 'o2', level: 'initiative', title: 'Reduce logistics cost per bag by ₹0.50', type: 'cost', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', result: { metric: 'Cost saved', unit: '₹/bag', baseline: 0, target: 0.5, current: 0.18 } },
  // Works (level 3)
  { id: 'w1W1', parentId: 'w1', level: 'work', title: 'Demand & selection', type: 'procurement', ownerId: 'u_roh' },
  { id: 'w1W2', parentId: 'w1', level: 'work', title: 'Procurement', type: 'procurement', ownerId: 'u_pri' },
  { id: 'w1W3', parentId: 'w1', level: 'work', title: 'Rollout', type: 'procurement', ownerId: 'u_roh' },
  { id: 'w1xW1', parentId: 'w1x', level: 'work', title: 'Assess & pilot', type: 'compliance', ownerId: 'u_roh' },
  { id: 'w2W1', parentId: 'w2', level: 'work', title: 'Freight optimisation', type: 'cost', ownerId: 'u_neh' },
  { id: 'w2W2', parentId: 'w2', level: 'work', title: 'Commercials', type: 'cost', ownerId: 'u_mee' },
  // Sub-works (level 4) — hold the activities
  { id: 'w1a', parentId: 'w1W1', level: 'subwork', title: 'Employee survey', type: 'procurement', ownerId: 'u_roh' },
  { id: 'w1b', parentId: 'w1W1', level: 'subwork', title: 'Shortlist & finalize models', type: 'procurement', ownerId: 'u_roh' },
  { id: 'w1c', parentId: 'w1W2', level: 'subwork', title: 'Budget approvals from HODs', type: 'procurement', ownerId: 'u_pri' },
  { id: 'w1d', parentId: 'w1W2', level: 'subwork', title: 'Raise PR–PO & procure', type: 'procurement', ownerId: 'u_pri' },
  { id: 'w1e', parentId: 'w1W3', level: 'subwork', title: 'Image, issue & hand over', type: 'procurement', ownerId: 'u_roh' },
  { id: 'w1xa', parentId: 'w1xW1', level: 'subwork', title: 'Assess current auth', type: 'compliance', ownerId: 'u_roh' },
  { id: 'w2a', parentId: 'w2W1', level: 'subwork', title: 'Optimize primary freight routes', type: 'cost', ownerId: 'u_neh' },
  { id: 'w2b', parentId: 'w2W1', level: 'subwork', title: 'Improve truck load factor', type: 'cost', ownerId: 'u_neh' },
  { id: 'w2c', parentId: 'w2W2', level: 'subwork', title: 'Renegotiate transporter contracts', type: 'cost', ownerId: 'u_mee' },
  // Operations objective — recurring plant routines (cement)
  { id: 'o3', parentId: null, level: 'objective', title: 'Run reliable, compliant plant operations', type: 'general', ownerId: 'u_vik' },
  { id: 'w3', parentId: 'o3', level: 'initiative', title: 'Plant routines & statutory upkeep', type: 'compliance', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', result: { metric: 'Plant uptime', unit: '%', baseline: 0, target: 95, current: 88 } },
  { id: 'w3W1', parentId: 'w3', level: 'work', title: 'Monthly payroll run', type: 'general', ownerId: 'u_ana', recurring: { cadence: 'monthly' } },
  { id: 'w3W2', parentId: 'w3', level: 'work', title: 'Kiln & raw-mill preventive maintenance', type: 'general', ownerId: 'u_sam', recurring: { cadence: 'monthly' } },
  { id: 'w3W3', parentId: 'w3', level: 'work', title: 'Statutory & pollution-control compliance', type: 'compliance', ownerId: 'u_div', recurring: { cadence: 'quarterly' } },
  { id: 'w3a', parentId: 'w3W1', level: 'subwork', title: 'July payroll cycle', type: 'general', ownerId: 'u_ana' },
  { id: 'w3b', parentId: 'w3W2', level: 'subwork', title: 'July PM cycle — Kiln line 2', type: 'general', ownerId: 'u_sam' },
  { id: 'w3c', parentId: 'w3W3', level: 'subwork', title: 'Q2 emissions returns & consent renewal', type: 'compliance', ownerId: 'u_div' },
];

const ACTIVITIES = [
  { id: 'a1', workId: 'w1a', title: 'Draft survey form', assigneeId: 'u_roh', date: D(-2), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'a2', workId: 'w1a', title: 'Send survey to all staff', assigneeId: 'u_roh', date: D(-1), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'self' },
  { id: 'a3', workId: 'w1a', title: 'Collate responses', assigneeId: 'u_roh', date: D(-1), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a3b', workId: 'w1a', title: 'Prepare survey summary', assigneeId: 'u_roh', date: D(0), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self' },
  { id: 'a4', workId: 'w1b', title: 'Gather target specs', assigneeId: 'u_roh', date: D(-1), status: 'executed', plannedHrs: 2, actualHrs: 3, actType: 'self' },
  { id: 'a5', workId: 'w1b', title: 'Get 3 vendor quotes', assigneeId: 'u_roh', date: D(2), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'call' },
  { id: 'a6', workId: 'w1c', title: 'Prepare BOM + cost sheet', assigneeId: null, date: null, status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a7', workId: 'w1c', title: 'HOD sign-offs (IT, Fin, Ops)', assigneeId: null, date: null, status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'meeting' },
  { id: 'a8', workId: 'w1d', title: 'Raise purchase requisition', assigneeId: 'u_roh', date: D(1), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self' },
  { id: 'a9', workId: 'w1d', title: 'Issue PO to vendor', assigneeId: 'u_roh', date: D(3), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self', blocked: true },
  { id: 'ae1', workId: 'w1e', title: 'Image devices', assigneeId: 'u_roh', date: D(4), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'ae2', workId: 'w1e', title: 'Hand over to users', assigneeId: 'u_roh', date: D(6), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'site' },
  { id: 'ax1', workId: 'w1xa', title: 'Inventory apps & auth methods', assigneeId: 'u_roh', date: D(0), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'ax2', workId: 'w1xa', title: 'Map risky logins', assigneeId: 'u_roh', date: D(2), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a10', workId: 'w2a', title: 'Map current lane costs', assigneeId: 'u_neh', date: D(-3), status: 'executed', plannedHrs: 3, actualHrs: 3, actType: 'self' },
  { id: 'a11', workId: 'w2a', title: 'Model route consolidation', assigneeId: 'u_neh', date: D(0), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'self' },
  { id: 'a12', workId: 'w2b', title: 'Pull load-factor data by plant', assigneeId: 'u_neh', date: D(-1), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'ap1', workId: 'w3a', title: 'Validate attendance & inputs', assigneeId: 'u_ana', date: D(-2), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'ap2', workId: 'w3a', title: 'Process & disburse salaries', assigneeId: 'u_ana', date: D(1), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'am1', workId: 'w3b', title: 'Inspect refractory & mill bearings', assigneeId: 'u_sam', date: D(0), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'site' },
  { id: 'am2', workId: 'w3b', title: 'Lubrication & alignment', assigneeId: 'u_sam', date: D(3), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'site' },
  { id: 'ac1', workId: 'w3c', title: 'Compile CPCB emission data', assigneeId: 'u_div', date: D(-2), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'ac2', workId: 'w3c', title: 'File returns & consent renewal', assigneeId: 'u_div', date: D(4), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
];

const CRS = [
  { id: 'cr1', workId: 'w1', subworkId: 'w1a', proposerId: 'u_roh', kind: 'add_activity', desc: 'Found 3 depots missed in the first survey — need a quick re-run.', payload: { title: 'Re-run survey for 3 missed depots', hrs: 2, type: 'self' }, status: 'pending' },
];

// Portal-only cross-user nudges. The bot ignores these, but they are persisted
// so multiple portal sessions stay consistent.
const REMARKS = [];

module.exports = { USERS, TEAMS, WORKS, ACTIVITIES, CRS, REMARKS, MSD, sod, addDays, iso, TODAY };
