// Seed data — the SINGLE source of truth shared by the Cadence portal (reads +
// writes) and the Teams agent (reads + writes). 4-level tree
//   objective -> initiative -> work -> activity
// with an explicit `level` on every work node. `deadline` (on works/initiatives)
// and each activity's `date` + `plannedHrs` drive the time-based tracking.
// Dates are relative to "today" and statuses are deliberately spread across
// done / on-track / at-risk / overdue / blocked so the portfolio shows a real
// range of health rather than everything red.
const MSD = 86400000;
const sod = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const TODAY = sod(new Date());
const addDays = (d, n) => sod(new Date(sod(d).getTime() + n * MSD));
// Local (not UTC) YYYY-MM-DD so seeded dates match the browser's local calendar.
const iso = (d) => { const x = sod(d); return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`; };
const D = (n) => iso(addDays(TODAY, n));

// Portal login credentials live here now (server-side). Members without a
// username/password can still be assigned work; they just don't sign in.
const USERS = [
  { id: 'u_vik', username: 'vikram', password: 'md@2026', email: 'vikram@contoso.com', name: 'Vikram Rao', title: 'MD', level: 'md', fn: 'Office of MD', reports_to: null },
  { id: 'u_mee', username: 'meera', password: 'vp@2026', email: 'meera@contoso.com', name: 'Meera Iyer', title: 'VP, Supply Chain', level: 'vp', fn: 'Supply Chain', reports_to: 'u_vik' },
  { id: 'u_pri', username: 'priya', password: 'vp@2026', email: 'priya@contoso.com', name: 'Priya Nair', title: 'VP, IT', level: 'vp', fn: 'IT', reports_to: 'u_vik' },
  { id: 'u_kar', username: 'karan', password: 'vp@2026', email: 'karan@contoso.com', name: 'Karan Malhotra', title: 'VP, Commercial', level: 'vp', fn: 'Commercial', reports_to: 'u_vik' },
  { id: 'u_roh', username: 'rohit', password: 'team@2026', email: 'rohit@contoso.com', name: 'Rohit Sharma', title: 'Executive, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_neh', username: 'neha', password: 'team@2026', email: 'neha@contoso.com', name: 'Neha Gupta', title: 'Executive, SC', level: 'member', fn: 'Supply Chain', reports_to: 'u_mee' },
  { id: 'u_tan', username: 'tanvi', password: 'team@2026', email: 'tanvi@contoso.com', name: 'Tanvi Shah', title: 'Executive, Commercial', level: 'member', fn: 'Commercial', reports_to: 'u_kar' },
  { id: 'u_arj', email: 'arjun@contoso.com', name: 'Arjun Mehta', title: 'Executive, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_kav', email: 'kavya@contoso.com', name: 'Kavya Reddy', title: 'Analyst, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_yash', email: 'yash@contoso.com', name: 'Yash Verma', title: 'Executive, IT', level: 'member', fn: 'IT', reports_to: 'u_pri' },
  { id: 'u_sam', email: 'sameer@contoso.com', name: 'Sameer Khan', title: 'Executive, SC', level: 'member', fn: 'Supply Chain', reports_to: 'u_mee' },
  { id: 'u_div', email: 'divya@contoso.com', name: 'Divya Rao', title: 'Analyst, SC', level: 'member', fn: 'Supply Chain', reports_to: 'u_mee' },
  { id: 'u_ana', email: 'anita@contoso.com', name: 'Anita Desai', title: 'Manager, Finance', level: 'member', fn: 'Finance', reports_to: 'u_vik' },
  { id: 'u_raj', email: 'raj@contoso.com', name: 'Raj Malhotra', title: 'Manager, HR', level: 'member', fn: 'HR', reports_to: 'u_vik' },
];

const TEAMS = [
  { id: 't_it', name: 'IT Ops', memberIds: ['u_pri', 'u_roh', 'u_arj', 'u_kav', 'u_yash'] },
  { id: 't_sc', name: 'Supply Chain', memberIds: ['u_mee', 'u_neh', 'u_sam', 'u_div'] },
  { id: 't_com', name: 'Commercial', memberIds: ['u_kar', 'u_tan', 'u_neh'] },
  { id: 't_fin', name: 'Finance & Controls', memberIds: ['u_ana', 'u_raj'] },
  { id: 't_dig', name: 'Digital Rollout (cross-fn)', memberIds: ['u_roh', 'u_sam', 'u_ana'] },
];

const WORKS = [
  // ===== Objectives (level 1) — MD-owned =====
  { id: 'o1', parentId: null, level: 'objective', title: 'Modernise field operations', type: 'general', ownerId: 'u_vik', deadline: D(50) },
  { id: 'o2', parentId: null, level: 'objective', title: 'Cut supply-chain cost', type: 'general', ownerId: 'u_vik', deadline: D(65) },
  { id: 'o3', parentId: null, level: 'objective', title: 'Run reliable, compliant plant operations', type: 'general', ownerId: 'u_vik', deadline: D(35) },
  { id: 'o4', parentId: null, level: 'objective', title: 'Strengthen dealer & customer experience', type: 'general', ownerId: 'u_vik', deadline: D(70) },
  { id: 'o5', parentId: null, level: 'objective', title: 'Build a high-performing team', type: 'general', ownerId: 'u_vik', deadline: D(6) },

  // ===== Initiatives (level 2) — VP-owned, carry the result metric + a deadline =====
  { id: 'w1', parentId: 'o1', level: 'initiative', title: 'Replace retired laptops (Oct-24 & older)', type: 'procurement', ownerId: 'u_pri', teamId: 't_it', scope: 'group', deadline: D(25), result: { metric: 'Laptops issued', unit: 'count', baseline: 0, target: 120, current: 55 } },
  { id: 'w1x', parentId: 'o1', level: 'initiative', title: 'Roll out MFA to field staff', type: 'compliance', ownerId: 'u_pri', teamId: 't_it', scope: 'group', deadline: D(35), result: { metric: 'Coverage', unit: '%', baseline: 0, target: 100, current: 60 } },
  { id: 'w2', parentId: 'o2', level: 'initiative', title: 'Reduce logistics cost per bag by ₹0.50', type: 'cost', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', deadline: D(50), result: { metric: 'Cost saved', unit: '₹/bag', baseline: 0, target: 0.5, current: 0.28 } },
  { id: 'w2b', parentId: 'o2', level: 'initiative', title: 'Warehouse automation pilot', type: 'general', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', deadline: D(55), result: { metric: 'Throughput gain', unit: '%', baseline: 0, target: 100, current: 30 } },
  { id: 'w3', parentId: 'o3', level: 'initiative', title: 'Plant routines & statutory upkeep', type: 'compliance', ownerId: 'u_mee', teamId: 't_sc', scope: 'group', deadline: D(28), result: { metric: 'Plant uptime', unit: '%', baseline: 0, target: 95, current: 91 } },
  { id: 'w3b', parentId: 'o3', level: 'initiative', title: 'FY24 statutory audit closure', type: 'compliance', ownerId: 'u_ana', teamId: 't_fin', scope: 'group', deadline: D(-3), result: { metric: 'Audit items closed', unit: '%', baseline: 0, target: 100, current: 100 } },
  { id: 'w4', parentId: 'o4', level: 'initiative', title: 'Launch dealer self-service portal', type: 'onboarding', ownerId: 'u_kar', teamId: 't_com', scope: 'group', deadline: D(60), result: { metric: 'Dealer adoption', unit: '%', baseline: 0, target: 100, current: 15 } },
  { id: 'w4b', parentId: 'o4', level: 'initiative', title: 'Reduce order-to-dispatch time', type: 'onboarding', ownerId: 'u_kar', teamId: 't_com', scope: 'group', deadline: D(40), result: { metric: 'Cycle time', unit: 'hrs', baseline: 48, target: 24, current: 40 } },
  { id: 'w5', parentId: 'o5', level: 'initiative', title: 'Roll out performance & L&D framework', type: 'compliance', ownerId: 'u_raj', teamId: 't_fin', scope: 'group', deadline: D(5), result: { metric: 'Coverage', unit: '%', baseline: 0, target: 100, current: 100 } },

  // ===== Works (level 3) — leaf work packages; hold the activities + a deadline =====
  // o1 / w1 — Replace retired laptops
  { id: 'w1W1', parentId: 'w1', level: 'work', title: 'Demand & selection', type: 'procurement', ownerId: 'u_roh', deadline: D(3) },
  { id: 'w1W2', parentId: 'w1', level: 'work', title: 'Procurement', type: 'procurement', ownerId: 'u_pri', deadline: D(12) },
  { id: 'w1W3', parentId: 'w1', level: 'work', title: 'Rollout', type: 'procurement', ownerId: 'u_roh', deadline: D(22) },
  // o1 / w1x — Roll out MFA
  { id: 'w1xW1', parentId: 'w1x', level: 'work', title: 'Assess & pilot', type: 'compliance', ownerId: 'u_arj', deadline: D(8) },
  { id: 'w1xW2', parentId: 'w1x', level: 'work', title: 'Enforce & monitor', type: 'compliance', ownerId: 'u_yash', deadline: D(30) },
  // o2 / w2 — Reduce logistics cost
  { id: 'w2W1', parentId: 'w2', level: 'work', title: 'Freight optimisation', type: 'cost', ownerId: 'u_neh', deadline: D(18) },
  { id: 'w2W2', parentId: 'w2', level: 'work', title: 'Commercials', type: 'cost', ownerId: 'u_mee', deadline: D(48) },
  // o2 / w2b — Warehouse automation pilot
  { id: 'w2bW1', parentId: 'w2b', level: 'work', title: 'Vendor selection', type: 'general', ownerId: 'u_sam', deadline: D(20) },
  { id: 'w2bW2', parentId: 'w2b', level: 'work', title: 'Pilot design', type: 'general', ownerId: 'u_div', deadline: D(40) },
  // o3 / w3 — Plant routines & statutory upkeep
  { id: 'w3W1', parentId: 'w3', level: 'work', title: 'Monthly payroll run', type: 'general', ownerId: 'u_ana', deadline: D(6), recurring: { cadence: 'monthly' } },
  { id: 'w3W2', parentId: 'w3', level: 'work', title: 'Kiln & raw-mill preventive maintenance', type: 'general', ownerId: 'u_sam', deadline: D(12), recurring: { cadence: 'monthly' } },
  { id: 'w3W3', parentId: 'w3', level: 'work', title: 'Statutory & pollution-control compliance', type: 'compliance', ownerId: 'u_div', deadline: D(26), recurring: { cadence: 'quarterly' } },
  // o3 / w3b — FY24 statutory audit closure (complete)
  { id: 'w3bW1', parentId: 'w3b', level: 'work', title: 'Statutory audit', type: 'compliance', ownerId: 'u_ana', deadline: D(-5) },
  // o4 / w4 — Dealer self-service portal
  { id: 'w4W1', parentId: 'w4', level: 'work', title: 'Requirements & scope', type: 'onboarding', ownerId: 'u_tan', deadline: D(5) },
  { id: 'w4W2', parentId: 'w4', level: 'work', title: 'Build MVP', type: 'onboarding', ownerId: 'u_tan', deadline: D(45) },
  // o4 / w4b — Reduce order-to-dispatch time
  { id: 'w4bW1', parentId: 'w4b', level: 'work', title: 'Process mapping', type: 'onboarding', ownerId: 'u_tan', deadline: D(0) },
  { id: 'w4bW2', parentId: 'w4b', level: 'work', title: 'Quick wins', type: 'onboarding', ownerId: 'u_neh', deadline: D(20) },
  // o5 / w5 — Performance & L&D framework (complete)
  { id: 'w5W1', parentId: 'w5', level: 'work', title: 'Competency framework', type: 'compliance', ownerId: 'u_raj', deadline: D(-4) },
  { id: 'w5W2', parentId: 'w5', level: 'work', title: 'L&D calendar', type: 'general', ownerId: 'u_raj', deadline: D(2) },
];

// Activities (level 4) hang directly off a work. `description` is the spec of what
// to produce (used when the AI scores the deliverable); `date` is the due date and
// `plannedHrs` the duration.
const ACTIVITIES = [
  // w1W1 — Demand & selection (done, one due-soon)
  { id: 'a1', workId: 'w1W1', title: 'Draft survey form', description: 'A short staff survey capturing current device, role, and mobility needs to size the refresh.', assigneeId: 'u_roh', date: D(-8), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'a2', workId: 'w1W1', title: 'Send survey to all staff', description: 'Distribute the survey to all in-scope staff and confirm delivery.', assigneeId: 'u_roh', date: D(-7), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'self' },
  { id: 'a3', workId: 'w1W1', title: 'Collate responses', description: 'Consolidate survey responses into a single dataset and flag gaps.', assigneeId: 'u_roh', date: D(-4), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'a3b', workId: 'w1W1', title: 'Prepare survey summary', description: 'A one-page demand summary: counts by model/spec and recommended quantities.', assigneeId: 'u_roh', date: D(-3), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'self' },
  { id: 'a4', workId: 'w1W1', title: 'Gather target specs', description: 'Define the standard laptop spec sheet (CPU/RAM/SSD/warranty).', assigneeId: 'u_roh', date: D(-5), status: 'executed', plannedHrs: 2, actualHrs: 3, actType: 'self' },
  { id: 'a5', workId: 'w1W1', title: 'Get 3 vendor quotes', description: 'Obtain three comparable vendor quotes against the spec sheet.', assigneeId: 'u_roh', date: D(2), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'call' },
  // w1W2 — Procurement (blocked, no overdue)
  { id: 'a6', workId: 'w1W2', title: 'Prepare BOM + cost sheet', description: 'Bill of materials and total cost sheet for the approved quantity.', assigneeId: 'u_roh', date: D(5), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'a7', workId: 'w1W2', title: 'HOD sign-offs (IT, Fin, Ops)', description: 'Signed budget approval from IT, Finance, and Operations heads.', assigneeId: 'u_pri', date: D(2), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'meeting' },
  { id: 'a8', workId: 'w1W2', title: 'Raise purchase requisition', description: 'PR raised in the procurement system referencing the approved cost sheet.', assigneeId: 'u_roh', date: D(4), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self' },
  { id: 'a9', workId: 'w1W2', title: 'Issue PO to vendor', description: 'Purchase order issued to the selected vendor with delivery terms.', assigneeId: 'u_roh', date: D(6), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'self', blocked: true },
  // w1W3 — Rollout (on track)
  { id: 'ae1', workId: 'w1W3', title: 'Image devices', description: 'Standard corporate image applied and validated on each device.', assigneeId: 'u_roh', date: D(10), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'ae2', workId: 'w1W3', title: 'Hand over to users', description: 'Devices handed over with sign-off; old devices collected.', assigneeId: 'u_roh', date: D(14), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'site' },
  { id: 'ae3', workId: 'w1W3', title: 'Update asset register', description: 'Asset register reconciled with issued/retired serials.', assigneeId: 'u_kav', date: D(16), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  // w1xW1 — Assess & pilot (one overdue)
  { id: 'ax1', workId: 'w1xW1', title: 'Inventory apps & auth methods', description: 'Inventory of applications and their current authentication methods.', assigneeId: 'u_arj', date: D(-3), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'ax2', workId: 'w1xW1', title: 'Map risky logins', description: 'Report of high-risk login patterns to prioritise for MFA.', assigneeId: 'u_arj', date: D(-2), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'ax3', workId: 'w1xW1', title: 'Pilot MFA with IT team', description: 'Run an MFA pilot with the IT team and capture issues.', assigneeId: 'u_arj', date: D(1), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  // w1xW2 — Enforce & monitor (on track)
  { id: 'ax4', workId: 'w1xW2', title: 'Roll out MFA to field staff', description: 'Enforce MFA for field staff in waves with support cover.', assigneeId: 'u_yash', date: D(20), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'self' },
  { id: 'ax5', workId: 'w1xW2', title: 'Build adoption dashboard', description: 'A dashboard tracking MFA coverage by function and site.', assigneeId: 'u_yash', date: D(26), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  // w2W1 — Freight optimisation (progressing, no overdue)
  { id: 'a10', workId: 'w2W1', title: 'Map current lane costs', description: 'Baseline cost-per-lane map across primary freight routes.', assigneeId: 'u_neh', date: D(-8), status: 'executed', plannedHrs: 3, actualHrs: 3, actType: 'self' },
  { id: 'a11', workId: 'w2W1', title: 'Model route consolidation', description: 'A model quantifying savings from consolidating overlapping routes.', assigneeId: 'u_neh', date: D(6), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'self' },
  { id: 'a12', workId: 'w2W1', title: 'Pull load-factor data by plant', description: 'Load-factor dataset by plant to find under-utilised trucks.', assigneeId: 'u_neh', date: D(-6), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'a13', workId: 'w2W1', title: 'Negotiate with 3PL carriers', description: 'Rate revisions agreed with third-party carriers on top lanes.', assigneeId: 'u_neh', date: D(12), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'call' },
  // w2W2 — Commercials (on track)
  { id: 'a14', workId: 'w2W2', title: 'Renegotiate transporter contracts', description: 'Updated transporter contracts reflecting the new rate card.', assigneeId: 'u_mee', date: D(20), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'meeting' },
  { id: 'a15', workId: 'w2W2', title: 'Board approval for savings plan', description: 'Board note and approval for the annualised savings plan.', assigneeId: 'u_mee', date: D(40), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'meeting' },
  // w2bW1 — Vendor selection (one due-soon)
  { id: 'b1', workId: 'w2bW1', title: 'Draft automation RFP', description: 'RFP for the warehouse automation pilot with scope and criteria.', assigneeId: 'u_sam', date: D(-3), status: 'executed', plannedHrs: 3, actualHrs: 3, actType: 'self' },
  { id: 'b2', workId: 'w2bW1', title: 'Shortlist 3 vendors', description: 'Evaluate responses and shortlist three vendors for site visits.', assigneeId: 'u_sam', date: D(2), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'b3', workId: 'w2bW1', title: 'Vendor site visits', description: 'Reference site visits with a scoring sheet per vendor.', assigneeId: 'u_sam', date: D(15), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'site' },
  // w2bW2 — Pilot design (on track)
  { id: 'b4', workId: 'w2bW2', title: 'Define pilot KPIs', description: 'KPIs and success thresholds for the automation pilot.', assigneeId: 'u_div', date: D(18), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'b5', workId: 'w2bW2', title: 'Draft pilot rollout plan', description: 'Rollout plan with milestones, owners, and go/no-go gates.', assigneeId: 'u_div', date: D(30), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  // w3W1 — Monthly payroll run (complete)
  { id: 'ap1', workId: 'w3W1', title: 'Validate attendance & inputs', description: 'Validated attendance and variable-pay inputs for the payroll cycle.', assigneeId: 'u_ana', date: D(-4), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'ap2', workId: 'w3W1', title: 'Process & disburse salaries', description: 'Payroll processed and disbursed; bank confirmation attached.', assigneeId: 'u_ana', date: D(-2), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'ap3', workId: 'w3W1', title: 'Distribute payslips', description: 'Payslips generated and distributed to all employees.', assigneeId: 'u_ana', date: D(-1), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'self' },
  // w3W2 — Kiln & raw-mill PM (progressing)
  { id: 'am1', workId: 'w3W2', title: 'Inspect refractory & mill bearings', description: 'Inspection report on refractory lining and mill bearing condition.', assigneeId: 'u_sam', date: D(-2), status: 'executed', plannedHrs: 4, actualHrs: 4, actType: 'site' },
  { id: 'am2', workId: 'w3W2', title: 'Lubrication & alignment', description: 'Lubrication and alignment completed with readings logged.', assigneeId: 'u_sam', date: D(3), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'site' },
  { id: 'am3', workId: 'w3W2', title: 'Vibration analysis', description: 'Vibration analysis on critical drives with trend chart.', assigneeId: 'u_sam', date: D(8), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'site' },
  // w3W3 — Statutory & pollution compliance (on track)
  { id: 'ac1', workId: 'w3W3', title: 'Compile CPCB emission data', description: 'Compiled CPCB emission dataset for the quarter.', assigneeId: 'u_div', date: D(5), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'ac2', workId: 'w3W3', title: 'File statutory returns', description: 'Statutory environmental returns filed for the quarter.', assigneeId: 'u_div', date: D(12), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  { id: 'ac3', workId: 'w3W3', title: 'Renew consent-to-operate', description: 'Consent-to-operate renewal submitted with supporting documents.', assigneeId: 'u_div', date: D(22), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  // w3bW1 — Statutory audit (complete → green)
  { id: 'd1', workId: 'w3bW1', title: 'Draft FY24 financials', description: 'Draft FY24 financial statements for audit.', assigneeId: 'u_ana', date: D(-15), status: 'executed', plannedHrs: 4, actualHrs: 5, actType: 'self' },
  { id: 'd2', workId: 'w3bW1', title: 'Resolve auditor queries', description: 'All auditor queries addressed and evidence provided.', assigneeId: 'u_ana', date: D(-10), status: 'executed', plannedHrs: 3, actualHrs: 4, actType: 'meeting' },
  { id: 'd3', workId: 'w3bW1', title: 'Management representation letter', description: 'Signed management representation letter to auditors.', assigneeId: 'u_ana', date: D(-7), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'self' },
  { id: 'd4', workId: 'w3bW1', title: 'Board sign-off', description: 'Board approval of audited financials.', assigneeId: 'u_ana', date: D(-5), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'meeting' },
  // w4W1 — Requirements & scope (one due-soon)
  { id: 'e1', workId: 'w4W1', title: 'Dealer interviews', description: 'Interview a representative set of dealers on portal needs.', assigneeId: 'u_tan', date: D(-4), status: 'executed', plannedHrs: 3, actualHrs: 3, actType: 'call' },
  { id: 'e2', workId: 'w4W1', title: 'Requirements document', description: 'A prioritised requirements document for the portal MVP.', assigneeId: 'u_tan', date: D(1), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'e3', workId: 'w4W1', title: 'Scope sign-off', description: 'Signed-off MVP scope from Commercial and IT.', assigneeId: 'u_tan', date: D(4), status: 'planned', plannedHrs: 1, actualHrs: null, actType: 'meeting' },
  // w4W2 — Build MVP (blocked)
  { id: 'e4', workId: 'w4W2', title: 'Wireframes & UX', description: 'Clickable wireframes for the core dealer journeys.', assigneeId: 'u_tan', date: D(15), status: 'planned', plannedHrs: 4, actualHrs: null, actType: 'self' },
  { id: 'e5', workId: 'w4W2', title: 'Vendor build kickoff', description: 'Kick off the build with the selected vendor.', assigneeId: 'u_tan', date: D(25), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'meeting', blocked: true },
  // w4bW1 — Process mapping (overdue)
  { id: 'f1', workId: 'w4bW1', title: 'Map current order-to-dispatch flow', description: 'End-to-end map of the current order-to-dispatch process.', assigneeId: 'u_tan', date: D(-3), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'f2', workId: 'w4bW1', title: 'Identify bottlenecks', description: 'Ranked list of bottlenecks with time impact.', assigneeId: 'u_neh', date: D(-1), status: 'planned', plannedHrs: 2, actualHrs: null, actType: 'self' },
  // w4bW2 — Quick wins (one overdue, one on track)
  { id: 'f3', workId: 'w4bW2', title: 'Automate order confirmations', description: 'Auto-confirmation emails on order capture.', assigneeId: 'u_neh', date: D(10), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  { id: 'f4', workId: 'w4bW2', title: 'Build SLA dashboard', description: 'Dashboard tracking order-to-dispatch SLA by region.', assigneeId: 'u_tan', date: D(-1), status: 'planned', plannedHrs: 3, actualHrs: null, actType: 'self' },
  // w5W1 — Competency framework (complete)
  { id: 'g1', workId: 'w5W1', title: 'Draft role competencies', description: 'Competency definitions for key roles.', assigneeId: 'u_raj', date: D(-6), status: 'executed', plannedHrs: 3, actualHrs: 3, actType: 'self' },
  { id: 'g2', workId: 'w5W1', title: 'Leadership review', description: 'Leadership review and approval of the competency framework.', assigneeId: 'u_raj', date: D(-2), status: 'executed', plannedHrs: 1, actualHrs: 1, actType: 'meeting' },
  // w5W2 — L&D calendar (complete)
  { id: 'g3', workId: 'w5W2', title: 'Training needs survey', description: 'Consolidated training-needs survey across functions.', assigneeId: 'u_raj', date: D(-3), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
  { id: 'g4', workId: 'w5W2', title: 'Publish annual L&D calendar', description: 'Published annual learning & development calendar.', assigneeId: 'u_raj', date: D(-1), status: 'executed', plannedHrs: 2, actualHrs: 2, actType: 'self' },
];

const CRS = [
  { id: 'cr1', workId: 'w1', targetWorkId: 'w1W1', proposerId: 'u_roh', kind: 'add_activity', desc: 'Found 3 depots missed in the first survey — need a quick re-run.', payload: { title: 'Re-run survey for 3 missed depots', hrs: 2, type: 'self' }, status: 'pending' },
  { id: 'cr2', workId: 'w2', targetWorkId: 'w2W1', proposerId: 'u_neh', kind: 'extend', desc: 'Route model needs another pass after the new lane data landed.', payload: { activityId: 'a11', hrs: 2 }, status: 'pending' },
  { id: 'cr3', workId: 'w4', targetWorkId: 'w4W1', proposerId: 'u_tan', kind: 'add_activity', desc: 'Dealers asked for a mobile view — add a scoping task.', payload: { title: 'Scope mobile dealer view', hrs: 3, type: 'self' }, status: 'pending' },
];

// Portal-only cross-user nudges. The bot ignores these, but they are persisted
// so multiple portal sessions stay consistent.
const REMARKS = [];

module.exports = { USERS, TEAMS, WORKS, ACTIVITIES, CRS, REMARKS, MSD, sod, addDays, iso, TODAY };
