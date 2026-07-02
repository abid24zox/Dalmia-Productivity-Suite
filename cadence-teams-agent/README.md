# Cadence — Teams Agent

An AI agent that runs on **Microsoft Teams**, linked to the Cadence portal, so an executive (e.g. the MD) can **plan and operate work by chatting**. Actionables requested in Teams are executed against a **shared Cadence service**, which the portal reads from — so what you do in Teams shows up in the portal.

```
  Microsoft Teams                     Bot host (Node/TS)                 Shared Cadence service          Cadence portal
 ┌───────────────┐   Bot Service    ┌────────────────────┐   HTTPS     ┌──────────────────────┐        ┌──────────────┐
 │  MD chats +   │◀───────────────▶│  CloudAdapter       │            │  REST API + store     │◀──────▶│  React app    │
 │  Adaptive     │   /api/messages  │  ├ Foundry loop ────┼──tools────▶│  /api/initiatives ... │  reads │  (v5 proto)   │
 │  Cards        │                  │  ├ tools → cadence  │   calls    │  meters / scoping /   │        │               │
 └───────────────┘                  │  └ card actions     │            │  assignment           │        └──────────────┘
                                    └─────────┬───────────┘            └──────────────────────┘
                                              │ chat + tool calling
                                              ▼
                                    Azure AI Foundry model
                                    (Azure OpenAI-compatible)
```

Three packages (the portal now lives in a sibling folder, `../cadence-portal`):

- **`cadence-service/`** — the shared source of truth (Node/Express, in-memory store, the SAME 5-level object model + meters as the portal). Both the bot **and** the portal read and write here, so changes flow both ways. It also hosts the **in-house AI** (Azure AI Foundry — the same model the bot uses) and a **speech-to-text** proxy (Deepgram), so no model/STT keys ever reach the browser. Swap the store for Postgres/Prisma in production without touching routes.
- **`teams-bot/`** — the agent (TypeScript, Bot Framework SDK). Runs the MD's conversation through your Foundry model with **tool calling**; each tool maps to a Cadence API action and returns an **Adaptive Card**.
- **`../cadence-portal/`** — the React portal (Vite + Tailwind), now **live**: it loads a snapshot from the service on login, persists every edit back, and polls every ~3.5s so anything done in Teams appears in the portal (and vice-versa).

> Framework note: I built on the **Bot Framework SDK** with an **explicit Foundry tool-calling loop** (rather than the higher-level Teams AI planner) so the routing of every tool to the Cadence API is transparent and easy to extend. It runs on your existing Bot Service registration. Moving to the Teams AI `ActionPlanner` later is straightforward.

---

## What the agent can do (v1)

| Intent (say this in Teams) | Tool | Cadence API |
|---|---|---|
| "How's the portfolio / what's at risk?" | `get_portfolio` | `GET /api/portfolio` |
| "Status of the laptop refresh?" | `get_initiative_status` | `GET /api/initiatives/:id` |
| "Plan an MFA rollout, assign to IT Ops, deadline in 2 weeks" | `plan_initiative` | `POST /api/initiatives` |
| "Schedule *Issue PO* to Rohit on Friday" | `schedule_activity` | `POST /api/activities/:id/schedule` |
| "Move *Collate responses* to Neha" | `reassign_activity` | `POST /api/activities/:id/reassign` |
| "What needs attention?" | `list_attention` | `GET /api/attention` |
| "Show pending approvals" (+ Approve/Reject buttons) | `list_approvals` / card action | `GET /api/approvals`, `POST /api/approvals/:id/decide` |
| "Who has capacity this week?" | `team_capacity` | `GET /api/capacity` |
| "List the teams" | `list_teams` | `GET /api/teams` |

When planning, the **model decomposes** the goal into Objective → Initiative → Work → Sub-work → Activity, shows the plan, asks for a yes, then the **service assigns** the activities across the chosen team balanced by load and dates them to the deadline — identical logic to the portal's capture flow.

Everything is **scoped to the caller**: MD sees the enterprise, a VP sees their function, a member sees only their own work.

---

## Setup

### 1) Shared Cadence service
```bash
cd cadence-service
npm install
npm start                 # http://localhost:4000
curl localhost:4000/api/health
```
(Optional) set `CADENCE_API_KEY` to require an `x-api-key` header on calls.

### 2) Teams bot
```bash
cd teams-bot
npm install
cp .env.example .env      # fill in the values below
npm run dev               # ts-node, or: npm run build && npm start
```

Fill `.env`:
- **Bot Service** (you already have this): `MICROSOFT_APP_ID`, `MICROSOFT_APP_PASSWORD`, `MICROSOFT_APP_TENANT_ID`, `MICROSOFT_APP_TYPE`.
- **Azure AI Foundry** (Azure OpenAI-compatible deployment with tool calling): `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_DEPLOYMENT`, `AZURE_OPENAI_API_VERSION`.
- **Cadence**: `CADENCE_API_URL` (default `http://localhost:4000`), optional `CADENCE_API_KEY`.
- **`PORTAL_URL`** for the "Open in portal" card button; **`DEFAULT_CADENCE_USER_ID`** demo fallback identity.

### 3) Connect to Teams
- Point your **Bot Service** messaging endpoint at `https://<your-host>/api/messages` (use a dev tunnel / ngrok locally).
- Sideload **`teams-bot/appManifest/manifest.json`** — replace `botId` with your Bot Service App ID and add `color.png` / `outline.png` icons. (You said the Teams app already exists — just point it at this bot.)

---

## Identity mapping

`getCadenceUserId` reads the Teams member's email/UPN and matches it to a Cadence user (`GET /api/resolve?kind=user&q=<email>`). If it can't (personal scope, emulator, or unlinked), it falls back to `DEFAULT_CADENCE_USER_ID` (the MD) so the demo works. For production, ensure Cadence user emails match AAD.

---

## Cadence API contract (v1)

```
GET  /api/health
GET  /api/users
GET  /api/teams
POST /api/teams                         { name, memberIds[] }               -> create/merge a team
GET  /api/portfolio?userId=             -> { tiles, initiatives[] }         (scoped)
GET  /api/initiatives                   -> { initiatives[] }                (scoped)
GET  /api/initiatives/:idOrTitle        -> full detail + subworks/activities
POST /api/initiatives                   { ownerId, title, type, objective?, deadline?, teamId, subworks[] }
POST /api/activities/:id/schedule       { assigneeId?, date? }
POST /api/activities/:id/reassign       { assigneeId }
GET  /api/attention?userId=             -> overdue / blocked / stuck
GET  /api/capacity?userId=              -> open hours per person
GET  /api/approvals?userId=             -> pending change requests
POST /api/approvals/:id/decide          { approve, remark?, spinoff?, approverId? }
GET  /api/resolve?kind=&q=              -> { match: { id, name } }          (user|team|initiative|activity)
```

`subworks[]` shape for `POST /api/initiatives`:
```json
[{ "title": "Assess current auth",
   "activities": [{ "title": "Inventory apps", "estimateHrs": 3, "type": "self" }] }]
```

---

## Example conversation (MD)

> **MD:** Plan a rollout of MFA to all field staff. Assign it to IT Ops, deadline in two weeks.
>
> **Agent:** Here's a plan — *Assess current auth* (inventory apps, map risky logins), *Pilot rollout* (enroll pilot, collect feedback), *Enforce & monitor* (enable policy, exception handling). Assign to IT Ops (Priya, Rohit, Arjun, Kavya), due 15 Jul? *(shows plan)*
>
> **MD:** Yes.
>
> **Agent:** ✅ Created & assigned — 3 sub-works, 6 activities distributed across IT Ops by load. *(createdCard, "View in portal")*
>
> **MD:** What's at risk right now?
>
> **Agent:** *(attentionCard)* Collate responses is overdue; the laptop refresh is stuck at Budget approvals. Want me to nudge the owner or reassign?

---

## Running the full system (portal + service + bot)

Three terminals:

```bash
# 1) shared service (source of truth + AI + STT)
cd cadence-teams-agent/cadence-service
npm install
cp .env.example .env      # fill Azure Foundry creds (copy from teams-bot/.env) + DEEPGRAM_API_KEY (optional, add later)
npm start                 # http://localhost:4000   (health shows foundry/deepgram flags)

# 2) portal (live React app)
cd cadence-portal
npm install
npm run dev               # http://localhost:5173   (VITE_CADENCE_API_URL points at :4000)

# 3) Teams bot
cd cadence-teams-agent/teams-bot
npm install
npm run dev               # ts-node; point your Bot Service messaging endpoint at /api/messages
```

Sign in to the portal with any test account (e.g. `vikram / md@2026`). Create/plan work in the portal or via the bot — it shows up on the other side within a few seconds.

## Model notes

- **One object model.** The service now carries the portal's full 5-level tree (`objective → initiative → work → sub-work → activity` with an explicit `level`) and identical meter/RAG/sufficiency logic, so the portal and the agent always agree.
- **One AI.** The portal's in-house AI (`POST /api/ai/complete`), document extraction (`/api/ai/extract` — pdf/docx/md/txt), and speech-to-text (`/api/ai/transcribe`) all run server-side through the same Foundry deployment the bot uses. `gpt-5.x` minis need `max_completion_tokens` (already handled).
- **Voice.** The portal records audio and transcribes via **Deepgram** when `DEEPGRAM_API_KEY` is set; until then it falls back to the browser's Web Speech API automatically. Drop the key into `cadence-service/.env` and it upgrades with no code change.

## What's mocked / next

- **Store is in-memory** and reseeds on restart — swap `cadence-service/store.js` for a real DB. Routes stay the same.
- **Live sync** is snapshot polling every ~3.5s + refetch-after-write. Swap for SSE/WebSocket if you want instant push.
- **Auth:** service supports a shared `x-api-key` (login/health/AI stay open); portal login is server-checked but passwords are demo-plaintext. Add AAD/JWT + hashing for production. The bot uses your Bot Service auth already.
- **Foundry:** requires a deployment that supports tool/function calling. Credentials via env only — nothing hardcoded.
