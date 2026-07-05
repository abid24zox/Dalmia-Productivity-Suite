# Cadence — work & initiative OS

Cadence is an execution platform where enterprise work is planned and operated across a **web portal** and a **Microsoft Teams agent** that both talk to one shared backend — so anything done in Teams shows up in the portal, and vice-versa, live.

```
   Microsoft Teams              Teams bot (Node/TS)          Shared Cadence service           Cadence portal
  ┌──────────────┐   Bot Svc   ┌──────────────────┐  HTTPS  ┌────────────────────────┐       ┌──────────────┐
  │ chat + cards │◀──────────▶│ Foundry tool loop  │───────▶│ REST + in-memory store  │◀─────▶│ React (Vite) │
  └──────────────┘ /api/messages└─────────┬────────┘  tools │ meters/scoping/AI/STT   │ poll  │  live portal  │
                                          │                 └────────────────────────┘       └──────────────┘
                                          ▼                   single source of truth
                                Azure AI Foundry model  ◀── same model powers the portal's in-house AI
```

## Layout

| Folder | What it is |
|---|---|
| [`cadence-portal/`](cadence-portal) | The **portal** — Vite + React + Tailwind. Live-wired to the service (loads a snapshot on login, persists every edit, polls every ~3.5s). See its [README](cadence-portal/README.md). |
| [`cadence-teams-agent/cadence-service/`](cadence-teams-agent/cadence-service) | The **shared service** — Express + in-memory store. Single source of truth for the 5-level model (`objective → initiative → work → sub-work → activity`). Also hosts the in-house **AI** (Azure Foundry) and **speech-to-text** (Deepgram). |
| [`cadence-teams-agent/teams-bot/`](cadence-teams-agent/teams-bot) | The **Teams bot** — Bot Framework + a Foundry tool-calling loop. Each tool maps to a service action and returns an Adaptive Card. See the [agent README](cadence-teams-agent/README.md). |
| `cadence-prototype-v13.jsx` | The original single-file portal prototype (reference only). `cadence-portal/src/App.jsx` is its live-wired evolution. |

## Run it locally (three terminals)

For local development each piece runs on its own port. (In production they run as **one process** — see [Deploy](#deploy--single-azure-app-service).)

```bash
# 1) shared service (source of truth + AI + STT)  → http://localhost:4000
cd cadence-teams-agent/cadence-service
npm install
cp .env.example .env      # fill AZURE_OPENAI_* and (optional) DEEPGRAM_API_KEY
npm start

# 2) portal                                        → http://localhost:5173
cd cadence-portal
npm install
npm run dev

# 3) Teams bot                                      → port 3978
cd cadence-teams-agent/teams-bot
npm install
cp .env.example .env      # Bot Service + Azure Foundry creds
npm run dev
```

Sign in to the portal with a test account, e.g. **`vikram / md@2026`** (MD), `priya / vp@2026` (VP IT), `rohit / team@2026` (member).

For the bot to receive Teams messages, expose port 3978 with a tunnel (e.g. `ngrok http 3978`) and set your Azure Bot's **messaging endpoint** to `https://<tunnel>/api/messages`.

## Deploy — single Azure App Service

In production the three pieces run as **one Node process**: the Express service *also* serves the built portal SPA and mounts the Teams bot at `/api/messages`. So a single App Service (one URL) hosts everything.

- **Entry point:** `node cadence-teams-agent/cadence-service/server.js` (root `npm start`).
- **Build:** root `npm run build` installs + builds the portal (`cadence-portal/dist`) and the bot (`teams-bot/dist`), then installs the service. `dist/` is git-ignored, so the host builds on deploy.
- One request map: `GET /` + `/assets/*` → portal · `/api/*` → REST + AI · `POST /api/messages` → bot.

### Steps

1. **Create an App Service** — Linux, **Node 20 LTS**, **Always On** = on (so the bot never sleeps).
2. **Deployment Center → GitHub →** branch **`deploy/single-appservice`**. Use the App Service (Oryx) build and add app setting `SCM_DO_BUILD_DURING_DEPLOYMENT=true`. Startup command: `npm start`. (Building on Azure is required — `dist/` is never committed.)
3. **App Settings** (env vars):
   - *Runtime (service + bot):* `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_DEPLOYMENT`, `AZURE_OPENAI_API_VERSION`, `DEEPGRAM_API_KEY`, `MICROSOFT_APP_ID`, `MICROSOFT_APP_PASSWORD`, `MICROSOFT_APP_TYPE`, `MICROSOFT_APP_TENANT_ID`.
   - *Build-time for the portal* (Vite bakes these into the bundle, so they must be set **before** the build): `VITE_MSAL_CLIENT_ID`, `VITE_MSAL_TENANT_ID`. Leave `VITE_CADENCE_API_URL` **unset** so the portal calls same-origin `/api`.
4. **Repoint the bot** — Azure Bot → messaging endpoint → `https://<app>.azurewebsites.net/api/messages`. And on the AAD app registration, add `https://<app>.azurewebsites.net` as an **SPA redirect URI** (so MSAL / OneDrive / Calendar work in prod).
5. **Verify** — `https://<app>.azurewebsites.net/api/health` returns `{status:"ok"}`, the site loads, and the Teams bot answers.

Keep the plan at a **single instance** — the store is in-memory (see Notes).

## How it stays consistent

- **One model, one store.** Both surfaces read/write the same service; the portal renders the raw tree and the bot reads summarized shapes — computed by identical logic so they never disagree.
- **One AI.** The portal's in-house AI, document extraction (PDF / DOCX / XLSX / CSV / EML / MSG / MD / TXT), and speech-to-text all run server-side through the same Azure Foundry deployment the bot uses. No model or STT keys in the browser.
- **Sync.** Optimistic local update → persist to the service → reconcile to the returned snapshot, plus a ~3.5s poll so cross-surface changes appear automatically.

## Notes

- The store is **in-memory and reseeds on restart** — swap `cadence-service/store.js` for a real DB without touching routes.
- `.env` files hold secrets and are **git-ignored**; commit only the `.env.example` templates.
- Voice uses **Deepgram** when `DEEPGRAM_API_KEY` is set, and falls back to the browser's Web Speech API otherwise.
