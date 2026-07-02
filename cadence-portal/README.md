# Cadence — Portal (live)

The React portal, now wired to the shared **Cadence service** (`../cadence-teams-agent/cadence-service`) instead of in-memory seed data. Everything you do here persists to the service, and anything the Teams agent does shows up here within a few seconds (polling every ~3.5s + refetch-after-write).

## Run

```bash
npm install
npm run dev        # http://localhost:5173
```

The service must be running on `http://localhost:4000` first (see the service README). Configure the URL in `.env`:

```
VITE_CADENCE_API_URL=http://localhost:4000
VITE_CADENCE_API_KEY=        # only if the service sets CADENCE_API_KEY
```

Sign in with a test account, e.g. `vikram / md@2026` (MD), `priya / vp@2026` (VP IT), `rohit / team@2026` (member).

## How it connects

- `src/api.js` — thin client for the service (login, snapshot, granular writes, AI, extract, transcribe).
- `src/App.jsx` — the portal. On login it loads `/api/snapshot`; a small `store` layer does an **optimistic** local update, **persists** each change to the service, then **reconciles** to the returned snapshot. A poller keeps it in sync with Teams-side changes.
- **AI** (`decompose`, `modifyPlan`, `insight`, `score`, …) runs through the service's Azure AI Foundry model — no keys in the browser.
- **File upload** (Capture, Quick-create, Deliverable) accepts **PDF, Word (.docx), .md, .txt**; PDFs/Word are extracted to text server-side.
- **Voice** uses **Deepgram** when the service has a key, else the browser's Web Speech API.

`src/App.jsx` is a direct evolution of `cadence-prototype-v13.jsx` — every screen and feature is preserved; only the data source and AI transport changed.
