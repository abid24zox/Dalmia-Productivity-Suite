// Standalone host: restify server + Bot Framework CloudAdapter. Uses your Azure
// Bot Service registration (App ID / password / tenant from env). This runs the
// bot on its OWN process/port — used for local dev against an ngrok tunnel. In
// production the bot is instead mounted in-process by the merged Cadence server
// (see cadence-service/server.js → teams-bot/dist/host registerBot), so a single
// Azure App Service serves the portal, the API, and /api/messages together.
import 'dotenv/config';
import * as restify from 'restify';
import { buildBot } from './host';

const { adapter, bot } = buildBot();

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', async (req, res) => {
  try {
    await adapter.process(req, res as any, (context) => bot.run(context));
  } catch (err: any) {
    console.error('[/api/messages] rejected:', err?.message || err);
    if (!res.headersSent) res.send(400, { error: 'invalid bot activity' });
  }
});

server.get('/healthz', (_req, res, next) => { res.send(200, { status: 'ok' }); return next(); });

const port = process.env.PORT || 3978;
server.listen(port, () => console.log(`Cadence Teams bot listening on ${port}`));
