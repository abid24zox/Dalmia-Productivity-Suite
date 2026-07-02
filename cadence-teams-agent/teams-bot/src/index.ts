// Host process: restify server + Bot Framework CloudAdapter. Uses your existing
// Azure Bot Service registration (App ID / password / tenant from env).
import 'dotenv/config';
import * as restify from 'restify';
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
} from 'botbuilder';
import { CadenceBot } from './bot';

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: process.env.MICROSOFT_APP_TYPE || 'MultiTenant',
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
});
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({}, credentialsFactory);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context: TurnContext, error: Error) => {
  console.error('[onTurnError]', error);
  await context.sendActivity('The Cadence agent hit an error. Please try again.');
};

const bot = new CadenceBot();

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res as any, (context) => bot.run(context));
});

server.get('/healthz', (_req, res, next) => { res.send(200, { status: 'ok' }); return next(); });

const port = process.env.PORT || 3978;
server.listen(port, () => console.log(`Cadence Teams bot listening on ${port}`));
