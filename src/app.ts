import * as restify from 'restify';
import { BotFrameworkAdapter, ConfigurationServiceClientCredentialFactory,  CloudAdapter, ConversationState,MemoryStorage } from 'botbuilder';
import { TeamsBot } from './bot';
const ngrok = require('ngrok');

const server = restify.createServer();
server.listen(3000, () => {
    console.log(`${server.name} listening to ${server.url}`);
});

ngrok.connect(3000).then((url: any) => console.log(`Public Url: ${url}`));


// const adapter = new BotFrameworkAdapter({  // BotFrameworkAdapter is depricated and will be removed in future versions by CloudAdapter 
//     appId: process.env.MicrosoftAppId,
//     appPassword: process.env.MicrosoftAppPassword
// });

const adapter = new CloudAdapter(
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.MicrosoftAppId,
        MicrosoftAppPassword: process.env.MicrosoftAppPassword,
        MicrosoftAppType: process.env.MicrosoftAppType,
        MicrosoftAppTenantId: process.env.MicrosoftAppTenantId
    }) as any
);

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);

const bot = new TeamsBot(conversationState);

server.post('/api/messages', (req, res, next) => {
    (adapter as any).processActivity(req, res, async (context: any) => {
        await bot.run(context);
    }).then(() => next()).catch((err: any) => next(err));
});


