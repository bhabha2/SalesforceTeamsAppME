// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
// index.js is used to setup and configure your bot
// Import required pckages
const path = require('path');
// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

const restify = require('restify');
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const { CloudAdapter, ConfigurationServiceClientCredentialFactory, ConfigurationBotFrameworkAuthentication,
  ConversationState, UserState, MemoryStorage} = require("botbuilder");
// const { TeamsMessagingExtensionsSearchAuthConfigBot } = require('./bots/teamsMessagingExtensionsSearchAuthConfigBot');
const { SearchApp } = require('./searchApp');

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppType: "MultiTenant"
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {}, 
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
await context.sendTraceActivity(
  'OnTurnError Trace',
  `${ error }`,
  'https://www.botframework.com/schemas/error',
  'TurnError'
);

// Uncomment below commented line for local debugging.
await context.sendActivity(`Sorry, it looks like something went wrong. Exception Caught: ${error}`);
};

const memoryStorage = new MemoryStorage();
const userState = new UserState(memoryStorage);
const conversationState = new ConversationState(memoryStorage);
// Create the bot that will handle incoming messages.
// const bot = new TeamsMessagingExtensionsSearchAuthConfigBot(userState);
// const bot = new SearchApp(userState);
const searchApp = new SearchApp(conversationState, userState);

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`\nBot started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await searchApp.run(context);
  });
});

// Serve up static files in the public directory (namely: searchSettings.html)
// server.get('/public/*', restify.plugins.serveStatic({
//     directory: __dirname
// }));