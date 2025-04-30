// Import required packages
const express = require("express");

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const { authorizeJWT, CloudAdapter, loadAuthConfigFromEnv } = require("@microsoft/agents-hosting");
const { LinkUnfurlingApp } = require("./linkUnfurlingApp");

// Create authentication configuration
const authConfig = loadAuthConfigFromEnv();

// Create adapter
const adapter = new CloudAdapter(authConfig);

adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a message to the user
  await context.sendActivity(`The bot encountered an unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Create the bot that will handle incoming messages.
const linkUnfurlingApp = new LinkUnfurlingApp();

// Create express application.
const server = express();
server.use(express.json());
server.use(authorizeJWT(authConfig));

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await linkUnfurlingApp.run(context);
  });
});

// Start the server
const port = process.env.PORT || 3978;
server
  .listen(port, () => {
    console.log(
      `\napp listening to port ${port} for appId ${authConfig.clientId} debug ${process.env.DEBUG}`
    );
  })
  .on("error", (err) => {
    console.error(err);
    process.exit(1);
  });
