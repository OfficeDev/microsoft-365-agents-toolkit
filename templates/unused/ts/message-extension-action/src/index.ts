// Import required packages
import {
  AuthConfiguration,
  authorizeJWT,
  CloudAdapter,
  loadAuthConfigFromEnv,
  TurnContext,
} from "@microsoft/agents-hosting";
import express from "express";
import { ActionApp } from "./actionApp";

// Create authentication configuration
const authConfig: AuthConfiguration = loadAuthConfigFromEnv();

// Create adapter
const adapter = new CloudAdapter(authConfig);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the app that will handle action commands.
const actionApp = new ActionApp();

// Create express application.
const server = express();
server.use(express.json());
server.use(authorizeJWT(authConfig));

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await actionApp.run(context);
  });
});

// Start the server
const port = process.env.PORT || 3978;
server
  .listen(port, () => {
    console.log(
      `Bot Started, listening to port ${port} for appId ${authConfig.clientId} debug ${process.env.DEBUG}`
    );
  })
  .on("error", (err) => {
    console.error(err);
    process.exit(1);
  });
