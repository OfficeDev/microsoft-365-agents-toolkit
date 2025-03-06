import {
  CloudAdapter,
  Request,
  TurnContext,
  authorizeJWT,
  loadBotAuthConfigFromEnv,
} from "@microsoft/agents-bot-hosting";
import express, { Response } from "express";
import rateLimit from "express-rate-limit";

// This bot's main dialog
import { TeamsBot } from "./teamsBot";

// Create authentication configuration
const authConfig = loadBotAuthConfigFromEnv();

// Create adapter
const adapter = new CloudAdapter(authConfig);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
  if (context.activity.type === "message") {
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
  }
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
const bot = new TeamsBot();

// Create express application with rate limiting
const app = express();
app.use(rateLimit({ validate: { xForwardedForHeader: false } }));
app.use(express.json());
app.use(authorizeJWT(authConfig));

// Listen for incoming requests.
app.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

// Start the server
const port = process.env.PORT || 3978;
app.listen(port, () => {
  console.log(`\napp listening to port ${port} for appId ${authConfig.clientId}`);
});
