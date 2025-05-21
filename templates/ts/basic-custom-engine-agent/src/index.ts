// Import required packages
import {
  AuthConfiguration,
  authorizeJWT,
  loadAuthConfigFromEnv,
  Request,
} from "@microsoft/agents-hosting";
import express, { Response } from "express";

// This bot's adapter
import adapter from "./adapter";

// This bot's main dialog.
import { agentApp } from "./agent";

// Create authentication configuration
const authConfig: AuthConfiguration = loadAuthConfigFromEnv();

// Create express application.
const expressApp = express();
expressApp.use(express.json());
expressApp.use(authorizeJWT(authConfig));

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nAgent started, ${expressApp.name} listening to`, server.address());
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await agentApp.run(context);
  });
});
