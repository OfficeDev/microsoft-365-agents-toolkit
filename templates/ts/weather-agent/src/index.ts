// Import required packages
import {
  AuthConfiguration,
  authorizeJWT,
  CloudAdapter,
  loadAuthConfigFromEnv,
  Request,
} from "@microsoft/agents-hosting";
import express, { Response } from "express";

// This bot's main dialog.
import { weatherAgent } from "./agent";

// Create authentication configuration
const authConfig: AuthConfiguration = loadAuthConfigFromEnv();
const adapter = new CloudAdapter(authConfig);

// Create express application.
const server = express();
server.use(express.json());
server.use(authorizeJWT(authConfig));

// Listen for incoming requests.
server.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await weatherAgent.run(context);
  });
});

const port = process.env.port || process.env.PORT || 3978;
server.listen(port, () => {
  console.log(
    `\nServer listening to port ${port} for appId ${authConfig.clientId} debug ${process.env.DEBUG}`
  );
});
