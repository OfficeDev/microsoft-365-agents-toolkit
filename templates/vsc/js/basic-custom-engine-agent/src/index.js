// Import required packages
const { authorizeJWT, loadAuthConfigFromEnv } = require("@microsoft/agents-hosting");
const express = require("express");

// This bot's adapter
const adapter = require("./adapter");

// This bot's main dialog.
const { agentApp } = require("./agent");

// Create authentication configuration
const authConfig = loadAuthConfigFromEnv();

// Create express application.
const expressApp = express();
expressApp.use(express.json());
expressApp.use(authorizeJWT(authConfig));

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nAgent started, ${expressApp.name} listening to`, server.address());
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await agentApp.run(context);
  });
});
