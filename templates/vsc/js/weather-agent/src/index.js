// Import required packages
const { authorizeJWT, CloudAdapter, loadAuthConfigFromEnv } = require("@microsoft/agents-hosting");
const express = require("express");

// This bot's main dialog.
const { weatherAgent } = require("./agent");

// Create authentication configuration
const authConfig = loadAuthConfigFromEnv();
const adapter = new CloudAdapter(authConfig);

// Create express application.
const server = express();
server.use(express.json());
server.use(authorizeJWT(authConfig));

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
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
