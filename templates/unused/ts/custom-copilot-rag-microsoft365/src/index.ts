// Import required packages
import express from "express";

// This bot's adapter
import adapter from "./adapter";

// This bot's main dialog.
import path from "path";
import send from "send";
import app from "./app/app";

// Create express application.
const expressApp = express();
expressApp.use(express.json());

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nAgent started, ${expressApp.name} listening to`, server.address());
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req, res) => {
  // Route received a request to adapter for processing
  await adapter.process(req, res as any, async (context) => {
    // Dispatch to application for routing
    await app.run(context);
  });
});

expressApp.get(["/auth-start.html", "/auth-end.html"], async (req, res) => {
  send(
    req,
    path.join(
      __dirname,
      "public",
      req.url.includes("auth-start.html") ? "auth-start.html" : "auth-end.html"
    )
  ).pipe(res);
});
