const { MemoryStorage } = require("botbuilder");
const { Application } = require("@microsoft/teams-ai");

// Define storage and application
const storage = new MemoryStorage();
const app = new Application({
  storage,
});

module.exports.app = app;
