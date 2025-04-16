const { AgentApplication } = require("@microsoft/agents-hosting");

const teamsBot = new AgentApplication();

teamsBot.conversationUpdate("membersAdded", async (context) => {
  await context.sendActivity(
    "Welcome to the Notification Bot! I am designed to send you updates and alerts using Adaptive Cards triggered by HTTP post requests. Please note that I am a notification-only bot and you can't interact with me. Follow the README in the project and stay tuned for notifications!"
  );
});

module.exports.teamsBot = teamsBot;
