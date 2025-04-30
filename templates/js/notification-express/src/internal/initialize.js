const { AgentBuilderCloudAdapter } = require("@microsoft/teamsfx");
const ConversationBot = AgentBuilderCloudAdapter.ConversationBot;

// Create bot.
const notificationApp = new ConversationBot({
  // Enable notification
  notification: {
    enabled: true,
  },
});

module.exports = {
  notificationApp,
};
