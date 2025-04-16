const { BotBuilderCloudAdapter } = require("@microsoft/teamsfx");
const ConversationBot = BotBuilderCloudAdapter.ConversationBot;

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
