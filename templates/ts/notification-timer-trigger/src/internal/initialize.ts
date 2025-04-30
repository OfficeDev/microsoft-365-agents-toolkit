import { BotBuilderCloudAdapter } from "@microsoft/teamsfx";
import ConversationBot = BotBuilderCloudAdapter.ConversationBot;

// Create bot.
export const notificationApp = new ConversationBot({
  // Enable notification
  notification: {
    enabled: true,
  },
});
