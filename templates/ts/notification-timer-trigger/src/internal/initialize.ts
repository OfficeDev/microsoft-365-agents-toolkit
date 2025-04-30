import { AgentBuilderCloudAdapter } from "@microsoft/teamsfx";
import ConversationBot = AgentBuilderCloudAdapter.ConversationBot;

// Create bot.
export const notificationApp = new ConversationBot({
  // Enable notification
  notification: {
    enabled: true,
  },
});
