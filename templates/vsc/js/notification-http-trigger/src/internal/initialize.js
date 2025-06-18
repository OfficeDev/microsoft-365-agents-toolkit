const { NotificationBot } = require("../notification/notification");
const { loadAuthConfigFromEnv, CloudAdapter } = require("@microsoft/agents-hosting");
const { LocalConversationReferenceStore } = require("../notification/storage");
const path = require("path");

const authConfig = loadAuthConfigFromEnv();
// Create adapter
const adapter = new CloudAdapter(authConfig);

const localStorage = new LocalConversationReferenceStore(
  path.resolve(process.env.RUNNING_ON_AZURE === "1" ? process.env.TEMP ?? "./" : "./")
);

const notificationApp = new NotificationBot(adapter, localStorage, authConfig.clientId);

module.exports = {
  adapter,
  notificationApp,
};
