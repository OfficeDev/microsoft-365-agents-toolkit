import { NotificationBot } from "../notification/notification";
import { AuthConfiguration, loadAuthConfigFromEnv, CloudAdapter } from "@microsoft/agents-hosting";
import { LocalConversationReferenceStore } from "../notification/storage";
import * as path from "path";

const authConfig: AuthConfiguration = loadAuthConfigFromEnv();
// Create adapter
export const adapter = new CloudAdapter(authConfig);

export const localStorage = new LocalConversationReferenceStore(
  path.resolve(process.env.RUNNING_ON_AZURE === "1" ? process.env.TEMP ?? "./" : "./")
);

export const notificationApp = new NotificationBot(adapter, localStorage, authConfig.clientId);
