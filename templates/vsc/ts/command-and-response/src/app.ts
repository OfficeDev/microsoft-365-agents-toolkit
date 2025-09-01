import { App } from "@microsoft/teams.apps";
import { HelloWorldCommandHandler } from "./helloworldCommandHandler";
import { GenericCommandHandler } from "./genericCommandHandler";
import { TokenCredentials } from "@microsoft/teams.api";
import { ManagedIdentityCredential } from "@azure/identity";
import config from "./internal/config";

const createTokenFactory = () => {
  return async (scope: string | string[], tenantId?: string): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;
export const app = new App({
  ...credentialOptions,
});

// Initialize command handlers
const helloWorldHandler = new HelloWorldCommandHandler();
const genericHandler = new GenericCommandHandler();

// Register message handler
app.on("message", async ({ activity, send }) => {
  const text = activity.text || "";

  // Check if helloWorld command
  if (helloWorldHandler.canHandle(text)) {
    const reply = await helloWorldHandler.handleCommandReceived(activity);
    if (reply) {
      await send(reply);
    }
    return;
  }

  // Handle all other messages with generic handler
  const reply = await genericHandler.handleCommandReceived(activity);
  if (reply) {
    await send(reply);
  }
});

export default app;
