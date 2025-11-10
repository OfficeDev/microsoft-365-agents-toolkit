# TeamsFx SDK for TypeScript/JavaScript

## Notification Bot

TeamsFx SDK provides `ConversationBot.notification` to send proactive notification message to Bot installations, such as personal installations, channels and groupChats.

The main thing to note is listen to `membersAdded` event, store `ConversationReference` somewhere, then you can use it to send proactive notification later.

With @microsoft/teamsfx SDK, you can simply create a ConversationBot:

```ts
export const notificationApp = new ConversationBot({
  // Enable notification
  notification: {
    enabled: true,
  },
});

const pagedData = await notificationApp.notification.getPagedInstallations(
      pageSize,
      continuationToken
);
```

Without TeamsFx SDK, you can put key classes `middleware.ts` and `notification.ts` into your Teams app source code, use `LocalConversationReferenceStore` or implement your own persistent storage.

```ts
export const notificationApp = new NotificationBot(adapter, localStorage, authConfig.clientId);

const pagedData = await notificationApp.getPagedInstallations(pageSize, continuationToken);
```

You can refer to [notification-express template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vsc/ts/notification-express/src/notification).

## Command Bot

With TeamsFx SDK, you can define a ConversationBot with commands.
```ts
import { CommandMessage, TeamsFxBotCommandHandler, TriggerPatterns, BotBuilderCloudAdapter } from "@microsoft/teamsfx";
import ConversationBot = BotBuilderCloudAdapter.ConversationBot;

export const commandApp = new ConversationBot({
  adapterConfig: config,
  command: {
    enabled: true,
    commands: [new HelloWorldCommandHandler()],
  },
});

export class HelloWorldCommandHandler implements TeamsFxBotCommandHandler {
  triggerPatterns: TriggerPatterns = "helloWorld";

  async handleCommandReceived(
    context: TurnContext,
    message: CommandMessage
  ): Promise<string | Partial<Activity> | void> {
    // Your logic here
  }
}
```

### Option 1: Use Agents SDK

[Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) provides a simple set of functions over the Microsoft Bot Framework to implement this scenario.

```ts
import { TurnContext, MemoryStorage, Selector } from "@microsoft/agents-hosting";
import { Activity } from "@microsoft/agents-activity";
import { TeamsApplication } from "@microsoft/agents-hosting-teams";

const storage = new MemoryStorage();
export const app = new TeamsApplication<ApplicationTurnState>({
  storage,
});

export class HelloWorldCommandHandler {
  triggerPatterns: string | RegExp | Selector | (string | RegExp | Selector)[] = "helloWorld";

  async handleCommandReceived(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<string | Partial<Activity> | void> {
    // Your logic here
  }
}

const helloworldCommandHandler = new HelloWorldCommandHandler();
app.message(
  helloworldCommandHandler.triggerPatterns,
  async (context: TurnContext, state: ApplicationTurnState) => {
    const reply = await helloworldCommandHandler.handleCommandReceived(context, state);

    if (reply) {
      await context.sendActivity(reply as Activity);
    }
  }
);
```

### Option 2: Use Microsoft Teams SDK

[Microsoft Teams SDK](https://www.npmjs.com/package/@microsoft/teams-ai) provides a simple set of functions over the Microsoft Bot Framework to implement this scenario.


```ts
import { Activity, CardFactory, MessageFactory, TurnContext, MemoryStorage } from "botbuilder";
import { Application, Selector } from "@microsoft/teams-ai";

const storage = new MemoryStorage();
export const app = new Application<ApplicationTurnState>({
  storage,
});

export class HelloWorldCommandHandler {
  triggerPatterns: string | RegExp | Selector | (string | RegExp | Selector)[] = "helloWorld";

  async handleCommandReceived(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<string | Partial<Activity> | void> {
    // Your logic here
  }
}

const helloworldCommandHandler = new HelloWorldCommandHandler();
app.message(
  helloworldCommandHandler.triggerPatterns,
  async (context: TurnContext, state: ApplicationTurnState) => {
    const reply = await helloworldCommandHandler.handleCommandReceived(context, state);

    if (reply) {
      await context.sendActivity(reply);
    }
  }
);
```


## Bot SSO and Message Extension SSO

### Option 1: Move TeamsBotSsoPrompt.ts into source code

TeamsFx SDK provides a `TeamsBotSsoPrompt` class to simplify the authentication process when you develop bot application. You can move `TeamsBotSsoPrompt.ts` to your Teams app source code.

You can refer to [bot-sso sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/tree/dev/bot-sso).

### Option 2: Use Microsoft Teams SDK

[Microsoft Teams SDK](https://www.npmjs.com/package/@microsoft/teams-ai) also integrates with `TeamsBotSsoPrompt`. You can add authentication configurations to Application.

```ts
const app = new ApplicationBuilder()
    .withStorage(storage)
    .withAuthentication(adapter, {
        settings: {
            graph: {
                scopes: ['User.Read'],
                msalConfig: {
                    auth: {
                        clientId: config.clientId,
                        clientSecret: config.clientSecret,
                        authority: `${config.authorityHost}/${config.tenantId}`
                    }
                },
                signInLink: `https://${config.botDomain}/auth-start.html`,
                endOnInvalidMessage: true
            }
        },
        autoSignIn: true
    })
    .build();

app.message("photo", async (context: TurnContext, state: TurnState) => {
    const token = state.temp.authTokens['graph'];
    if (!token) {
        await context.sendActivity('No auth token found in state. Authentication failed.');
        return;
    }

    const response = await handlePhotoCommand(context, token);
    await context.sendActivity(response);
});
```

You can refer to [command-bot-with-sso sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/tree/dev/command-bot-with-sso) and [query-org-user-with-message-extension-sso sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/blob/dev/query-org-user-with-message-extension-sso).


## Tab SSO

### Option 1: Move TeamsUserCredential.ts into source code

`TeamsUserCredential` represents Teams current user's identity. Using this credential will request user consent at the first time. It leverages the Teams SSO and On-Behalf-Of flow to do token exchange. SDK uses this credential when developer choose "User" identity in browser environment.You can copy this `TeamsUserCredential.ts` into your source code.

### Option 2: Use @microsoft/teams-js SDK

With @microsoft/teams-js SDK:
```ts
import { app, authentication } from "@microsoft/teams-js";
await app.initialize();

const scopes = ["User.Read"];
const params = {
    url: `${
        config.initiateLoginEndpoint ? config.initiateLoginEndpoint : ""
    }?clientId=${config.clientId ? config.clientId : ""}&scope=${encodeURI(
        scopes.join(" ")
    )}`,
    width: 600,
    height: 535,
} as authentication.AuthenticatePopUpParameters;

await authentication.authenticate(params);
const ssoToken = await authentication.getAuthToken({
   resources: scopes
});

```
You can refer to [hello-world-tab-with-backend sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/tree/dev/hello-world-tab-with-backend).

### NAA

We recommend [Nested App Auth](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/authentication/nested-authentication) to implement SSO. 

```ts
import { app } from "@microsoft/teams-js";
import { createNestablePublicClientApplication } from "@azure/msal-browser";

await app.initialize();

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        supportsNestedAppAuth: true,
    },
};
const msalClient = await createNestablePublicClientApplication(msalConfig);
const result = await msalClient.loginPopup({
    scopes: ["User.Read"],
});
const account = result.account;
msalClient.setActiveAccount(account);

const result = await msalClient.acquireTokenSilent({
    scopes: ["User.Read"],
    account: account,
});

```

You can refer to [sso-tab-naa template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vsc/ts/sso-tab-naa).


## API client

You can just create your own client with `axios` library. E.g.

With TeamsFx SDK:
```ts
import { createApiClient, BearerTokenAuthProvider } from "@microsoft/teamsfx"
const apiClient = createApiClient(
  apiBaseUrl,
  new BearerTokenAuthProvider(
    async () => (await credential.getToken(""))!.token
  )
);

```

Without TeamsFx SDK:
```ts
import axios, { AxiosInstance } from "axios";
const apiClient = axios.create({ baseURL: apiBaseUrl });
  apiClient.interceptors.request.use(async (config) => {
      config.headers["Authorization"] = `Bearer ${ssoToken}`;
      return config;
    });

```