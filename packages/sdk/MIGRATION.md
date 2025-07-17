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

Without TeamsFx SDK, you can put key classes `middleware.ts` and `notification.ts` into your Teams app source code, use `LocalConversationReferenceStore` or implement your own persistant storage.

```ts
export const notificationApp = new NotificationBot(adapter, localStorage, authConfig.clientId);

const pagedData = await notificationApp.getPagedInstallations(pageSize, continuationToken);
```

You can reference to [notification-express template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vsc/ts/notification-express/src/notification).

## Bot SSO and Message Extension SSO

### Move TeamsBotSsoPrompt.ts into source code

Teamsfx SDK provides a `TeamsBotSsoPrompt` class to simply the authentication process when you develop bot application. You can move `TeamsBotSsoPrompt.ts` to your Teams app source code. 

You can reference to [bot-sso sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/tree/dev/bot-sso).

### Use Teams AI Library

[Teams AI Library](https://www.npmjs.com/package/@microsoft/teams-ai) also integrates with `TeamsBotSsoPrompt`. You can add authentication configurations to Application.

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

You can reference [command-bot-with-sso sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/tree/dev/command-bot-with-sso) and [query-org-user-with-messaage-extension-sso sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/blob/dev/query-org-user-with-message-extension-sso).


## Tab SSO

`TeamsUserCredential` represents Teams current user's identity. Using this credential will request user consent at the first time. It leverages the Teams SSO and On-Behalf-Of flow to do token exchange. SDK uses this credential when developer choose "User" identity in browser environment.You can copy this `TeamsUserCredential.ts` into your source code, or directly use `@microsoft/teams-js` with `NAA(Nested App Auth)`.

### Use @microsoft/teams-js SDK

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
You can reference to [hello-world-tab-with-backend sample](https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples/tree/dev/hello-world-tab-with-backend).

### NAA

We recomment [Nested App Auth](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/authentication/nested-authentication) to implement SSO. 

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

You can reference to [sso-tab-naa template](https://github.com/OfficeDev/microsoft-365-agents-toolkit/tree/dev/templates/vsc/ts/sso-tab-naa).


## API client

You can just create your own client with `axios` library. E.g.

With Teamsfx SDK:
```ts
import { createApiClient, BearerTokenAuthProvider } from "@microsoft/teamsfx"
const apiClient = createApiClient(
  apiBaseUrl,
  new BearerTokenAuthProvider(
    async () => (await credential.getToken(""))!.token
  )
);

```

Without Teamsfx SDK:
```ts
import axios, { AxiosInstance } from "axios";
const apiClient = axios.create({ baseURL: apiBaseUrl });
  apiClient.interceptors.request.use(async (config) => {
      config.headers["Authorization"] = `Bearer ${ssoToken}`;
      return config;
    });

```