# Overview of Custom Search Results app template

## Build a message extension from OpenAPI description document

This app template allows Teams to interact directly with third-party data, apps, and services, enhancing its capabilities and broadening its range of capabilities. It allows Teams to:

- Retrieve real-time information, for example, latest news coverage on a product launch.
- Retrieve knowledge-based information, for example, my team’s design files in Figma.

## Get started with the template

> **Prerequisites**
>
> To run this app template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts).
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Create Teams app by clicking `Provision` in "Lifecycle" section.
4. Select `Preview in Teams (Edge)` or `Preview in Teams (Chrome)` from the launch configuration dropdown.
5. To trigger the Message Extension, you can click the `+` under compose message area to find your message extension.
   > Note: Please make sure to switch to New Teams when Teams web client has launched
{{#ApiKey}}
> [!NOTE]
> Microsoft 365 Agents Toolkit will ask you for your API key during provision. The API key will be securely stored with [Teams Developer Portal](https://dev.teams.microsoft.com/home) and used by Teams client to access your API in runtime. Microsoft 365 Agents Toolkit will not store your API key.
{{/ApiKey}}

{{#OAuth}}
> [!NOTE]
> If your identity server needs Proof of Key Code Exchange (PKCE) for token exchange, uncomment the `isPKCEEnabled` property in the` oauth/register` section of the `teamsapp.yml` file shown as below:
```yaml
  - uses: oauth/register
    with:
      name: {{ApiSpecAuthName}}
      flow: authorizationCode
      # Teams app ID
      appId: ${{TEAMS_APP_ID}}
      # Path to OpenAPI description document
      apiSpecPath: {{{ApiSpecPath}}}
      # Uncomment below property to use proof key for code exchange (PKCE)
      isPKCEEnabled: true
    writeToEnvironmentFile:
      configurationId: {{ApiSpecAuthRegistrationIdEnvName}}
```
> Microsoft 365 Agents Toolkit will ask you for your Client ID and Client Secret for Oauth2 during provision. These information will be securely stored with [Teams Developer Portal](https://dev.teams.microsoft.com/home) and used by Teams client to access your API in runtime. Microsoft 365 Agents Toolkit will not store your Client ID and Client Secret.
{{/OAuth}}

## What's included in the template

| Folder       | Contents                                     |
| ------------ | -------------------------------------------- |
| `.vscode`    | VSCode files for debugging                   |
| `appPackage` | Templates for the Teams application manifest, the API specification and response templates for API responses |
| `env`        | Environment files                            |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |

## Addition information and references

- [Extend Teams platform with APIs](https://aka.ms/teamsfx-api-plugin)
