# Overview of Bot Based Message Extension template

A Message Extension allows users to interact with your web service while composing messages in the Microsoft Teams client. Users can invoke your web service to assist message composition, from the message compose box, or from the search bar.

This app template has a search command, an action command and a link unfurling.

1. The search command allows users to search an external system and share results through the compose message area of the Microsoft Teams client.
2. The action command allows you to present your users with a modal pop-up called a task module in Teams. The task module collects or displays information, processes the interaction, and sends the information back to Teams.
3. With link unfurling, an app can unfurl a link into an adaptive card when URLs with a particular domain are pasted into the compose message area in Microsoft Teams or email body in Outlook.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
{{^enableMETestToolByDefault}}
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
{{/enableMETestToolByDefault}}
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
{{#enableMETestToolByDefault}}
2. Press F5 to start debugging which launches your app in Microsoft 365 Agents Playground using a web browser.
3. To trigger the Message Extension to invoke commands:
   1. To trigger search commands, click the `+` in compose message area and select `Search command`.
   2. To trigger action commands, click the `+` in compose message area or `...` above a message and select `Action command`.
   3. To trigger link unfurling, click the `+` in compose message area and select `Link unfurling`.

![Search app demo](https://github.com/OfficeDev/TeamsFx/assets/9698542/5275e5bc-492f-4365-b602-5803938a9780)

![action-ME](https://github.com/OfficeDev/TeamsFx/assets/9698542/c0afbd89-7fbb-4e73-98a2-f018be4ca88c)
{{/enableMETestToolByDefault}}
{{^enableMETestToolByDefault}}
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. To trigger the Message Extension to invoke commands, you can:
   1. In Teams: `@mention` Your message extension from the `search box area`, `@mention` your message extension from the `compose message area` or click the `...` under compose message area to find your message extension.
   2. In Outlook: click the `More apps` icon under compose email area to find your message extension. Only search command and link unfurling works in Outlook.
6. Paste a link ending with `.botframework.com` into compose message area in Teams or email body in Outlook. You should see an adaptive card unfurled.

![Search app demo](https://user-images.githubusercontent.com/11220663/167868361-40ffaaa3-0300-4313-ae22-0f0bab49c329.png)

![action-ME](https://github.com/OfficeDev/TeamsFx/assets/25220706/378ea4d7-9332-4aec-9f85-59891d086b80)
{{/enableMETestToolByDefault}}

## What's included in the template

| Folder       | Contents                                     |
| ------------ | -------------------------------------------- |
| `.vscode`    | VSCode files for debugging                   |
| `appPackage` | Templates for the application manifest |
| `env`        | Environment files                            |
| `infra`      | Templates for provisioning Azure resources   |

The following files can be customized and demonstrate an example implementation to get you started.

| File                | Contents                                                                                                |
| ------------------- | ------------------------------------------------------------------------------------------------------- |
| `./src/teamsBot.ts` | Handles the business logic for this app template to query npm registry and return result list to Teams. |
| `./src/index.ts`    | `index.ts` is used to setup and configure the Message Extension.                                        |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `m365agents.yml`       | This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `teamsapp.local.yml` | This overrides `m365agents.yml` with actions that enable local execution and debugging.                                                     |
| `teamsapp.testtool.yml`| This overrides `m365agents.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.                            |

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Microsoft 365 Agents Toolkit CLI](https://aka.ms/teams-toolkit-cli/debug)
- [Preview the app on mobile clients](https://aka.ms/teamsfx-mobile)
