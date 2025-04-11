# Overview of Collect Form Input and Process Data template

A Message Extension allows users to interact with your web service while composing messages in the Microsoft Teams client. Users can search or initiate actions in an external system from compose message area, the command box, or directly from a message.

This app template implements action command that allows you to present your users with a modal pop-up called a task module in Teams. The task module collects or displays information, processes the interaction, and sends the information back to Teams.

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
2. Press F5 to start debugging which launches your app in Teams App Test Tool using a web browser.
3. To trigger the action command, you can click the `+` under compose message area and select `Action Command`.

**Congratulations**! You are running an application that can share information in rich format by creating an Adaptive Card in Teams App Test Tool.

![action-ME](https://github.com/OfficeDev/TeamsFx/assets/9698542/c0afbd89-7fbb-4e73-98a2-f018be4ca88c)
{{/enableMETestToolByDefault}}
{{^enableMETestToolByDefault}}
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. To trigger the action command, you can click the `...` under compose message area to find your message extension.

**Congratulations**! You are running an application that can share information in rich format by creating an Adaptive Card in Teams.

![action-ME](https://github.com/OfficeDev/TeamsFx/assets/25220706/378ea4d7-9332-4aec-9f85-59891d086b80)
{{/enableMETestToolByDefault}}

## What's included in the template

| Folder        | Contents                                     |
| ------------- | -------------------------------------------- |
| `.vscode/`    | VSCode files for debugging                   |
| `appPackage/` | Templates for the Teams application manifest |
| `env/`        | Environment files                            |
| `infra/`      | Templates for provisioning Azure resources   |
| `src/`        | The source code for the action application   |

The following files can be customized and demonstrate an example implementation to get you started.

| File               | Contents                                                                                 |
| ------------------ | ---------------------------------------------------------------------------------------- |
| `src/actionApp.js` | Handles the business logic for this app template to collect form input and process data. |
| `src/index.js`     | `index.js` is used to setup and configure the Message Extension.                         |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                     |
| `teamsapp.testtool.yml`| This overrides `teamsapp.yml` with actions that enable local execution and debugging in Teams App Test Tool.                            |

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Microsoft 365 Agents Toolkit CLI](https://aka.ms/teams-toolkit-cli/debug)
- [Preview the app on mobile clients](https://aka.ms/teamsfx-mobile)

