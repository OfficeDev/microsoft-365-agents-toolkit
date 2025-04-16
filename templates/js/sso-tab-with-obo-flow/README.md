# Overview of the React with Fluent UI template

This app showcases how to craft a visually appealing web page that can be embedded in Microsoft Teams, Outlook and the Microsoft 365 app with React and Fluent UI. The app also enhances the end-user experiences with built-in single sign-on and data from Microsoft Graph.

This app has adopted [On-Behalf-Of flow](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to implement SSO, and uses Azure Functions as middle-tier service, and make authenticated requests to call Graph from Azure Functions.

## Get started with the React with Fluent UI template

> **Prerequisites**
>
> To run the command bot template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
>   Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.

**Congratulations**! You are running an application that can now show a beautiful web page in Teams, Outlook and the Microsoft 365 app.

![Personal tab demo](https://github.com/OfficeDev/TeamsFx/assets/63089166/9599b53c-8f89-493f-9f7e-9edae1f9be54)

## What's included in the template

| Folder       | Contents                                                                                                               |
| ------------ | ---------------------------------------------------------------------------------------------------------------------- |
| `.vscode`    | VSCode files for debugging                                                                                             |
| `appPackage` | Templates for the application manifest                                                                           |
| `env`        | Environment files                                                                                                      |
| `infra`      | Templates for provisioning Azure resources                                                                             |
| `src`        | The source code for the frontend of the Tab application. Implemented with Fluent UI Framework.                         |
| `api`        | The source code for the backend of the Tab application. Implemented single-sign-on with OBO flow using Azure Functions. |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                 | Contents                                                                                                                                                                                                                                                |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `m365agents.yml`       | This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions.                                                                                                               |
| `m365agents.local.yml` | This overrides `m365agents.yml` with actions that enable local execution and debugging.                                                                                                                                                                   |
| `aad.manifest.json`  | This file defines the configuration of Microsoft Entra app. This template will only provision [single tenant](https://learn.microsoft.com/azure/active-directory/develop/single-and-multi-tenant-apps#who-can-sign-in-to-your-app) Microsoft Entra app. |

## Extend the React with Fluent UI template

Following documentation will help you to extend the React with Fluent UI template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Enable the app for multi-tenant](https://github.com/OfficeDev/TeamsFx/wiki/Multi-tenancy-Support-for-Azure-AD-app)
- [Preview the app on mobile clients](https://aka.ms/teamsfx-mobile)
