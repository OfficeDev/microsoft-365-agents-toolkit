# Overview of Custom Search Results app template

## Build a message extension from a new API with Azure Functions

This app template allows Teams to interact directly with third-party data, apps, and services, enhancing its capabilities and broadening its range of capabilities. It allows Teams to:

- Retrieve real-time information, for example, latest news coverage on a product launch.
- Retrieve knowledge-based information, for example, my team’s design files in Figma.

## Get started with the template

> **Prerequisites**
>
> To run this app template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)` from the launch configuration dropdown.
4. When Teams launches in the browser, you can navigate to a chat message and [trigger your search commands from compose message area](https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=dotnet#search-commands).
   > Note: Please make sure to switch to New Teams when Teams web client has launched

## What's included in the template

| Folder       | Contents                                                                                                    |
| ------------ | ----------------------------------------------------------------------------------------------------------- |
| `.vscode`    | VSCode files for debugging                                                                                  |
| `appPackage` | Templates for the Teams application manifest, the API specification and response template for API responses |
| `env`        | Environment files                                                                                           |
| `infra`      | Templates for provisioning Azure resources                                                                  |
| `src`        | The source code for the repair API                                                                          |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                            | Contents                                                                                                                     |
| ----------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `src/functions/repair.ts`                       | The main file of a function in Azure Functions.                                                                              |
| `src/functions/middleware/tokenCacheWrapper.ts` | A wrapper class that handles caching of JWT signing keys to improve performance of token validation.                         |
| `src/functions/middleware/tokenValidator.ts`    | Core class for validating JWT tokens from Microsoft Entra, including checks for claims, scopes, roles and tenant validation. |
| `src/functions/middleware/authMiddleware.ts`    | Middleware function that handles authorization using JWT tokens, integrating with the token validator.                       |
| `src/functions/middleware/utils.ts`             | Utility functions for authentication, including retrieving JWKS URIs for different cloud environments.                       |
| `src/functions/middleware/config.ts`            | Configuration file that exports Microsoft Entra app settings from environment variables.                                     |
| `src/functions/middleware/authMiddleware.ts`    | A middleware function that implements Microsoft Entra authentication for Azure Functions.                                    |
| `src/repairsData.json`                          | The data source for the repair API.                                                                                          |
| `appPackage/apiSpecificationFile/repair.yml`    | A file that describes the structure and behavior of the repair API.                                                          |
| `appPackage/responseTemplates/repair.json`      | A generated Adaptive Card that used to render API response.                                                                  |

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                 | Contents                                                                                                                                                                                                                                                |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Teams Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions.                                                                                                               |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                                                                                                                                   |
| `aad.manifest.json`  | This file defines the configuration of Microsoft Entra app. This template will only provision [single tenant](https://learn.microsoft.com/azure/active-directory/develop/single-and-multi-tenant-apps#who-can-sign-in-to-your-app) Microsoft Entra app. |

## How Microsoft Entra works

![microsoft-entra-flow](https://github.com/OfficeDev/TeamsFx/assets/107838226/846e7a60-8cc1-4d8b-852e-2aec93b61fe9)

## Addition information and references

- [Extend Teams platform with APIs](https://aka.ms/teamsfx-api-plugin)
