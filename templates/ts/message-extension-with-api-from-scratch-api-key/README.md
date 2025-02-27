# Overview of Custom Search Results app template

## Build a message extension from a new API with Azure Function and API key authentication

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
4. To trigger the Message Extension, you can click the `+` under compose message area to find your message extension.
   > Note: Please make sure to switch to New Teams when Teams web client has launched

### How to add your own API Key

1. Open terminal and run command `npm install` to install all dependency packages

   ```
   > npm install
   ```

2. After `npm install` completed, run command `npm run keygen`
   ```
   > npm run keygen
   ```
3. The above command will output something like "Generated a new API Key: xxx..."
4. Fill in API Key into `env/.env.*.user`
   ```
   SECRET_API_KEY=<your-api-key>
   ```

## What's included in the template

| Folder       | Contents                                                                                                    |
| ------------ | ----------------------------------------------------------------------------------------------------------- |
| `.vscode`    | VSCode files for debugging                                                                                  |
| `appPackage` | Templates for the Teams application manifest, the API specification and response template for API responses |
| `env`        | Environment files                                                                                           |
| `infra`      | Templates for provisioning Azure resources                                                                  |
| `src`        | The source code for the repair API                                                                          |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                         | Contents                                                            |
| -------------------------------------------- | ------------------------------------------------------------------- |
| `src/functions/repair.ts`                    | The main file of a function in Azure Functions.                     |
| `src/repairsData.json`                       | The data source for the repair API.                                 |
| `src/keyGen.ts`                              | Designed to generate a API key used for authorization.              |
| `appPackage/apiSpecificationFile/repair.yml` | A file that describes the structure and behavior of the repair API. |
| `appPackage/responseTemplates/repair.json`   | A template file for rendering API response.                         |

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Teams Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                     |

## Addition information and references

- [Extend Teams platform with APIs](https://aka.ms/teamsfx-api-plugin)
