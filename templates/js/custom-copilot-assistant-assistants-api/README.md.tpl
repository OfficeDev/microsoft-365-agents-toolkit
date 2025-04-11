# Overview of the AI Agent template

This app template is built on top of [Teams AI library](https://aka.ms/teams-ai-library) and [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview).
It showcases how to build an AI agent in Teams capable of helping users accomplish specific tasks using natural language right in the Teams conversations, such as solving a math problem, call functions to get city weather, etc.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
{{^enableTestToolByDefault}}
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
{{/enableTestToolByDefault}}
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
{{#useAzureOpenAI}}
> - An account with [Azure OpenAI](https://aka.ms/oai/access).
{{/useAzureOpenAI}}
{{#useOpenAI}}
> - An account with [OpenAI](https://platform.openai.com/).
{{/useOpenAI}}
>
> **Note**
>
> The `AssistantsPlanner` in Teams AI Library is currently in preview version.

### Create your own OpenAI Assistant

{{#useOpenAI}}
Before running or debugging your bot, please follow these steps to setup your own [OpenAI Assistant](https://platform.openai.com/docs/assistants/overview).
{{/useOpenAI}}
{{#useAzureOpenAI}}
Before running or debugging your bot, please follow these steps to setup your own [Azure OpenAI Assistant](https://learn.microsoft.com/en-us/azure/ai-services/openai/how-to/assistant).
{{/useAzureOpenAI}}

**If you haven't setup any Assistant yet**

> This app template provides script `src/creator.js` to help create assistant. You can change the instructions and settings in the script to customize the assistant.
> 
{{#useOpenAI}}
> After creation, you can change and manage your assistants on [OpenAI](https://platform.openai.com/assistants).
{{/useOpenAI}}

1. Open terminal and run command `npm install` to install all dependency packages
   ```
   > npm install
   ```
{{#useOpenAI}}
1. After `npm install` completed, run command `npm run assistant:create -- <your-openai-api-key>`
   ```
   > npm run assistant:create -- xxxxxx
   ```
1. The above command will output something like "*Created a new assistant with an ID of: **asst_xxx...***"
1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.*.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```

**If you already have an Assistant created**

1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.*.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. After `npm install` completed, fill in both Azure OpenAI API Endpoint and Azure OpenAI API Deployment name into `src/creator.js`
   ```
   const azureOpenAIEndpoint="<your-azure-openai-endpoint>";
   const azureOpenAIDeploymentName="<your-azure-openai-delopyment-name>";
   ```
1. Run command `npm run assistant:create -- <your-azure-openai-api-key>`
   ```
   > npm run assistant:create -- xxxxxx
   ```
1. The above command will output something like "*Created a new assistant with an ID of: **asst_xxx...***"
1. Fill in Azure OpenAI API Key, endpoint, and the created Assistant ID into `env/.env.*.user`
   ```
   SECRET_AZURE_OPENAI_API_KEY=<your-azure-openai-api-key>
   AZURE_OPENAI_ENDPOINT=<your-azure-openai-endpoint>
   AZURE_OPENAI_ASSISTANT_ID=<your-azure-openai-assistant-id>
   ```

**If you already have an Assistant created**

1. Fill in Azure OpenAI API Key, endpoint, and the created Assistant ID into `env/.env.*.user`
   ```
   SECRET_AZURE_OPENAI_API_KEY=<your-azure-openai-api-key>
   AZURE_OPENAI_ENDPOINT=<your-azure-openai-endpoint>
   AZURE_OPENAI_ASSISTANT_ID=<your-azure-openai-assistant-id>
   ```
{{/useAzureOpenAI}}

### Run Teams Bot locally

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
{{#enableTestToolByDefault}}
1. Press F5 to start debugging which launches your app in Microsoft 365 Agents Playground using a web browser. Select `Debug in Microsoft 365 Agents Playground`.
1. You can send any message to get a response from the bot.

**Congratulations**! You are running an application that can now interact with users in Microsoft 365 Agents Playground:

![AI Agent in Microsoft 365 Agents Playground](https://github.com/OfficeDev/TeamsFx/assets/37978464/e3b458f3-5e74-460d-9df2-bf77ed8d9c54)
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
1. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't yet.
1. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
1. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
1. You can send any message to get a response from the bot.

**Congratulations**! You are running an application that can now interact with users in Teams:

![AI Agent in Teams](https://github.com/OfficeDev/TeamsFx/assets/37978464/fd1cf673-e7d8-4826-9cac-e9481a74ee1e)
{{/enableTestToolByDefault}}

## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the Teams application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `src`        | The source code for the application                 |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/index.js`| Sets up the bot app server.|
|`src/adapter.js`| Sets up the bot adapter.|
|`src/config.js`| Defines the environment variables.|
|`src/creator.js`| One-time tool to create OpenAI Assistant.|
|`src/app/app.js`| Handles business logics for the AI Agent.|
|`src/app/messages.js`| Defines the message activity handlers.|
|`src/app/actions.js`| Defines the AI actions.|

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`teamsapp.yml`|This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`teamsapp.local.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging.|
|`teamsapp.testtool.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.|

## Extend the template

You can follow [Build an AI Agent in Teams](https://aka.ms/teamsfx-ai-agent) to extend the AI Agent template with more AI capabilities, like:
- [Customize assistant creation](https://aka.ms/teamsfx-ai-agent#customize-assistant-creation)
- [Add functions](https://aka.ms/teamsfx-ai-agent#add-functions-with-assistants-api)

## Additional information and references

- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
- [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview)
