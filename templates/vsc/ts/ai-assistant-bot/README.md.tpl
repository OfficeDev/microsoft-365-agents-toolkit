# Overview of the AI Assistant Bot template

This app template is built on top of [Teams AI library](https://aka.ms/teams-ai-library) and [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview).
It showcases how to build an intelligent chat bot in Teams capable of helping users accomplish a specific task using natural language right in the Teams conversations, such as solving a math problem.

- [Overview of the AI Assistant Bot template](#overview-of-the-ai-assistant-bot-template)
  - [Get started with the AI Assistant Bot template](#get-started-with-the-ai-assistant-bot-template)
  - [What's included in the template](#whats-included-in-the-template)
  - [Extend the AI Assistant Bot template with more AI capabilities](#extend-the-ai-assistant-bot-template-with-more-ai-capabilities)
  - [Additional information and references](#additional-information-and-references)

## Get started with the AI Assistant Bot template

> **Prerequisites**
>
> To run the AI Assistant Bot template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
{{^enableTestToolByDefault}}
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
{{/enableTestToolByDefault}}
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
> - An account with [OpenAI](https://platform.openai.com/).
>
> **Note**
>
> The `AssistantsPlanner` in Teams AI Library is currently in preview version.
> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).


### Create your own OpenAI Assistant

Before running or debugging your bot, please follow these steps to setup your own [OpenAI Assistant](https://platform.openai.com/docs/assistants/overview).

**If you haven't setup any Assistant yet**

> This app template provides script `src/creator.ts` to help create assistant. You can change the instructions and settings in the script to customize the assistant.
> 
> After creation, you can change and manage your assistants on [OpenAI](https://platform.openai.com/assistants).

1. Open terminal and run command `npm install` to install all dependency packages

   ```
   > npm install
   ```
1. After `npm install` completed, run command `npm run assistant:create -- <your-openai-api-key>`
   ```
   > npm run assistant:create -- xxxxxx
   ```
1. The above command will output something like "*Created a new assistant with an ID of: **asst_xxx...***"
1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.*.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   SECRET_OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```

**If you already have an Assistant created**

1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.*.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   SECRET_OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```

### Run Teams Bot locally

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
{{#enableTestToolByDefault}}
1. Press F5 to start debugging which launches your app in Microsoft 365 Agents Playground using a web browser. Select `Debug in Microsoft 365 Agents Playground`.
1. You will receive a welcome message from the bot, or send any message to get a response.

**Congratulations**! You are running an application that can now interact with users in Microsoft 365 Agents Playground:

![ai assistant bot in Microsoft 365 Agents Playground](https://github.com/user-attachments/assets/2618217b-1adc-48cc-ae6e-d411369e4deb)
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
1. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
1. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
1. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
1. You will receive a welcome message from the bot, or send any message to get a response.

**Congratulations**! You are running an application that can now interact with users in Teams:

![ai assistant bot in Teams](https://github.com/OfficeDev/TeamsFx/assets/7642967/21480d90-307d-4074-84e0-c68a20e38134)
{{/enableTestToolByDefault}}

## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `src`        | The source code for the application                 |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/index.ts`| Sets up and configures the AI Assistant Bot.|
|`src/app.ts`| Handles business logics for the AI Assistant Bot.|
|`src/config.ts`| Defines the environment variables.|
|`src/creator.ts`| One-time tool to create OpenAI Assistant.|

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`m365agents.yml`|This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`m365agents.local.yml`|This overrides `m365agents.yml` with actions that enable local execution and debugging.|
|`m365agents.playground.yml`|This overrides `m365agents.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.|

## Extend the AI Assistant Bot template with more AI capabilities

You can follow [Get started with Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/how-conversation-ai-get-started) to extend the AI Assistant Bot template with more AI capabilities.

## Additional information and references
- [Teams AI library](https://aka.ms/teams-ai-library)
- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
- [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview)
