# Overview of the AI Agent template

This app template is built on top of [Teams AI library](https://aka.ms/teams-ai-library) and [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview).
It showcases how to build an AI agent in Teams capable of helping users accomplish specific tasks using natural language right in the Teams conversations, such as solving a math problem, call functions to get city weather, etc.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Python](https://www.python.org/), version 3.8 or higher
> - [Python extension](https://code.visualstudio.com/docs/languages/python), version v2024.0.1 or higher
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
{{#useAzureOpenAI}}
> - An account with [Azure OpenAI](https://aka.ms/oai/access).
{{/useAzureOpenAI}}
{{#useOpenAI}}
> - An account with [OpenAI](https://platform.openai.com/).
{{/useOpenAI}}
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts).

{{#useAzureOpenAI}}
> Please make sure you are using model version 0613 or newer (0613, 1106, 0125) or gpt-4 turbo or gpt-35 turbo. Lower versions do NOT support assistants.
{{/useAzureOpenAI}}

### Configurations
1. Open the command box and enter `Python: Create Environment` to create and activate your desired virtual environment. Remember to select `src/requirements.txt` as dependencies to install when creating the virtual environment.
{{#useAzureOpenAI}}
1. In file *env/.env.local.user*, fill in your Azure OpenAI key `SECRET_AZURE_OPENAI_API_KEY`, deployment name `AZURE_OPENAI_MODEL_DEPLOYMENT_NAME` and endpoint `AZURE_OPENAI_ENDPOINT`.
{{/useAzureOpenAI}}
{{#useOpenAI}}
1. In file *env/.env.local.user*, fill in your OpenAI key `SECRET_OPENAI_API_KEY`. 
1. In this template, default model name is `gpt-3.5-turbo`. If you want to use a different model from OpenAI, fill in your model name in [src/config.py](./src/config.py).
{{/useOpenAI}}

### Create your own OpenAI Assistant

{{#useOpenAI}}
Before running or debugging your bot, please follow these steps to setup your own [OpenAI Assistant](https://platform.openai.com/docs/assistants/overview).
{{/useOpenAI}}
{{#useAzureOpenAI}}
Before running or debugging your bot, please follow these steps to setup your own [Azure OpenAI Assistant](https://learn.microsoft.com/en-us/azure/ai-services/openai/how-to/assistant).
{{/useAzureOpenAI}}

**If you haven't setup any Assistant yet**

> This app template provides script `src/utils/creator.py` to help create assistant. You can change the instructions and settings in the script to customize the assistant.
> 
{{#useOpenAI}}
> After creation, you can change and manage your assistants on [OpenAI](https://platform.openai.com/assistants).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> After creation, you can change and manage your assistants on [Azure OpenAI Studio](https://oai.azure.com/).
{{/useAzureOpenAI}}

{{#useOpenAI}}
1. Run command `python src/utils/creator.py`. Remember to input your **OpenAI key** in command parameter.
   ```
   python src/utils/creator.py --api-key <your-openai-api-key>
   ```
1. The above command will output something like "*Created a new assistant with an ID of: **asst_xxx...***".
1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.local.user`.
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```

**If you already have an Assistant created**

1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.local.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Run command `python src/utils/creator.py`. Remember to input your **Azure OpenAI key** in command parameter.
   ```
   python src/utils/creator.py --api-key <your-azure-openai-api-key>
   ```
1. The above command will output something like "*Created a new assistant with an ID of: **asst_xxx...***".
1. Fill in both Azure OpenAI API Key, endpoint, deployment name and the created Assistant ID into `env/.env.local.user`.
   ```
   SECRET_AZURE_OPENAI_API_KEY=<your-azure-openai-api-key>
   AZURE_OPENAI_ENDPOINT=<your-azure-openai-endpoint>
   AZURE_OPENAI_MODEL_DEPLOYMENT_NAME=<your-azure-openai-model-delopyment-name>
   AZURE_OPENAI_ASSISTANT_ID=<your-azure-openai-assistant-id>
   ```

**If you already have an Assistant created**

1. Fill in both Azure OpenAI API Key, endpoint, deployment name and the created Assistant ID into `env/.env.local.user`.
   ```
   SECRET_AZURE_OPENAI_API_KEY=<your-azure-openai-api-key>
   AZURE_OPENAI_ENDPOINT=<your-azure-openai-endpoint>
   AZURE_OPENAI_MODEL_DEPLOYMENT_NAME=<your-azure-openai-model-delopyment-name>
   AZURE_OPENAI_ASSISTANT_ID=<your-azure-openai-assistant-id>
   ```
{{/useAzureOpenAI}}

### Conversation with bot
1. Select the Teams Toolkit icon on the left in the VS Code toolbar.
1. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
1. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
1. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
1. You will receive a welcome message from the bot, or send any message to get a response.

**Congratulations**! You are running an application that can now interact with users in Teams:

> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

![AI Agent in Teams](https://github.com/OfficeDev/TeamsFx/assets/37978464/fd1cf673-e7d8-4826-9cac-e9481a74ee1e)

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
|`src/app.py`| Hosts an aiohttp api server and exports an app module.|
|`src/bot.py`| Handles business logics for the AI Agent.|
|`src/config.py`| Defines the environment variables.|

The following file is a script that helps you to prepare an OpenAI assistant.

| File                                 | Contents                                           |
| - | - |
|`src/utils/creator.py`| Create an OpenAI assistant with defined functions and prompts.|

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`teamsapp.yml`|This is the main Teams Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`teamsapp.local.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging.|
|`teamsapp.testtool.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging in Teams App Test Tool.|

## Extend the template

You can follow [Build an AI Agent in Teams](https://aka.ms/teamsfx-ai-agent) to extend the AI Agent template with more AI capabilities, like:
- [Add functions](https://aka.ms/teamsfx-ai-agent#add-functions-build-new)

## Additional information and references

- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Known issue
- If you use `Debug in Test Tool` to local debug, you might get an error `InternalServiceError: connect ECONNREFUSED 127.0.0.1:3978` in Test Tool console log or error message `Error: Cannot connect to your app,
please make sure your app is running or restart your app` in log panel of Test Tool web page. You can wait for Python launch console ready and then refresh the front end web page.
- When you use `Launch Remote in Teams` to remote debug after deployment, you might loose interaction with your bot. This is because the remote service needs to restart. Please wait for several minutes to retry it.