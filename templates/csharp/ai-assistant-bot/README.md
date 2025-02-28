# Overview of the AI Assistant Bot template

This app template is built on top of [Teams AI library](https://aka.ms/teams-ai-library) and [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview).
It showcases how to build an intelligent chat bot in Teams capable of helping users accomplish a specific task using natural language right in the Teams conversations, such as solving a math problem.

## Quick Start

**Prerequisites**
> To run the AI Assistant Bot template in your local dev machine, you will need:
>
> - An account with [OpenAI](https://platform.openai.com).
> 
> **Note**
>
> The `AssistantsPlanner` in `Microsoft.Teams.AI` is currently experimental.

### Create your own OpenAI Assistant

Before running or debugging your bot, please follow these steps to setup your own [OpenAI Assistant](https://platform.openai.com/docs/assistants/overview).

**If you haven't setup any Assistant yet**

> This app template provides script `Create-Assistant.ps1` to help create assistant. You can change the instructions and settings in the script to customize the assistant.
> 
> After creation, you can change and manage your assistants on [OpenAI](https://platform.openai.com/assistants).

1. Open PowerShell, change the current working directory to this project root and run command `. ./Create-Assistant.ps1 -OPENAI_API_KEY <your-openai-api-key>`
   ```
   > . ./Create-Assistant.ps1 -OPENAI_API_KEY xxxxxx
   ```
2. The above command will output the properties of the new created assistant, including the ID like "id: asst_xxx..."

### Debug bot app in Teams Web Client

1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.local.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   SECRET_OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```
2. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
1. Right-click the 'TeamsApp' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
2. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
3. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
4. In the opened web browser, select Add button to install the app in Teams
7. In the chat bar, type and send anything to your bot to trigger a response

> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

### Debug bot app in Teams App Test Tool
1. Fill in both OpenAI API Key and the created Assistant ID into `appsettings.TestTool.json`
   ```
   "OpenAI": {
      "ApiKey": "<your-openai-api-key>",
      "AssistantId": "<your-openai-assistant-id>"
   }
   ```
2. Select `Teams App Test Tool (browser)` in debug dropdown menu
3. Press F5, or select the Debug > Start Debugging menu in Visual Studio
4. In Teams App Test Tool from the launched browser, type and send anything to your bot to trigger a response 

## Extend the AI Assistant Bot template with more AI capabilities

You can follow [Get started with Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/how-conversation-ai-get-started) to extend the AI Assistant Bot template with more AI capabilities.

## Additional information and references
- [Teams AI library](https://aka.ms/teams-ai-library)
- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Learn more

New to Teams app development or Teams Toolkit? Learn more about 
Teams app manifests, deploying to the cloud, and more in the documentation 
at https://aka.ms/teams-toolkit-vs-docs.

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, you can create an issue directly in our GitHub repository: 
https://github.com/OfficeDev/TeamsFx/issues.
