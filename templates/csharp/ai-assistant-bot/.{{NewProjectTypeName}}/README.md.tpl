# Overview of the AI Assistant Bot template

This app template is built using [Teams AI library](https://aka.ms/teams-ai-library) and [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview) to help you build an intelligent chat bot in Teams that can help users accomplish a task using natural language within Teams conversations.

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

Before running or debugging your bot, follow these steps to setup your own [OpenAI Assistant](https://platform.openai.com/docs/assistants/overview).

**If you haven't setup any Assistant yet**

> This app template includes a script `Create-Assistant.ps1` to help you create an assistant. You can customize the assistant by changing the instructions and settings in the script.
> 
> After creation, you can change and manage your assistants on [OpenAI](https://platform.openai.com/assistants).

1. Open PowerShell, change the current working directory to this project root and run command `. ./Create-Assistant.ps1 -OPENAI_API_KEY <your-openai-api-key>`
   ```
   > . ./Create-Assistant.ps1 -OPENAI_API_KEY xxxxxx
   ```
2. The above command will display the properties of the newly created assistant, including the ID like "id: asst_xxx..."

### Debug bot app in Teams Web Client

1. Fill in both OpenAI API Key and the created Assistant ID into `env/.env.local.user`
   ```
   SECRET_OPENAI_API_KEY=<your-openai-api-key>
   SECRET_OPENAI_ASSISTANT_ID=<your-openai-assistant-id>
   ```
2. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/create-devtunnel-button.png)
1. Right-click the `{{NewProjectTypeName}}` project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
2. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
3. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
4. In the opened web browser, select Add button to install the app in Teams
7. In the message input field, type and send anything to your bot to get a response

## Run the app on other platforms

The Teams app can run in other platforms like Outlook and Microsoft 365 app. See https://aka.ms/vs-ttk-debug-multi-profiles for more details.

## Extend the AI Assistant Bot template with more AI capabilities

Follow [Get started with Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/how-conversation-ai-get-started) to enhance the AI Assistant Bot template with advanced features.

## Additional information and references
- [Teams AI library](https://aka.ms/teams-ai-library)
- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Get more info

New to Teams app development or Teams Toolkit? Explore Teams app manifests, cloud deployment, and much more in the https://aka.ms/teams-toolkit-vs-docs.

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues.
