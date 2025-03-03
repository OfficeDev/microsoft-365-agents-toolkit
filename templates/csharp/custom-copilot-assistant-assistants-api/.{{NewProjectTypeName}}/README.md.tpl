# Overview of the AI Agent template

This template showcases a bot app that responds to user questions like an AI assistant. This enables your users to talk with the AI assistant in Teams to find information.

The app template is built using the Teams AI library, which provides the capabilities to build AI-based Teams applications.

## Quick Start

**Prerequisites**
> To run the AI Chat Bot template in your local dev machine, you will need:
>
{{#useOpenAI}}
> - an account with [OpenAI](https://platform.openai.com).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - [Azure OpenAI](https://aka.ms/oai/access) resource
{{/useAzureOpenAI}}

## Create your Assistant 
**Before running or debugging your bot, follow these steps to setup your own Assistant.**

> This app template provides script `Create-Assistant.ps1` to help create assistant. The app assistant provides capabilities such as solving a math problem, calling functions for city weather and city nickname. Get more information in `ActionHandlers.cs`. You can change the instructions and settings in the script to customize assistant.
{{#useOpenAI}}
1. Ensure your OpenAI settings filled in `appsettings.TestTool.json`.
    ```
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>"
    }
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings filled in `appsettings.TestTool.json`.
    ```
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>",
      "OpenAIDeploymentName": "<your-azure-openai-deployment-name>"
    }
    ```
{{/useAzureOpenAI}}
1. Open PowerShell, change the current working directory to this project root and run command `. ./Create-Assistant.ps1`.
   ```
   > . ./Create-Assistant.ps1
   ```
1. The above command will display the properties of the newly created assistant, including the ID like "id: asst_xxx...". 
1. Fill in your assistant id in `env/.env.local.user`, `env/.env.dev.user` and `appsettings.TestTool.json`.

### Debug bot app in Teams App Test Tool
{{#useOpenAI}}
1. Ensure your OpenAI API Key filled in `appsettings.TestTool.json`.
    ```
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>",
      "AssistantId": "<your-openai-assistant-id>"
    }
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings filled in `appsettings.TestTool.json`.
    ```
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>",
      "OpenAIDeploymentName": "<your-azure-openai-deployment-name>",
      "OpenAIAssistantId": "<your-azure-openai-assistant-id>"
    }
    ```
{{/useAzureOpenAI}}
1. Set `Startup Item` as `Teams App Test Tool (browser)`.
![image](https://github.com/user-attachments/assets/41121a95-5c13-4bda-8f69-3b970a4e8f78)
1. Press F5, or select the Debug > Start Debugging menu in Visual Studio.
1. In Teams App Test Tool from the launched browser, type and send anything to your bot to trigger a response.
**Congratulations**! You are running an application that can now interact with users in Teams App Test Tool:

![AI Agent in Teams App Test Tool](https://github.com/OfficeDev/TeamsFx/assets/37978464/e3b458f3-5e74-460d-9df2-bf77ed8d9c54)

### Debug bot app in Teams Web Client

{{#useOpenAI}}
1. Ensure your OpenAI API Key filled in `env/.env.local.user`.
    ```
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"
    OPENAI_ASSISTANT_ID="<your-openai-assistant-id>"
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings filled in `env/.env.local.user`.
    ```
    SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
    AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
    AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
    AZURE_OPENAI_ASSISTANT_ID="<your-azure-openai-assistant-id>"
    ```
{{/useAzureOpenAI}}
1. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel.
2. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
3. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
4. Set `Startup Item` as `Microsoft Teams (browser)`.
5. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
6. In the opened web browser, select Add button to install the app in Teams
7. In the chat bar, type and send anything to your bot to trigger a response.

> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

## Extend the AI Chat Bot template with more AI capabilities
You can follow [Build an AI Agent in Teams](https://aka.ms/teamsfx-ai-agent) to extend the AI Agent template with more AI capabilities, like:
- [Customize assistant creation](https://aka.ms/teamsfx-ai-agent#customize-assistant-creation)
- [Add functions](https://aka.ms/teamsfx-ai-agent#add-functions-with-assistants-api)

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
