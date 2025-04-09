# Overview of the Weather Agent template

This template showcases a bot app that responds to user questions about weather like an AI agent. This enables your users to talk with the AI agent in Teams to find weather information.

The app template is built using the Microsoft 365 Agents SDK and Semantic Kernel, which provides the capabilities to build AI-based Teams applications.

## Quick Start

**Prerequisites**
> To run the Weather Agent template in your local dev machine, you will need:
>
{{#useOpenAI}}
> - an account with [OpenAI](https://platform.openai.com).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - [Azure OpenAI](https://aka.ms/oai/access) resource
{{/useAzureOpenAI}}

### Debug bot app in Teams App Test Tool
{{#useOpenAI}}
1. Ensure your OpenAI API Key is filled in `appsettings.TestTool.json`.
    ```
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>"
    }
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings are filled in `appsettings.TestTool.json`.
    ```
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>",
      "OpenAIDeploymentName": "<your-azure-openai-deployment-name>"
    }
    ```
  > The domain of your endpoint should be as `.cognitiveservices.azure.com`. You can find it from [Azure AI Foundry](https://ai.azure.com/resource/overview/) 
{{/useAzureOpenAI}}
1. Set `Startup Item` as `Teams App Test Tool (browser)`.
![image](https://github.com/user-attachments/assets/41121a95-5c13-4bda-8f69-3b970a4e8f78)
1. Press F5, or select the Debug > Start Debugging menu in Visual Studio.
1. In Teams App Test Tool from the launched browser, type and send anything to your bot to trigger a response.

### Debug bot app in Teams Web Client
{{#useOpenAI}}
1. Ensure your OpenAI API Key is filled in `env/.env.local.user`.
    ```
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings are filled in `env/.env.local.user`.
    ```
    SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
    AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
    AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
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

## Additional information and references
- [Microsoft 365 Agent SDK](https://github.com/microsoft/Agents)
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
