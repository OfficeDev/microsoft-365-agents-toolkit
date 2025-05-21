# Overview of the AI Agent template

This template showcases an agent that responds to user questions like an AI assistant. This enables your users to talk with the AI assistant in Teams to find information.

The app template is built using the Teams AI library, which provides the capabilities to build AI-based applications.

## Quick Start

**Prerequisites**
> To run the agent template in your local dev machine, you will need:
>
{{#useOpenAI}}
> - an account with [OpenAI](https://platform.openai.com).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - [Azure OpenAI](https://aka.ms/oai/access) resource
{{/useAzureOpenAI}}

### Debug agent in Microsoft 365 Agents Playground
{{#useOpenAI}}
1. Ensure your OpenAI API Key filled in `appsettings.Playground.json`
    ```
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>"
    }
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings filled in `appsettings.Playground.json`
    ```
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>",
      "OpenAIDeploymentName": "<your-azure-openai-deployment-name>"
    }
    ```
{{/useAzureOpenAI}}
1. Set `Startup Item` as `Microsoft 365 Agents Playground (browser)`.
![image](https://github.com/user-attachments/assets/41121a95-5c13-4bda-8f69-3b970a4e8f78)
1. Press F5, or select the Debug > Start Debugging menu in Visual Studio
1. In Microsoft 365 Agents Playground from the launched browser, type and send anything to your agent to trigger a response

**Congratulations**! You are running an application that can now interact with users in Microsoft 365 Agents Playground:

![AI Agent](https://github.com/user-attachments/assets/5c38d129-f033-4f17-8ecc-23f827d7c1fd)

### Debug agent in Teams Web Client

{{#useOpenAI}}
1. Ensure your OpenAI API Key filled in `env/.env.local.user`
    ```
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings filled in `env/.env.local.user`
    ```
    SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
    AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
    AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
    ```
{{/useAzureOpenAI}}
1. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
2. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
3. Sign in to Microsoft 365 Agents Toolkit with a **Microsoft 365 work or school account**
4. Set `Startup Item` as `Microsoft Teams (browser)`.
5. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
6. In the opened web browser, select Add button to install the app in Teams
7. In the chat bar, type and send anything to your agent to trigger a response

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

## Extend the agent template with more AI capabilities

You can follow [Get started with Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/how-conversation-ai-get-started) to extend the agent template with more AI capabilities.

## Additional information and references
- [Teams AI library](https://aka.ms/teams-ai-library)
- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Learn more

New to app development or Microsoft 365 Agents Toolkit? Learn more about app manifests, deploying to the cloud, and more in the documentation 
at https://aka.ms/teams-toolkit-vs-docs.

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, you can create an issue directly in our GitHub repository: 
https://github.com/OfficeDev/TeamsFx/issues.
