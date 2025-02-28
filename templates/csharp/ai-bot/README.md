# Overview of the AI Chat Bot template

This template showcases a bot app that responds to user questions like an AI assistant. This enables your users to talk with the AI assistant in Teams to find information.

The app template is built using the Teams AI library, which provides the capabilities to build AI-based Teams applications.

## Quick Start

**Prerequisites**
> To run the AI Chat Bot template in your local dev machine, you will need:
>
> - [Azure OpenAI](https://aka.ms/oai/access) resource or an account with [OpenAI](https://platform.openai.com).

### Debug bot app in Teams Web Client

1. Fill in your OpenAI API Key or Azure OpenAI settings in `env/.env.local.user`
    ```
    # If using OpenAI
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"

    # If using Azure OpenAI
    SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
    SECRET_AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
    ```

2. If using Azure OpenAI, update "gpt-35-turbo" in `Program.cs` to your own model deployment name
3. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
1. Right-click the 'TeamsApp' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
2. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
3. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
4. In the opened web browser, select Add button to install the app in Teams
8. In the chat bar, type and send anything to your bot to trigger a response

> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

### Debug bot app in Teams App Test Tool

1. Fill in your OpenAI API Key or Azure OpenAI settings in `appsettings.TestTool.json`
    ```
    # If using OpenAI
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>"
    },

    # If using Azure OpenAI
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>"
    }
    ```

2. If using Azure OpenAI, update "gpt-35-turbo" in `Program.cs` to your own model deployment name
3. Select `Teams App Test Tool (browser)` in debug dropdown menu
4. Press F5, or select the Debug > Start Debugging menu in Visual Studio
5. In Teams App Test Tool from the launched browser, type and send anything to your bot to trigger a response

## Extend the AI Chat Bot template with more AI capabilities

You can follow [Get started with Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/how-conversation-ai-get-started) to extend the AI Chat Bot template with more AI capabilities.

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
