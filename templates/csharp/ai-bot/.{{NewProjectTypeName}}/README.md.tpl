# Overview of the AI Chat Bot template

This template shows a bot app that acts like an AI agent, answering user questions. Your users can chat with this AI agent in Teams to get information.

The app template is built using the Teams AI library, which helps build AI-based Teams applications.

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
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/create-devtunnel-button.png)
4. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
5. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
6. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
7. In the opened web browser, select Add button to install the app in Teams
8. In the message input field, type and send anything to your bot to get a response

## Run the app on other platforms

The Teams app can run in other platforms like Outlook and Microsoft 365 app. See https://aka.ms/vs-ttk-debug-multi-profiles for more details.

## Extend AI Chat Bot template with more AI capabilities

Follow [Get started with Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/how-conversation-ai-get-started) to enhance the AI Chat Bot template with advanced features.

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
