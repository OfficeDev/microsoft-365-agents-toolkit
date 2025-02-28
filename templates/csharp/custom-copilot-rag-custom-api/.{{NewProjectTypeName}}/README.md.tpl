# Overview of the Chat With Your Data (Using Custom API) template

This template showcases how to build an AI-powered intelligent chatbot that can understand natural language to invoke the API defined in the OpenAPI description document, so you can enable your users to chat with the data provided through API service.
The app template is built using the Teams AI library, which provides the capabilities to build AI-based Teams applications.
## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
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
{{/useAzureOpenAI}}
1. Set `Startup Item` as `Teams App Test Tool (browser)`.
![image](https://github.com/user-attachments/assets/41121a95-5c13-4bda-8f69-3b970a4e8f78)
1. Press F5, or select the Debug > Start Debugging menu in Visual Studio.
1. In Teams App Test Tool from the launched browser, type and send anything to your bot to trigger a response.

**Congratulations**! You are running an application that can now interact with users in Teams App Test Tool:

![custom api template](https://github.com/OfficeDev/TeamsFx/assets/63089166/81f985a1-b81d-4c27-a82a-73a9b65ece1f)

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
1. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
2. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
1. Set `Startup Item` as `Microsoft Teams (browser)`.
3. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
4. In the opened web browser, select Add button to install the app in Teams
1. In the chat bar, type and send anything to your bot to trigger a response.

> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

## Extend the template

- Follow [Build a Basic AI Chatbot in Teams](https://aka.ms/teamsfx-basic-ai-chatbot) to extend the template with more AI capabilities.
- Understand more about [build your own data ingestion](https://aka.ms/teamsfx-rag-bot#build-your-own-data-ingestion).

## Additional information and references

- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, you can create an issue directly in our GitHub repository: 
https://github.com/OfficeDev/TeamsFx/issues.
