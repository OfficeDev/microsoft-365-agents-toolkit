# Collaborator Agent for Microsoft Teams

This intelligent collaboration assistant is built with the [Microsoft Teams SDK](https://aka.ms/teamsai-v2), and showcases how to create a sophisticated bot that can analyze conversations, manage tasks, and search through chat history using advanced AI capabilities and natural language processing.

This agent can listen to all messages in a group chat (even without being @mentioned) using RSC (Resource Specific Control) permissions defined in [App Manifest](appPackage/manifest.json). For more details, see the documentation [RSC Documentation](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/resource-specific-consent).

## Key Features

- 📋 **Intelligent Summarization** - Analyze conversations and provide structured summaries with proper participant attribution and topic identification
- ✅ **Action Items** - Automatically identify and create action items from team discussions with smart assignment
- 🔍 **Conversation Search** - Search through chat history using natural language queries with time-based filtering and deep linking to original messages

## Adding Custom Capabilities

Adding your own capabilities only requires a few steps:

1. Copy the template folder under capabilities [template](Capabilities\Template\TemplateCapability.cs)
2. Customize your capability to do what you want (helpful to look at existing capabilities)
3. Make sure to create a CapabilityDefinition at the bottom of your main file
4. Register your capability by importing the CapabilityDefinition and adding to the definition list in [registry](Agent\CapabilityRegistry.cs)
5. The manager will automatically be instantiated with the capability you defined!

## Agent Architecture

![architecture](./img/architecture.png)

## Flow of the Agent

![flow](./img/flow.png)

If Collab Agent is added to a groupchat or private message, it will always listen and log each message to its database. The messages are stored in an SQLite DB by the conversation ID of the given conversation. 
The agent will respond whenever @mentioned in groupchats and will always respond in 1-on-1 messages. When the agent responds, the request is first passed through a manger prompt.
This manager may route to a capability based on the request--this capability returns its result back to the manager where it will be passed back to the user.

### Running the Bot

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
1. Ensure your OpenAI API Key is filled in `appsettings.Playground.json`.
    ```
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>"
    }
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings are filled in `appsettings.Playground.json`.
    ```
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>",
      "OpenAIDeploymentName": "<your-azure-openai-deployment-name>"
    }
    ```
{{/useAzureOpenAI}}
1. Set `Startup Item` as `Microsoft 365 Agents Playground (browser)`.
![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/switch-to-test-tool.png)
1. Press F5, or select the Debug > Start Debugging menu in Visual Studio.
1. In Microsoft 365 Agents Playground from the launched browser, type and send anything to your agent to trigger a response.

### Debug agent in Teams Web Client
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
2. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
3. Sign in to Microsoft 365 Agents Toolkit with a **Microsoft 365 work or school account**
4. Set `Startup Item` as `Microsoft Teams (browser)`.
5. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
6. In the opened web browser, select Add button to install the app in Teams
7. In the chat bar, type and send anything to your agent to trigger a response.

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).
#### Sample Questions

You can ask the Collaborator agent questions like:

**Summarization:**
- "@Collaborator summarize yesterday's discussion"
- "@Collaborator what were the main topics from last week?"
- "@Collaborator give me an overview of recent messages"

**Action Items:**
- "@Collaborator find action items from the past 3 days"
- "@Collaborator create a task to review the proposal by Friday"
- "@Collaborator what tasks are assigned to me?"

**Search:**
- "@Collaborator find messages about the project deadline"
- "@Collaborator search for conversations between Alice and Bob"
- "@Collaborator locate discussions from this morning about the budget"

## Architecture

The Collaborator agent uses a sophisticated multi-capability architecture:

- **Manager**: Coordinates between specialized capabilities and handles natural language time parsing
- **Summarizer**: Analyzes conversation content and provides structured summaries
- **Action Items**: Identifies tasks, manages assignments, and tracks completion
- **Search**: Performs semantic search across conversation history with citation support
- **Context Management**: Global message context handling for concurrent request support

## Deployment

The agent can be deployed to Azure App Service for production use. See following documentation for detailed instructions on setting up Azure resources and configuring the production environment.
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- Azure SQL Database is used to store data, you can set admin password in `env/.env.dev.user`.

If you are trying to local debug / preview deployed version, then either of the two conditions must be met as a min-bar:
1. The user doing the operation should be an admin in the org.
2. The Entra app ID specified in webAppInfo.Id must be homed in the same tenant.
