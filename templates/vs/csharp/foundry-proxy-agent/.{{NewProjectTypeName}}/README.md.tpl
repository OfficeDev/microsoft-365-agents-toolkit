# Overview of the Microsoft Foundry Proxy Agent

This template showcases an agent that proxies requests to Microsoft Azure AI Foundry agents, enabling Teams and M365 Copilot integration with your Foundry agents.

The app template is built using the Microsoft Agents SDK, which provides the capabilities to build AI-based applications.

## Quick Start

**Prerequisites**
To run the agent template in your local dev machine, you will need:

- An [Azure AI Foundry](https://ai.azure.com) project with an agent deployed
- The Foundry Project Endpoint and Agent ID from your Azure AI Foundry project
- Azure subscription with appropriate permissions

> **Important:** The Microsoft 365 account used must have the **Azure AI User** role on your Azure AI Foundry project. You can assign this permission in the [Azure AI Foundry portal](https://ai.azure.com) under **Management Center > Resource > Users**.

### Debug agent in Teams Web Client

1. Ensure your Azure AI Foundry settings are filled in `env/.env.local.user`:
    ```
    AZURE_AI_FOUNDRY_PROJECT_ENDPOINT="<your-foundry-project-endpoint>"
    AGENT_ID="<your-foundry-agent-id>"
    ```
2. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel.
3. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
4. Sign in to Microsoft 365 Agents Toolkit with a **Microsoft 365 work or school account**
5. Set `Startup Item` as `Microsoft Teams (browser)`.
6. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
7. In the opened web browser, select Add button to install the app in Teams
8. In the chat bar, type and send anything to your agent to trigger a response.

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

## Deploy to Azure

To deploy this project to Azure for production use:

1. Ensure your Azure AI Foundry settings are filled in `env/.env.dev.user`:
    ```
    AZURE_AI_FOUNDRY_PROJECT_ENDPOINT="<your-foundry-project-endpoint>"
    AGENT_ID="<your-foundry-agent-id>"
    ```
1. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Provision to the Cloud**
2. Select your Azure subscription and resource group
3. After provisioning completes, right-click the '{{NewProjectTypeName}}' project and select **Microsoft 365 Agents Toolkit > Deploy to the Cloud**
4. To preview, right-click the '{{NewProjectTypeName}}' project and select **Microsoft 365 Agents Toolkit > Preview in > Teams**

- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)

## How it works

This agent acts as a proxy between Microsoft Teams/M365 Copilot and your Azure AI Foundry agent:

1. User sends a message in Teams or M365 Copilot
2. The agent authenticates the user via SSO
3. The user's token is exchanged for an Azure AI Foundry-compatible token
4. The message is forwarded to your Foundry agent
5. The response is streamed back to the user in real-time

## Configuration

The agent requires the following configuration in `appsettings.json`:

- `AIServices:AzureAIFoundryProjectEndpoint` - Your Azure AI Foundry project endpoint
- `AIServices:AgentID` - The ID of your deployed Foundry agent

## Additional information and references
- [Azure AI Foundry Documentation](https://learn.microsoft.com/azure/ai-foundry/)
- [Microsoft Agents SDK](https://github.com/microsoft/agents)
- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

## Learn more

New to app development or Microsoft 365 Agents Toolkit? Learn more about app manifests, deploying to the cloud, and more in the documentation
at https://aka.ms/teams-toolkit-vs-docs.

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem.
Or, you can create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues.
