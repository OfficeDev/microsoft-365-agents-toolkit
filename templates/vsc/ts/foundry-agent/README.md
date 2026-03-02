# Overview of the Basic Foundry Agent template

This app template is built on top of [Microsoft 365 Agents SDK](https://github.com/Microsoft/Agents).
It showcases an agent that responds to user questions by connecting to a Microsoft Foundry (formerly Azure AI Foundry) agent using the Responses API.

## Microsoft Foundry Configuration

This template is configured to connect to a Microsoft Foundry agent using agent reference in the Responses API. The configuration requires:

### Prerequisites
- A Microsoft Foundry project with a deployed agent
- Azure credentials configured for `DefaultAzureCredential` (e.g., Azure CLI login, managed identity, or environment variables)
- Required environment variables:
  - `FOUNDRY_PROJECT_ENDPOINT`: Your Foundry project endpoint (e.g., `https://your-resource.services.ai.azure.com/api/projects/your-project`)
  - `FOUNDRY_AGENT_NAME`: The name of your agent in Foundry (e.g., `mail-assistant`)

### How It Works
The Teams app uses the Azure AI Projects SDK (`@azure/ai-projects`) to:
1. Connect to your Microsoft Foundry project using `AIProjectClient`
2. Retrieve the agent configuration by name
3. Create an Azure OpenAI client from the project
4. Send user messages to the agent using the Responses API with agent reference
5. Return the agent's response to the Teams user

This approach allows you to leverage your existing Foundry agents directly within Microsoft Teams without duplicating agent logic.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22.
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) latest version or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli).
> - A Microsoft Foundry project with a deployed agent (e.g., `mail-assistant`)
> - Azure credentials configured (run `az login` if using Azure CLI)

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
1. Configure your Microsoft Foundry connection in the environment files:
  - In `env/.env.playground.user` (recommended) or `env/.env.playground`, update:
     - `FOUNDRY_PROJECT_ENDPOINT`: Your Foundry project endpoint
     - `FOUNDRY_AGENT_NAME`: Your agent name (default: `mail-assistant`)
1. Ensure you're authenticated to Azure (the app uses `DefaultAzureCredential`):
   ```bash
   az login
   ```
1. Press F5 to start debugging which launches your agent in Microsoft 365 Agents Playground using a web browser. Select `Debug in Microsoft 365 Agents Playground`.
1. You can send any message to get a response from the agent powered by your Microsoft Foundry agent.

**Congratulations**! You are running an agent that can now interact with users in Microsoft 365 Agents Playground using Microsoft Foundry:

![Basic AI Agent](https://github.com/user-attachments/assets/984af126-222b-4c98-9578-0744790b103a)

## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `src`        | The source code for the application                 |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/index.ts`| Sets up the agent server.|
|`src/adapter.ts`| Sets up the agent adapter.|
|`src/config.ts`| Defines the environment variables.|
|`src/agent.ts`| Handles business logics for the Basic Foundry Agent.|

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`m365agents.yml`|This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`m365agents.local.yml`|This overrides `m365agents.yml` with actions that enable local execution and debugging.|
|`m365agents.playground.yml`| This overrides `m365agents.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.|

## Additional information and references

- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Known issue
- The agent is currently not working in any Teams group chats or Teams channels when the stream response is enabled.
