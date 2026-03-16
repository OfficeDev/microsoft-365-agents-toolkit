# Welcome to Microsoft 365 Agents Toolkit!

## Overview

This project is a **Microsoft Foundry Proxy Agent** that enables your Azure AI Foundry agents to work seamlessly with Microsoft Teams and Microsoft 365 Copilot. It acts as a bridge between Microsoft 365 platforms and your Foundry-hosted AI agent.

### This template illustrates
- How to connect an AI Foundry Agent to M365 Copilot
- How to setup and use the Agent SDK with managed Identity so you no longer maintain secrets
- How to setup and use SSO in M365 Copilot & Teams and pass the user token to AI Foundry using Agent SDK
- How to Configure SSO with Federated Credentials so your SSO flow does not have any secrets (Single Tenant Only)

### 🔄 Architecture Flow

```mermaid
sequenceDiagram
    %% Groups
    box "User"
        participant U as Copilot User
    end

    box "Microsoft 365"
        participant M as Microsoft 365 Copilot
    end

    box "Custom Engine Agent"
        participant B as Azure Bot Service
        participant P as Proxy Agent (Agents SDK)
    end

    box "Microsoft Foundry"
        participant A as AI Agent Backend
    end

    %% Flow
    U->>M: User prompt (e.g., "Create a report")
    M->>B: Activity (Message)
    B->>P: POST /api/messages (Message)
    P->>B: 202 Accepted
    P-->>M: Start Streaming Session with information
    P->>A: POST /process { prompt }
    A-->>P: { content }
    P-->>M: Stream(content)
    M-->>U: Display result
```

This proxy pattern allows you to:
- ✅ Connect existing AI agents to Microsoft 365 Copilot
- ✅ Maintain your AI logic in Microsoft Foundry
- ✅ Provide seamless user experience in Teams and Copilot with SSO
- ✅ Handle authentication and message routing automatically

## Prerequisites

Before you begin, ensure you have:
- An **Azure AI Foundry** project with a deployed agent, ensure the AI Foundry is in the same tenant with your m365 and your azure account. SSO is only available in single tenant.
- Your **Foundry Project Endpoint** (e.g., `https://your-project.services.ai.azure.com/api/projects/your-project`)
- Your **Agent ID** (e.g., `asst_xxxxxxxxxxxxx`)

### Required Azure Permissions

| Permission | Scope | Purpose |
|------------|-------|---------|
| **Contributor** | Subscription or Resource Group | Deploy Bot Service |
| **Application Administrator** | Entra ID | Create app registrations |
| **Azure AI User** | Azure AI Foundry Project | Allow M365 account to access the Foundry agent |

> **Important:** The Microsoft 365 account used must have the **Azure AI User** role on your Azure AI Foundry project. You can assign this permission in the [Azure AI Foundry portal](https://ai.azure.com) under **Management Center > Resource > Users**.

## Quick Start

### Step 1: Configure Dev Tunnel (Required for Local Debug)

Before debugging, you need to create a Dev Tunnel in Visual Studio:

1. In the **Debug** dropdown menu, select **Dev Tunnels > Create A Tunnel**
2. Configure the tunnel:
   - **Access:** Public
   - **Persistence:** Permanent (recommended)
3. Visual Studio will automatically update your environment with the tunnel URL

> **Note:** Dev Tunnel provides a secure HTTPS endpoint that allows Microsoft Teams/Copilot to communicate with your locally running bot.

### Step 2 (Optional): Verify Azure AI Foundry Settings

Your Foundry configuration should already be set in `env/.env.local.user` (from project creation). If you need to update these values, edit `env/.env.local.user` directly:

```
AZURE_AI_FOUNDRY_PROJECT_ENDPOINT=<your-foundry-endpoint>
AGENT_ID=<your-agent-id>
```

If you need to update these values, edit `env/.env.local.user` directly.

### Step 3: Start Debugging

1. Right-click the **M365Agent** project in Solution Explorer
2. Select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
3. Sign in with a **Microsoft 365 work or school account**
4. Set **Startup Item** to `Microsoft Teams (browser)` or `Microsoft M365 Copilot (browser)`
5. Press **F5** to start debugging

> **Having issues?** If you encounter any errors during or after debugging, see the [Troubleshooting](#troubleshooting) section below.

## What Happens During Local Debug (F5)

When you press F5, Microsoft 365 Agents Toolkit automates several steps:

### First-Time Provisioning (ARM Deploy)
On first run, the toolkit deploys Azure infrastructure for local development:

1. **Creates Bot App Registration** - An Entra ID application for bot authentication
2. **Deploys Azure Bot Service** - Routes messages between Teams/Copilot and your local bot
3. **Creates SSO App Registration** - Enables single sign-on with your Foundry agent
4. **Configures OAuth Connection** - Sets up token exchange for Foundry authentication

These resources are created in Azure but your bot code runs locally, enabling full debugging with breakpoints.

### Every Debug Session
- Validates prerequisites (Node.js, M365 account, ports)
- Updates Teams app manifest with current tunnel URL
- Starts your bot application with debugger attached
- Opens Teams/Copilot with your agent ready to test

## Debug in Teams

1. Ensure Dev Tunnel is configured (Step 1 above)
2. Right-click **M365Agent** project > **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
3. Sign in with a **Microsoft 365 work or school account**
4. Set `Startup Item` as `Microsoft Teams (browser)`
5. Press **F5** to start debugging
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
6. In the browser, select **Add** to install the app in Teams
7. Chat with your agent - messages are proxied to your Azure AI Foundry agent

## Debug in Microsoft 365 Copilot

1. Ensure Dev Tunnel is configured (Step 1 above)
2. Right-click **M365Agent** project > **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
3. Sign in with a **Microsoft 365 work or school account**
4. Set `Startup Item` as `Microsoft M365 Copilot (browser)`
5. Press **F5** to start debugging
6. In Copilot, @mention your agent to interact with it

## Deploy to Azure

You can deploy this project to Azure App Service for production use.

### Prerequisites

- An Azure subscription with **Contributor** access
- Your Azure AI Foundry Project Endpoint and Agent ID

### Provision and Deploy


2. Right-click the **M365Agent** project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Provision to the Cloud**
3. Select your Azure subscription and resource group (or create a new one)
4. Wait for provisioning to complete — this deploys:
   - Managed Identity
   - Azure App Service
   - Azure Bot Service
   - App Registration
   - OAuth Connection
5. Right-click the **M365Agent** project and select **Microsoft 365 Agents Toolkit > Deploy to the Cloud**
6. Wait for the deployment to complete

### Preview in Teams (Remote)

After deployment:

1. Right-click the **M365Agent** project and select **Microsoft 365 Agents Toolkit > Preview in > Teams**
2. The app will open in Teams using the deployed Azure resources
3. You can also test via **Microsoft 365 Copilot** by @mentioning your agent

> **Note:** Remote deployment uses Managed Identity for authentication instead of client secrets, providing a more secure production configuration.

## Project Structure

```
├── M365Agent/                    # Microsoft 365 Agents Toolkit project
│   ├── appPackage/              # Teams app manifest
│   ├── env/                     # Environment configuration
│   │   ├── .env.local           # Auto-generated local settings
│   │   ├── .env.local.user      # Your Foundry configuration (gitignored)
│   │   ├── .env.dev             # Dev environment settings
│   │   └── .env.dev.user        # Dev Foundry configuration (gitignored)
│   ├── infra/                   # Azure Bicep templates
│   │   ├── azure-local.bicep    # Local debug infrastructure
│   │   └── azure.bicep          # Production infrastructure
│   └── m365agents.local.yml     # Local debug workflow
├── Agents/                      # Agent implementation
│   └── AzureAgent.cs           # Foundry proxy logic
├── Program.cs                   # Application entry point
└── appsettings.json            # Application configuration
```

## Environment Files Explained

| File | Purpose | Git Status |
|------|---------|------------|
| `.env.local` | Auto-generated values (Bot ID, App IDs, tunnel URL) | Ignored |
| `.env.local.user` | Your secrets and Foundry config | Ignored |
| `.env.dev` | Dev environment auto-generated values | Committed |
| `.env.dev.user` | Dev environment secrets | Ignored |

## Troubleshooting

### "No Dev Tunnel configured"
Create a Dev Tunnel via **Debug > Dev Tunnels > Create A Tunnel** with Public access.

### "Agent not responding"
1. Verify your `AZURE_AI_FOUNDRY_PROJECT_ENDPOINT` is correct
2. Verify your `AGENT_ID` exists in Foundry
3. Check that your Foundry agent is deployed and running

### "Authentication failed"
Ensure you're signed into Microsoft 365 Agents Toolkit with a work/school account that has access to your Azure subscription.

### "AADSTS650052: Azure Machine Learning Services lacks a service principal"

**Symptom:** The SSO token exchange fails at runtime with:
```
AuthenticationFailed / invalid_client
AADSTS650052: The app is trying to access a service '18a66f5f-dbdf-4c17-9dd7-1634712a9cbe'
(Azure Machine Learning Services) that your organization lacks a service principal for.
```

**Root cause:** This proxy agent exchanges the signed-in user's token for an `Azure Machine Learning | user_impersonation` token to call Azure AI Foundry. If your tenant has never provisioned any Azure ML or AI Foundry resources, the **Azure Machine Learning Services** enterprise application doesn't exist in your Entra ID tenant, and the token exchange fails.

**Solution (requires Tenant Admin — one-time):**
```powershell
# Option 1 – Azure CLI
az ad sp create --id 18a66f5f-dbdf-4c17-9dd7-1634712a9cbe

# Option 2 – Azure PowerShell
New-AzADServicePrincipal -ApplicationId 18a66f5f-dbdf-4c17-9dd7-1634712a9cbe
```

> **Note:** If the command fails with _"The service principal cannot be created because the service principal name `https://containeragents.ai.azure.com` is already in use"_, this means the **Azure Machine Learning Services** enterprise application is already present in your tenant. You can ignore the error and retry the agent directly — no further action is needed.

After the command completes (or if the SP already exists), retry the agent — no restart is required.

## Learn More

- [Microsoft 365 Agents Toolkit Documentation](https://aka.ms/teams-toolkit-vs-docs)
- [Azure AI Foundry Documentation](https://learn.microsoft.com/azure/ai-studio/)
- [Teams App Development](https://learn.microsoft.com/microsoftteams/platform/)

## Report an Issue

Select **Visual Studio > Help > Send Feedback > Report a Problem**
Or create an issue at: https://github.com/OfficeDev/TeamsFx/issues
