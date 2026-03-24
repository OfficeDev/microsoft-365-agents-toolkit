---
name: microsoft-365-agents-toolkit
description: Builds, tests, and deploys Microsoft 365 apps or agents working on Teams and Copilot. Used when working with Teams agent, bot, tab, message extension, Declarative Agents, Custom Engine Agents using Microsoft 365 Agents Toolkit and Playground. It helps to do local testing, Azure resource provision, remote deployment, or any Microsoft 365 application development workflow.
---

# Microsoft 365 Agents Toolkit Skill

Use this skill to run end-to-end Microsoft 365 app and agent workflows with the Microsoft 365 Agents Toolkit (ATK) CLI.

# Prerequisites

ATK CLI is required:

```bash
npm i -g @microsoft/m365agentstoolkit-cli@latest
atk --version  # Requires >= 1.1.5
```

# CLI Global Options

| Option | Meaning | Recommendation |
| --- | --- | --- |
| `-i` | Interactive mode | Always use `-i false` in automation to avoid hanging |
| `-f` | Project folder | Default to be current directory, used when specifying a custom folder. When scaffolding a new project, this is the parent folder where the project folder will be created under. |
| `-h` | Command help | Use `atk <command> -h` for quick syntax checks |

# Agent/App Development Lifecycle

Match user intent to the smallest valid workflow.

| User Intent                                 | Workflow (read in order)                        |
| ------------------------------------------- | ----------------------------------------------- |
| Create a new agent from scratch             | [create](#create) → [debug](#debug)             |
| Debug an agent (code already exists)        | debug → [remote](#remote) → [preview](#preview) |
| Update/redeploy an agent after code changes | remote → preview                                |
| Test/chat with a deployed agent/app         | preview                                         |

# Workflow

## Create

Create from template:

```bash
# language options: typescript, javascript, python
atk new -c <template-id> -n <project-name> -f <folder-name> -l <language> -i false
```

Create from sample:

```bash
atk new sample <sample-id>
```

### Templates

| Capability                                   | Description                                               |
| -------------------------------------------- | --------------------------------------------------------- |
| `declarative-agent`                          | Declarative Agent                                         |
| `declarative-agent-action`                   | Declarative Agent with Action from Scratch                |
| `declarative-agent-action-bearer`            | Declarative Agent with Action from Scratch (Bearer Token) |
| `declarative-agent-action-oauth`             | Declarative Agent with Action from Scratch (OAuth)        |
| `declarative-agent-action-from-existing-api` | Declarative Agent with Action from Existing API           |
| `declarative-agent-with-action-from-mcp`     | Declarative Agent with Action from MCP Server             |
| `declarative-agent-with-graph-connector`     | Declarative Agent with Copilot Connector                  |
| `declarative-agent-meta-os-new-project`      | Declarative Agent for MetaOS (New Project)                |
| `declarative-agent-typespec`                 | Declarative Agent from TypeSpec                           |
| `basic-custom-engine-agent`                  | Basic Custom Engine Agent                                 |
| `weather-agent`                              | Weather Agent                                             |
| `foundry-agent-to-m365`                      | Foundry Agent to M365                                     |
| `copilot-connector`                          | Copilot Connector                                         |
| `teams-agent`                                | General Teams Agent                                       |
| `teams-agent-rag-customize`                  | Teams Agent with Data from Customized Source              |
| `teams-agent-rag-azure-ai-search`            | Teams Agent with Data from Azure AI Search                |
| `teams-agent-rag-custom-api`                 | Teams Agent with Data from Custom API using OpenAPI Spec  |
| `teams-collaborator-agent`                   | Teams Collaborator Agent                                  |
| `tab`                                        | Tab                                                       |
| `bot`                                        | Simple Bot                                                |
| `message-extension`                          | Message Extension                                         |

Use `atk list templates` if a template ID fails. Refer to [Templates](./references/templates.md) for complete details.

### Template Selection Guide

| Family | Use When |
| --- | --- |
| Declarative Agent | Copilot extension with low/no custom runtime |
| Custom Engine Agent | Cross-platform logic, custom LLM orchestration, advanced multi-turn behavior |
| Teams Agent | Teams-first experiences, RAG in Teams, Teams-specific capabilities |

### Samples

| Sample                                            | Sample ID (`atk new sample <sample-id>`) | Tags                                                                                      |
| ------------------------------------------------- | ---------------------------------------- | ----------------------------------------------------------------------------------------- |
| Langchain Agent with Agent365 SDK in NodeJS       | `agent365-langchain-nodejs`              | Agent365, TS                                                                              |
| Agent Framework Agent with Agent365 SDK in Python | `agent365-agentframework-python`         | Agent365, Python                                                                          |
| OpenAI Agent with Agent365 SDK in Python          | `agent365-openai-python`                 | Agent365, Python                                                                          |
| Claude Agent with Agent365 SDK in NodeJS          | `agent365-claude-nodejs`                 | Agent365, TS                                                                              |
| Tab App with Azure Backend                        | `hello-world-tab-with-backend`           | Tab, TS, Azure Functions, Dev Proxy                                                       |
| Bot App with SSO Enabled                          | `bot-sso`                                | Bot, TS, Adaptive Cards, SSO                                                              |
| Team Central Dashboard                            | `team-central-dashboard`                 | Tab, TS, Azure Functions, SSO                                                             |
| Copilot connector App                             | `copilot-connector-app`                  | Tab, Azure Functions, TS, SSO, Copilot connector                                          |
| Teams Conversation Bot using Python               | `bot-conversation-python`                | Python, Bot, Bot Framework                                                                |
| Teams Messaging Extensions Search using Python    | `msgext-search-python`                   | Python, Message extension, Bot Framework                                                  |
| Travel Agent                                      | `travel-agent`                           | C#, Custom Engine Agent, M365 Copilot Retrieval API, Agents SDK, Agent Framework          |
| Coffee Agent                                      | `coffee-agent`                           | TS, Custom Engine Agent, Adaptive Cards, Microsoft Teams SDK                              |
| Data Analyst Agent v2                             | `data-analyst-agent-v2`                  | TS, Custom Engine Agent, Data Visualization, Adaptive Cards, LLM SQL, Microsoft Teams SDK |

List all samples with `atk list samples`.

## Debug

Use one of two local debug paths.

| Path | Best For | Account Needed |
| --- | --- | --- |
| Agents Playground | Fast bot/message-extension inner loop | No |
| Teams sideloading | Full-fidelity Teams behavior and auth/SSO | Yes |

### Common bot startup check (Bash)

For bot projects using port `3978`, run this shared Bash flow before either Playground debug or Teams sideloading. It frees the port, starts the bot, and verifies the listener before continuing.

```bash
# Save as start-bot-safe.sh, then run:
#   bash ./start-bot-safe.sh "npm run dev:teamsfx:playground"
#   bash ./start-bot-safe.sh "npm run dev:teamsfx"

set -euo pipefail

PORT=3978
START_CMD=${1:-"npm run dev:teamsfx:playground"}

# 1) Kill any process currently listening on the bot port.
if command -v lsof >/dev/null 2>&1; then
  PIDS=$(lsof -ti tcp:${PORT} || true)
elif command -v netstat >/dev/null 2>&1; then
  # Git Bash on Windows often has netstat even when lsof is missing.
  PIDS=$(netstat -ano 2>/dev/null | grep ":${PORT}" | grep LISTENING | awk '{print $5}' | tr -d '\r' | sort -u)
else
  PIDS=""
fi

if [ -n "${PIDS:-}" ]; then
  kill -9 ${PIDS}
fi

# 2) Start bot command in background.
bash -lc "${START_CMD}" &
BOT_PID=$!

# 3) Wait up to 30s for listener readiness.
READY=0
for _ in $(seq 1 30); do
  if command -v lsof >/dev/null 2>&1 && lsof -i tcp:${PORT} -sTCP:LISTEN >/dev/null 2>&1; then
    READY=1
    break
  fi
  if command -v netstat >/dev/null 2>&1 && netstat -ano 2>/dev/null | grep ":${PORT}" | grep LISTENING >/dev/null 2>&1; then
    READY=1
    break
  fi
  sleep 1
done

# 4) Fail fast if startup did not succeed.
if [ ${READY} -ne 1 ]; then
  echo "Bot did not start on port ${PORT} within 30 seconds." >&2
  kill -9 ${BOT_PID} >/dev/null 2>&1 || trues
  exit 1
fi

echo "Bot is listening on port ${PORT}."
echo "Bot PID: ${BOT_PID}"
```

This avoids stale listener issues and works as a common pre-step for both debug paths.

### Debug in Agents Playground

```bash
winget install agentsplayground
# For Agents Toolkit project
atk deploy --env playground -i false
npm run dev:teamsfx:playground
# For customized project
npm run dev # Check package.json for similar command to start the bot
# In a new terminal, after bot starts:
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

Use `bash ./start-bot-safe.sh "npm run dev:teamsfx:playground"` first, then run Playground launch. For troubleshooting, see [playground](./references/playground.md).

### Debug in Teams

For bot projects, expose port `3978` and set `BOT_ENDPOINT`:

```bash
devtunnel host -p 3978 --allow-anonymous
# Set BOT_ENDPOINT in env/.env.local with the tunnel URL
```

Provision and deploy local environment:

```bash
atk provision --env local -i false
atk deploy --env local -i false
```

Start the service:

```bash
# Check package.json for start command: npm run dev:teamsfx, npm run dev, npm start, etc.
npm run dev:teamsfx
```

Use `bash ./start-bot-safe.sh "npm run dev:teamsfx"` to apply the same port cleanup and startup verification before opening a dev tunnel.

Open Teams sideloading URL:

```
https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true
```

For declarative agents (no backend service), provision and open M365 Copilot URL directly.

#### Opening Declarative Agents in M365 Copilot

Use `M365_APP_ID` (not `TEAMS_APP_ID`), generated during provisioning.

Sideloading URL format:

```
https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&developerMode=Basic
```

`${agent-hint}` is Base64-encoded JSON. Example payload:

```json
{
  "id": "${M365_APP_ID}",
  "scenario": "launchcopilotextension",
  "properties": { "clickTimestamp": "${CURRENT_TIMESTAMP}" },
  "version": 1
}
```

See [teams-debug.md](./references/teams-debug.md) for complete steps.

### Comparison

| Feature                | Agents Playground           | Teams Direct Launch                       |
| ---------------------- | --------------------------- | ----------------------------------------- |
| Setup complexity       | Simple                      | Requires provisioning                     |
| M365 account needed    | No                          | Yes                                       |
| HTTPS required         | No                          | Yes (for bots)                            |
| Real Teams environment | No (simulated)              | Yes                                       |
| SSO testing            | No                          | Yes                                       |
| Speed                  | Fast                        | Slower (tunnel setup)                     |
| Recommended for        | Testing first (recommended) | When user explicitly asks to run on Teams |

## Remote

Deploy to Azure after local validation.

### Provision and Deploy

`atk provision` and `atk deploy` execute actions defined in `m365agentstoolkit*.yml` files:

0. Check required environment variables referenced in `m365agentstoolkit*.yml`.
Copy needed values from local env files into `env/.env.dev`.

1. Configure Azure subscription in `env/.env.dev`:
`AZURE_SUBSCRIPTION_ID=your-subscription-id`

2. Create resource group if needed:
`az group create --name <rg> --location <region>`
IMPORTANT: Verify az account matches atk account: `az account show` vs `atk auth list`

3. Provision Azure and M365 resources:
`atk provision --env dev --resource-group <rg> --region <region> -i false`

4. Deploy code to Azure:
`atk deploy --env dev -i false`

**Environment variables are stored in:**
- `env/.env.dev` - Non-secret configuration
- `env/.env.dev.user` - Secrets (prefixed with `SECRET_`)

Customize environments by editing `m365agentstoolkit*.yml` actions and corresponding env files.

## Preview

Open app in Teams or Copilot:

- Teams: `https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}`
- Copilot: `https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${{M365_APP_ID}}?auth=2&developerMode=Basic`

# ATK Project Context Resolution

Resolve config values only when missing. If a value is already known in the session, reuse it.

## Step 1: Detect ATK Project

If `m365agentstoolkit*.yml` exists in the current folder, treat it as an ATK project and parse configuration.

## Step 2: Resolve Common Configuration

Resolve variables referenced in `m365agentstoolkit*.yml`. Common variables:
AZURE_OPENAI_API_KEY
AZURE_OPENAI_ENDPOINT
AZURE_OPENAI_DEPLOYMENT_NAME

## Step 3: Collect Missing Values

If required values are missing, ask the user for only the missing ones.

## ATK project config files
Refer to [config-files](./references/config-files.md) for full config-file details.