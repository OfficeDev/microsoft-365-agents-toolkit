# Create Project

Scaffold a new Microsoft 365 agent or Teams app from an ATK template.

## Template Selection Guide

### Declarative Agents

| User Wants | Lifecycle Route | Capability or Command |
| --- | --- | --- |
| Extend M365 Copilot with custom instructions | WIQD | `wiqd agent create` |
| Declarative Agent with new API backend source code | ATK | `declarative-agent-action` |
| Declarative Agent with new API backend and Bearer Token | ATK | `declarative-agent-action-bearer` |
| Declarative Agent with new API backend and OAuth | ATK | `declarative-agent-action-oauth` |
| Declarative Agent with existing OpenAPI spec | WIQD | `wiqd agent create` then `wiqd agent add action` |
| Connect existing MCP Server to Copilot | WIQD | `wiqd agent create` then `wiqd agent add action` |
| Declarative Agent using an existing Copilot Connector connection | WIQD | `wiqd agent create` |
| Declarative Agent with a new project-owned Copilot Connector backend | ATK | `declarative-agent-with-graph-connector` |
| Declarative Agent for MetaOS | ATK | `declarative-agent-meta-os-new-project` |
| Declarative Agent from TypeSpec | ATK | `declarative-agent-typespec` |

### Other Agents and Apps

| User Wants | Capability |
|------------|------------|
| Agent with custom LLM (Azure OpenAI, etc.) | `basic-custom-engine-agent` |
| Weather forecast agent | `weather-agent` |
| Agent using Azure AI Foundry | `foundry-agent-to-m365` |
| Teams chatbot with AI | `teams-agent` |
| Teams bot with RAG/knowledge base | `teams-agent-rag-customize` |
| Teams Agent with Azure AI Search | `teams-agent-rag-azure-ai-search` |
| Teams Agent with Custom API | `teams-agent-rag-custom-api` |
| Teams Collaborator Agent | `teams-collaborator-agent` |
| Simple Teams echo bot | `bot` |
| Teams tab app | `tab` |
| Teams message extension | `message-extension` |
| Copilot Connector | `copilot-connector` |
| Office task pane add-in (Word/Excel/PowerPoint/Outlook) | `office-addin-wxpo-taskpane` |
| Excel custom functions (shared runtime) | `office-addin-excel-cfshortcut` |
| Excel custom functions (JavaScript-only runtime) | `office-addin-excel-customfunctions` |
| Office add-in with SSO via Nested App Auth | `office-addin-sso-naa` |

See [../toolkit/templates.md](../toolkit/templates.md) for the complete template catalog with language support and descriptions.

## Creating Projects

For a pure Declarative Agent or a DA that uses an existing OpenAPI, MCP, or Copilot Connector backend, use WIQD and then follow [declarative-agent-lifecycle.md](../toolkit/declarative-agent-lifecycle.md) for actions and lifecycle operations:

```bash
wiqd agent create --name <project-name> --output <parent-folder>
```

For a hybrid DA that needs project-owned API backend source code and deployment, use the ATK from-scratch template. The resulting project uses ATK for its complete lifecycle:

```bash
atk new -c declarative-agent-action -l <typescript|javascript|csharp> -n <project-name> -f <parent-folder> -i false
```

Use `declarative-agent-action-bearer` or `declarative-agent-action-oauth` when the new backend requires that authentication model. Detect the hybrid from its local backend source plus backend deployment actions, not from whether the user invoked WIQD or ATK. Do not switch a hybrid project to WIQD lifecycle commands.

Use `declarative-agent-with-graph-connector` when the project must scaffold and deploy a new Copilot Connector ingestion backend. That generated backend makes the project hybrid; merely referencing an existing Connector connection does not.

For all other ATK entries, create templates in the current directory with this flow:

```bash
# 1) Scaffold into a temporary parent folder
atk new -c <template-id> -n <project-name> -f /tmp -l <language> -i false

# 2) Move generated files from the scaffold subfolder into current directory
mv /tmp/<project-name>/. .

# 3) Remove the empty scaffold folder
rmdir /tmp/<project-name>
```

Common examples:

```bash
# Custom Engine Agent
atk new -c basic-custom-engine-agent -l typescript -n my-cea -f /tmp -i false

# Teams Agent with RAG
atk new -c teams-agent-rag-customize -l typescript -n my-rag-agent -f /tmp -i false

# Office task pane add-in (all Office apps by default)
atk new -c office-addin-wxpo-taskpane -l typescript -n my-addin -f /tmp -i false

# Office task pane add-in for specific apps only (comma-separated, no spaces)
atk new -c office-addin-wxpo-taskpane -l typescript -n my-addin --office-addin-hosts word,excel -f /tmp -i false

# Excel custom functions (JavaScript-only runtime)
atk new -c office-addin-excel-customfunctions -l typescript -n my-addin -f /tmp -i false

# Office add-in with SSO via Nested App Auth
atk new -c office-addin-sso-naa -l typescript -n my-addin -f /tmp -i false
```

PowerShell equivalent:

```powershell
# 1) Scaffold into temporary folder
atk new -c <template-id> -n <project-name> -f $env:TEMP -l <language> -i false

# 2) Move files into current directory
Move-Item "$env:TEMP\<project-name>\*" .
Move-Item "$env:TEMP\<project-name>\.*" . -ErrorAction SilentlyContinue

# 3) Remove scaffold folder
Remove-Item "$env:TEMP\<project-name>" -Force
```

## Creating from Samples

```bash
atk new sample <sample-id>
```

To place sample files in current directory, scaffold first and then move files from the sample output folder into `.` using the same move pattern as above.

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

## Notes

- Office add-in capabilities (`office-addin-*`) are TypeScript only; `office-addin-wxpo-taskpane` accepts `--office-addin-hosts word,excel` (comma-separated) to limit which Office apps it supports, defaulting to all four
- Always use `-i false` for non-interactive scripted creation
- `atk new` can take several minutes — wait for completion (timeout 120000ms+)
- If template/sample already matches the requirement, do not run dependency install by default; continue only when user asks for next steps

## After Scaffolding

Once the project is created:
- For a DA without project-owned backend deployment → continue with [declarative-agent-lifecycle.md](../toolkit/declarative-agent-lifecycle.md)
- For a hybrid DA with project-owned backend source and deployment → keep its complete lifecycle on ATK; use [../provision-deploy/provision-deploy.md](../provision-deploy/provision-deploy.md)
- For other projects, to test locally → see [../test-playground/test-playground.md](../test-playground/test-playground.md)
- To understand project files → see [../toolkit/manifest-and-yaml.md](../toolkit/manifest-and-yaml.md)

## Expert Deep Dives

> **Applies to: code-based Teams bots/agents only** (templates: `bot`, `teams-agent*`, `basic-custom-engine-agent`, `weather-agent`, `coffee-agent`, `bot-sso`, `msgext-*`, `tab*`).
>
> Does **not** apply to pure DAs, existing-API/MCP actions, DAs that only reference existing Copilot Connector connections, or other templates without application source code. Hybrid DA projects contain project-owned backend source and deployment actions; keep their complete lifecycle on ATK and follow the generated project instructions.

For deeper guidance on what `atk new` produces and how to extend it, consult the Teams expert micro-files:

| Topic | Expert |
|---|---|
| Project file layout, `package.json`, `tsconfig.json`, npm scripts, `appPackage/` | [../experts/teams/project.scaffold-files-ts.md](../experts/teams/project.scaffold-files-ts.md) |
| `App` constructor, plugins, credentials, runtime initialization | [../experts/teams/runtime.app-init-ts.md](../experts/teams/runtime.app-init-ts.md) |
| Teams app manifest schema, scopes, bots/composeExtensions/staticTabs | [../experts/teams/runtime.manifest-ts.md](../experts/teams/runtime.manifest-ts.md) |
| Routing handlers (`app.on('message')`, activity types, invokes) | [../experts/teams/runtime.routing-handlers-ts.md](../experts/teams/runtime.routing-handlers-ts.md) |
