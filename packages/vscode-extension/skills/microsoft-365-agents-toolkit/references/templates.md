# Agent Templates Reference

## Contents
- CLI Capabilities (all atk new -c options)
- Declarative Agents (creating, options)
- Custom Engine Agents (creating, languages)
- Teams Agents (creating, languages)
- Other Templates (bot, tab, message extension)
- Best Practices (language matching)
- Template Selection Guide

## CLI Capabilities (atk new -c)

Use `atk new -c <capability>` to create projects. Available capabilities:

| Capability | Description |
|------------|-------------|
| `copilot-gpt-basic` | Declarative Agent |
| `api-plugin-from-scratch` | Declarative Agent with new API (creates backend service) |
| `api-plugin-from-scratch-bearer` | Declarative Agent with new API (Bearer Token auth) |
| `api-plugin-from-scratch-oauth` | Declarative Agent with new API (OAuth auth) |
| `api-plugin-from-existing-api` | Declarative Agent with existing OpenAPI spec |
| `declarative-agent-with-action-from-mcp` | Declarative Agent with Action from MCP Server |
| `declarative-agent-with-graph-connector` | Declarative Agent with Copilot Connector |
| `declarative-agent-meta-os-new-project` | Declarative Agent for MetaOS (New Project) |
| `declarative-agent-meta-os-upgrade-project` | Declarative Agent for MetaOS (Upgrade Project) |
| `declarative-agent-typespec` | Declarative Agent from TypeSpec |
| `basic-custom-engine-agent` | Basic Custom Engine Agent |
| `weather-agent` | Weather Agent |
| `foundry-agent` | Foundry Agent |
| `graph-connector` | Copilot Connector |
| `custom-copilot-basic` | General Teams Agent |
| `custom-copilot-rag-customize` | Teams Agent with Data from Customized Source |
| `custom-copilot-rag-azure-ai-search` | Teams Agent with Data from Azure AI Search |
| `custom-copilot-rag-custom-api` | Teams Agent with Data from Custom API using OpenAPI Spec |
| `teams-collaborator-agent` | Teams Collaborator Agent |
| `non-sso-tab` | Tab |
| `default-bot` | Simple Bot |
| `default-message-extension` | Message Extension |
| `office-addin-outlook-taskpane` | Outlook Task Pane Add-in |
| `office-addin-wxpo-taskpane` | Office Task Pane Add-in |
| `office-addin-excel-cfshortcut` | Excel Custom Functions |
| `office-addin-config` | Office Add-in Common Configuration |

## Declarative Agents (Copilot Extensions)

### Creating a Declarative Agent

```bash
# Basic declarative agent (no backend service needed)
atk new -c copilot-gpt-basic -n myagent -i false

# Declarative agent with new API plugin (creates backend)
atk new -c api-plugin-from-scratch -l typescript -n myagent -i false

# Declarative agent with existing OpenAPI spec (requires -a and -o with operation IDs)
# First inspect the OpenAPI spec to find operation IDs, then pass them:
atk new -c api-plugin-from-existing-api -n myagent -a <openapi-spec-url-or-path> -o "GET /repairs" -o "POST /repairs" -i false

# Declarative agent with MCP Server
atk new -c declarative-agent-with-action-from-mcp -n myagent -i false
```

**Important Notes:**
- Basic declarative agents (`copilot-gpt-basic`) do NOT require a programming language
- `api-plugin-from-scratch`: Use `-l typescript/javascript/csharp` (creates new backend API)
- `api-plugin-from-existing-api`: Requires `-a` (OpenAPI spec) and `-o` (operation IDs from the spec, e.g., `"GET /repairs"`)

### Declarative Agent Options

| Option | Values | Description |
|--------|--------|-------------|
| `--openapi-spec-location -a` | file path or URL | **Required for existing API**: OpenAPI spec location |
| `--api-operation -o` | operation IDs (e.g., `"GET /path"`) | **Required for existing API**: Actual operation IDs from OpenAPI spec. Use multiple `-o` for multiple operations |
| `--api-auth` | `none`, `api-key`, `bearer-token`, `oauth` | API authentication type |

## Custom Engine Agents (M365 SDK-based)

### Creating a Custom Engine Agent

```bash
# Basic custom engine agent
atk new -c basic-custom-engine-agent -l typescript -n myagent -i false

# Weather agent sample
atk new -c weather-agent -l typescript -n myagent -i false
```

| Capability | Languages | Description |
|------------|-----------|-------------|
| `basic-custom-engine-agent` | typescript, javascript, python | Basic agent with M365 SDK and LLM |
| `weather-agent` | typescript, javascript, csharp | Weather forecast agent with LangChain |

## Teams Agents (Teams AI Library)

### Creating a Teams Agent

```bash
# Basic Teams chatbot
atk new -c custom-copilot-basic -l typescript -n mybot -i false

# Teams Agent with RAG (custom data source)
atk new -c custom-copilot-rag-customize -l typescript -n mybot -i false

# Teams Agent with Azure AI Search
atk new -c custom-copilot-rag-azure-ai-search -l typescript -n mybot -i false
```

| Capability | Languages | Description |
|------------|-----------|-------------|
| `custom-copilot-basic` | typescript, javascript, csharp, python | General Teams Agent |
| `custom-copilot-rag-customize` | typescript, javascript, csharp, python | Teams Agent with Customized Data Source |
| `custom-copilot-rag-azure-ai-search` | typescript, javascript, csharp, python | Teams Agent with Azure AI Search |
| `custom-copilot-rag-custom-api` | typescript, javascript, csharp, python | Teams Agent with Custom API |
| `teams-collaborator-agent` | typescript, csharp | Teams Collaborator Agent |

## Other Templates

```bash
# Simple Bot
atk new -c default-bot -l typescript -n mybot -i false

# Tab
atk new -c non-sso-tab -l typescript -n mytab -i false

# Message Extension
atk new -c default-message-extension -l typescript -n myme -i false
```

| Capability | Languages | Description |
|------------|-----------|-------------|
| `default-bot` | typescript, javascript, python, csharp | Simple Bot |
| `non-sso-tab` | typescript, csharp | Tab |
| `default-message-extension` | typescript, python, csharp | Message Extension |
| `graph-connector` | typescript, csharp | Copilot Connector |

## Best Practices

### Before Creating a Project

1. **Use non-interactive mode** - Always use `-i false` for scripted creation

2. **Match language to capability**:
   - Basic declarative agents (`copilot-gpt-basic`): NO language flag needed
   - API plugin agents (`api-plugin-from-scratch`): `-l typescript/javascript/csharp`
   - Custom Engine agents: `-l typescript/javascript/python`
   - Teams agents: `-l typescript/javascript/csharp/python`

## Template Selection Guide

**Choose Declarative Agents when:**
- Extending Microsoft 365 Copilot with custom instructions
- Integrating APIs as actions without running custom code
- Need zero-infrastructure deployment

**Choose Custom Engine Agents when:**
- Need custom LLM integration (Azure OpenAI, OpenAI, etc.)
- Require complex multi-turn conversations
- Building with LangChain or other AI frameworks

**Choose Teams Agents when:**
- Building chat bots specifically for Microsoft Teams
- Need RAG (Retrieval Augmented Generation) capabilities
- Require Teams-specific features (channels, meetings, etc.)
