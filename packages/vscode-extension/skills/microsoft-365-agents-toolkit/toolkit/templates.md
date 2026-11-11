# Agent Templates Reference

Delegate every DA to the WIQD plugin except an explicitly requested TypeSpec DA. The plugin coordinates Azure and Entra skills when the DA needs a new backend.

## Contents
- CLI Capabilities (ATK-routed options)
- Declarative Agents (creating, options)
- Custom Engine Agents (creating, languages)
- Teams Agents (creating, languages)
- Other Templates (bot, tab, message extension)
- Best Practices (language matching)
- Template Selection Guide

## CLI Capabilities (atk new -c)

For entries routed to ATK, use `atk new -c <capability>`. Among DAs, only `declarative-agent-typespec` is routed to ATK; the other DA entries below are retained for legacy recognition, not direct selection.

| Capability | Description |
|------------|-------------|
| `declarative-agent-action` | Declarative Agent with project-owned Action from Scratch |
| `declarative-agent-action-bearer` | Declarative Agent with project-owned Action from Scratch (Bearer Token) |
| `declarative-agent-action-oauth` | Declarative Agent with project-owned Action from Scratch (OAuth) |
| `declarative-agent-with-graph-connector` | DA with a new project-owned Copilot Connector backend |
| `declarative-agent-typespec` | Declarative Agent from TypeSpec |
| `basic-custom-engine-agent` | Basic Custom Engine Agent |
| `weather-agent` | Weather Agent |
| `foundry-agent-to-m365` | Foundry Agent to M365 |
| `copilot-connector` | Copilot Connector |
| `teams-agent` | General Teams Agent |
| `teams-agent-rag-customize` | Teams Agent with Data from Customized Source |
| `teams-agent-rag-azure-ai-search` | Teams Agent with Data from Azure AI Search |
| `teams-agent-rag-custom-api` | Teams Agent with Data from Custom API using OpenAPI Spec |
| `teams-collaborator-agent` | Teams Collaborator Agent |
| `tab` | Tab |
| `bot` | Simple Bot |
| `message-extension` | Message Extension |
| `office-addin-outlook-taskpane` | Outlook Task Pane Add-in |
| `office-addin-wxpo-taskpane` | Office Task Pane Add-in |
| `office-addin-excel-cfshortcut` | Excel Custom Functions |
| `office-addin-excel-customfunctions` | Excel Custom Functions (JavaScript-only Runtime) |
| `office-addin-sso-naa` | Single Sign-On with Nested App Auth |
| `office-addin-config` | Office Add-in Common Configuration |

## Declarative Agents (Copilot Extensions)

### Creating a Declarative Agent

Invoke the `wiqd` skill for every DA except TypeSpec and let the plugin own creation, actions, and the remaining lifecycle. This includes new Azure-hosted API backends, existing OpenAPI or MCP services, and Copilot Connector scenarios. The plugin coordinates Azure and Entra skills when external backend or identity work is required. Follow [declarative-agent-lifecycle.md](declarative-agent-lifecycle.md) for the routing boundary.

Use `atk new -c declarative-agent-typespec` only when the user explicitly requests a TypeSpec DA.

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
atk new -c teams-agent -l typescript -n mybot -i false

# Teams Agent with RAG (custom data source)
atk new -c teams-agent-rag-customize -l typescript -n mybot -i false

# Teams Agent with Azure AI Search
atk new -c teams-agent-rag-azure-ai-search -l typescript -n mybot -i false
```

| Capability | Languages | Description |
|------------|-----------|-------------|
| `teams-agent` | typescript, javascript, csharp, python | General Teams Agent |
| `teams-agent-rag-customize` | typescript, javascript, csharp, python | Teams Agent with Customized Data Source |
| `teams-agent-rag-azure-ai-search` | typescript, javascript, csharp, python | Teams Agent with Azure AI Search |
| `teams-agent-rag-custom-api` | typescript, javascript, csharp, python | Teams Agent with Custom API |
| `teams-collaborator-agent` | typescript, csharp | Teams Collaborator Agent |

## Other Templates

```bash
# Simple Bot
atk new -c bot -l typescript -n mybot -i false

# Tab
atk new -c tab -l typescript -n mytab -i false

# Message Extension
atk new -c message-extension -l typescript -n myme -i false
```

| Capability | Languages | Description |
|------------|-----------|-------------|
| `bot` | typescript, javascript, python, csharp | Simple Bot |
| `tab` | typescript, csharp | Tab |
| `message-extension` | typescript, python, csharp | Message Extension |
| `copilot-connector` | typescript, csharp | Copilot Connector |

## Office Add-ins

Office add-ins extend Office applications (Word, Excel, PowerPoint, Outlook) with custom task panes and functions. All office add-in capabilities are TypeScript only.

```bash
# Task pane add-in for Word, Excel, PowerPoint, and Outlook (all hosts by default)
atk new -c office-addin-wxpo-taskpane -l typescript -n myaddin -i false

# Task pane supporting only specific Office apps (comma-separated, no spaces)
atk new -c office-addin-wxpo-taskpane -l typescript -n myaddin --office-addin-hosts word,excel -i false

# Excel custom functions with keyboard shortcuts
atk new -c office-addin-excel-cfshortcut -l typescript -n myaddin -i false

# Excel custom functions with a JavaScript-only runtime
atk new -c office-addin-excel-customfunctions -l typescript -n myaddin -i false

# Task pane add-in with single sign-on via Nested App Auth
atk new -c office-addin-sso-naa -l typescript -n myaddin -i false
```

| Capability | Languages | Description |
|------------|-----------|-------------|
| `office-addin-wxpo-taskpane` | typescript | Task pane add-in for Word, Excel, PowerPoint, and/or Outlook |
| `office-addin-excel-cfshortcut` | typescript | Excel custom functions with keyboard shortcuts (shared runtime) |
| `office-addin-excel-customfunctions` | typescript | Excel custom functions using a JavaScript-only runtime |
| `office-addin-sso-naa` | typescript | Task pane add-in with single sign-on via Nested App Auth |
| `office-addin-config` | typescript | Common configuration for Office add-ins |

### Office Add-in Options

| Option | Values | Description |
|--------|--------|-------------|
| `--office-addin-hosts` | `word`, `powerpoint`, `outlook`, `excel` (comma-separated) | Office apps the task pane supports. Only applies to `office-addin-wxpo-taskpane`. Optional; defaults to all four hosts. Source files and debug configs for unselected hosts are pruned. |

## Best Practices

### Before Creating a Project

1. **Use non-interactive mode** - Always use `-i false` for scripted creation

2. **Match language to capability**:
   - API plugin agents (`declarative-agent-action`): `-l typescript/javascript/csharp`
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
