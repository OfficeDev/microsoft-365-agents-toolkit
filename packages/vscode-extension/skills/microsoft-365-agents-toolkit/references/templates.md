# Agent Templates Reference

## Contents
- Declarative Agents (creating, options)
- Custom Engine Agents (creating, languages)
- Teams Agents (creating, languages)
- Other Templates (bot, tab, message extension)
- Best Practices (language matching)

## Declarative Agents (Copilot Extensions)

### Creating a Declarative Agent

```bash
# Basic declarative agent (no backend service needed)
atk new -c declarative-agent -n myagent -i false

# Declarative agent with new API plugin (creates backend)
atk new -c declarative-agent-action -l typescript -n myagent -i false

# Declarative agent with existing OpenAPI spec (requires -a and -o with operation IDs)
# First inspect the OpenAPI spec to find operation IDs, then pass them:
atk new -c declarative-agent-action-from-existing-api -n myagent -a <openapi-spec-url-or-path> -o "GET /repairs" -o "POST /repairs" -i false

# Declarative agent with MCP Server
atk new -c declarative-agent-with-action-from-mcp -n myagent -i false
```

**Important Notes:**
- Basic declarative agents (`declarative-agent`) do NOT require a programming language
- `declarative-agent-action`: Use `-l typescript/javascript/csharp` (creates new backend API)
- `declarative-agent-action-from-existing-api`: Requires `-a` (OpenAPI spec) and `-o` (operation IDs from the spec, e.g., `"GET /repairs"`)

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

## Other Teams Templates

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

## Best Practices

### Before Creating a Project

1. **Use non-interactive mode** - Always use `-i false` for scripted creation

2. **Match language to capability**:
   - Basic declarative agents (`declarative-agent`): NO language flag needed
   - API plugin agents (`declarative-agent-action`): `-l typescript/javascript/csharp`
   - Custom Engine agents: `-l typescript/javascript/python`
   - Teams agents: `-l typescript/javascript/csharp/python`
