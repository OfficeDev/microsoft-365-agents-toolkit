# Create Project

Scaffold a new Microsoft 365 agent or Teams app from an ATK template.

## Template Selection Guide

| User Wants | Capability |
|------------|------------|
| Extend M365 Copilot with custom instructions | `declarative-agent` |
| Declarative Agent with new API | `declarative-agent-action` |
| Declarative Agent with new API (Bearer Token) | `declarative-agent-action-bearer` |
| Declarative Agent with new API (OAuth) | `declarative-agent-action-oauth` |
| Declarative Agent with existing OpenAPI spec | `declarative-agent-action-from-existing-api` |
| Connect MCP Server to Copilot | `declarative-agent-with-action-from-mcp` |
| Declarative Agent with Copilot Connector | `declarative-agent-with-graph-connector` |
| Declarative Agent for MetaOS | `declarative-agent-meta-os-new-project` |
| Declarative Agent from TypeSpec | `declarative-agent-typespec` |
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

See [../references/templates.md](../references/templates.md) for the complete template catalog with language support and descriptions.

## Creating Projects

**Default project folder:** `~/AgentsToolkitProjects`

```bash
# Declarative Agent (no backend, no -l flag needed)
atk new -c declarative-agent -n my-agent -f ~/AgentsToolkitProjects -i false

# Declarative Agent with new API
atk new -c declarative-agent-action -l typescript -n my-api-agent -f ~/AgentsToolkitProjects -i false

# Declarative Agent with existing OpenAPI spec (requires -a and -o with operation IDs)
atk new -c declarative-agent-action-from-existing-api -n my-agent -a <openapi-spec-url-or-path> -o "GET /repairs" -o "POST /repairs" -f ~/AgentsToolkitProjects -i false

# Custom Engine Agent
atk new -c basic-custom-engine-agent -l typescript -n my-cea -f ~/AgentsToolkitProjects -i false

# Teams Agent with RAG
atk new -c teams-agent-rag-customize -l typescript -n my-rag-agent -f ~/AgentsToolkitProjects -i false
```

## Creating from Samples

```bash
atk new sample <sample-id>
```

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

- `declarative-agent` does NOT require `-l` language flag
- `declarative-agent-action-from-existing-api` requires `-a` (OpenAPI spec) and `-o` (operation IDs like `"GET /path"`)
- Always use `-i false` for non-interactive scripted creation
- `atk new` can take several minutes — wait for completion (timeout 120000ms+)

## After Scaffolding

Once the project is created:
- To test locally → see [../test-playground/test-playground.md](../test-playground/test-playground.md)
- To understand project files → see [../references/manifest-and-yaml.md](../references/manifest-and-yaml.md)
