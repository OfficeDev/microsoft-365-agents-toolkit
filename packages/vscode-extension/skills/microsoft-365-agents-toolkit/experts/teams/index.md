# teams-router

## purpose

Route Teams bot/agent tasks to the minimal set of micro-expert files. Read only the clusters that match the user's request. For multi-area requests, combine files from each relevant cluster.

## task clusters

### Runtime: App Initialisation
When: creating a new bot, setting up `app.ts`, configuring Teams AI library, bot entry point
Read:
- `runtime.app-init-ts.md`
- `project.scaffold-files-ts.md` (only if scaffolding a new project)
- `dev.debug-test-ts.md` (only if setting up local dev environment)

### Runtime: Routing & Handlers
When: adding activity handlers, message routing, turn handling, middleware
Read:
- `runtime.routing-handlers-ts.md`
Depends on: `runtime.app-init-ts.md` (App must be initialized before registering handlers)

### Runtime: Manifest
When: Teams app manifest, `manifest.json`, app registration, scopes, permissions
Read:
- `runtime.manifest-ts.md`
- `project.scaffold-files-ts.md` (for appPackage directory structure)

### Runtime: Proactive Messaging
When: sending messages outside a conversation turn, proactive notifications, scheduled messages
Read:
- `runtime.proactive-messaging-ts.md`
- `state.storage-patterns-ts.md` (for persisting conversation IDs across restarts)
Depends on: `runtime.app-init-ts.md` (requires App credentials for proactive sends)

### UI: Adaptive Cards
When: building cards, card actions, card templates, card rendering
Read:
- `ui.adaptive-cards-ts.md`
- `runtime.routing-handlers-ts.md` (for `card.action` handler registration)
- `ui.dialogs-task-modules-ts.md` (only if cards open dialogs/task modules)

### UI: Dialogs & Task Modules
When: dialogs, task modules, modal popups, multi-step forms
Read:
- `ui.dialogs-task-modules-ts.md`
- `ui.adaptive-cards-ts.md` (task modules render Adaptive Cards)
- `runtime.routing-handlers-ts.md` (for `dialog.open`/`dialog.submit` routes)

### UI: Message Extensions
When: message extensions, search commands, action commands, link unfurling
Read:
- `ui.message-extensions-ts.md`
- `runtime.manifest-ts.md` (composeExtensions must be declared in manifest)
- `ui.adaptive-cards-ts.md` (extensions return card attachments)

### Auth & Graph
When: SSO, OAuth, authentication, Graph API, user profile, app-only token
Read:
- `auth.oauth-sso-ts.md`
- `graph.usergraph-appgraph-ts.md`
Depends on: `runtime.app-init-ts.md` (oauth config set in App constructor)

### State & Storage
When: conversation state, user state, storage, persistence, memory patterns
Read:
- `state.storage-patterns-ts.md`
- `ai.memory-localmemory-ts.md` (only if combining state with AI conversation history)
Depends on: `runtime.app-init-ts.md` (storage passed to App constructor)

### AI: ChatPrompt & Model
When: ChatPrompt setup, prompt templates, model configuration, OpenAI/Azure OpenAI
Read:
- `ai.chatprompt-basics-ts.md`
- `ai.model-setup-ts.md`
Depends on: `runtime.app-init-ts.md` (ChatPrompt used inside message handlers)

### AI: Function Calling
When: defining tools/functions for the LLM, JSON schema, function design, function implementation
Read:
- `ai.function-calling-design-ts.md`
- `ai.function-calling-implementation-ts.md`
Depends on: `ai.chatprompt-basics-ts.md` (functions chain off ChatPrompt), `ai.model-setup-ts.md`

### AI: RAG & Retrieval
When: retrieval-augmented generation, embeddings, vector stores, knowledge base
Read:
- `ai.rag-retrieval-ts.md`
- `ai.rag-vectorstores-ts.md`
- `ai.citations-feedback-ts.md` (for annotating RAG responses with source citations)
Depends on: `ai.function-calling-implementation-ts.md` (RAG uses search as a function), `ai.chatprompt-basics-ts.md`

### AI: Streaming & Citations
When: streaming responses, SSE, citation rendering, feedback loops, thumbs up/down
Read:
- `ai.streaming-ts.md`
- `ai.citations-feedback-ts.md`
Depends on: `ai.chatprompt-basics-ts.md` (streaming wraps prompt.send()), `runtime.routing-handlers-ts.md` (for ctx.stream)

### AI: Memory
When: LocalMemory, conversation memory, chat history, context window management
Read:
- `ai.memory-localmemory-ts.md`
- `state.storage-patterns-ts.md` (only if persisting memory across restarts)
Depends on: `ai.chatprompt-basics-ts.md` (memory passed to ChatPrompt constructor)

### MCP: Model Context Protocol
When: MCP server, MCP client, exposing tools via MCP, MCP security
Read:
- `mcp.server-basics-ts.md`
- `mcp.client-basics-ts.md`
- `mcp.expose-chatprompt-tools-ts.md`
- `mcp.security-ts.md` (only if security/auth questions)
Depends on: `runtime.app-init-ts.md` (McpPlugin added to App plugins). MCP client also depends on `ai.chatprompt-basics-ts.md` (McpClientPlugin is a ChatPrompt plugin). `mcp.expose-chatprompt-tools-ts.md` depends on `ai.function-calling-implementation-ts.md` (bridges prompt functions to MCP tools).

### A2A: Agent-to-Agent
When: A2A protocol, agent orchestration, multi-agent, agent discovery
Read:
- `a2a.server-basics-ts.md`
- `a2a.client-basics-ts.md`
- `a2a.orchestrator-patterns-ts.md` (only if orchestrating multiple agents)
Depends on: `runtime.app-init-ts.md` (A2APlugin added to App plugins). A2A client also depends on `ai.chatprompt-basics-ts.md` (A2AClientPlugin is a ChatPrompt plugin).

### Compatibility: BotBuilder Interop
When: mixing BotBuilder SDK with Teams AI, legacy bot code, adapter patterns
Read:
- `compat.botbuilder-interop-ts.md`
- `runtime.app-init-ts.md` (for understanding the target SDK v2 patterns)
- `runtime.routing-handlers-ts.md` (for mapping TeamsActivityHandler to SDK v2 routes)

### Dev: Debug & Test
When: debugging, testing, ATK, Agents Toolkit, local tunnel, dev tools, unit tests
Read:
- `dev.debug-test-ts.md`
- `project.scaffold-files-ts.md` (for npm scripts and build verification)

### Scaffolding
When: new project, file structure, folder layout, boilerplate, starter template
Read:
- `project.scaffold-files-ts.md`
- `runtime.app-init-ts.md`
- `runtime.manifest-ts.md` (for appPackage/manifest.json setup)

### Teams SDK for Python
When: Python, `microsoft_teams`, `microsoft_teams.apps`, `microsoft_teams.ai`, `ActivityContext`, `@app.on_message`, `@app.on_message_pattern`, `ChatPrompt`, `OpenAICompletionsAIModel`, Pydantic, FastAPI, Python Teams bot
Read:
- `teams-python.md`
Note: All TS experts provide architectural patterns. This expert provides Python API mappings. Load the relevant TS expert for concepts, then this expert for Python translation.

### Teams SDK for .NET (C#)
When: C#, .NET, `Microsoft.Teams.Apps`, `Microsoft.Teams.AI`, `AddTeams()`, `UseTeams()`, `IContext<TActivity>`, `OnMessage`, `OnAdaptiveCardAction`, `OpenAIChatPrompt`, `[Prompt]`, `[Function]`, ASP.NET Core, NuGet
Read:
- `teams-dotnet.md`
Note: C# has SDK support for Teams only (Tier 3). For the Slack side, route to `../bridge/rest-only-integration-ts.md`.

### Toolkit: ATK CLI, Environments, Playground, Publishing
When: `m365agents.yml`, `atk` CLI, `atk provision`, `atk deploy`, `atk share`, `atk new`, `atk auth`, `atk doctor`, lifecycle hooks, CI/CD pipeline, built-in actions, env files, environment variables, `${{VAR}}`, `SECRET_` prefix, multi-environment, Agents Playground, `agentsplayground`, local testing, publish to org, Teams Store, Partner Center
Read:
- Parent skill references at [../../references/](../../references/) — covers ATK CLI commands, environment config, local debugging, templates, and publishing
- `runtime.manifest-ts.md` (for manifest.json v1.25 schema and packaging)
Note: ATK toolkit details are maintained in the parent microsoft-365-agents-toolkit skill to avoid duplication.

## cross-platform bridging

If the developer wants to **add Slack support** to an existing Teams bot, route to `../bridge/index.md` for cross-platform bridging experts. The bridge domain covers Teams↔Slack feature mapping, UI conversion, identity bridging, and infrastructure migration. For ATK toolkit patterns (environments, lifecycle, playground), see the parent skill's [../../references/](../../references/) folder.

## combining rule

If a request spans multiple clusters (e.g., "add a function-calling tool that returns an Adaptive Card"), read files from **every** matching cluster. Avoid duplicates.

## file inventory

`a2a.client-basics-ts.md` | `a2a.orchestrator-patterns-ts.md` | `a2a.server-basics-ts.md` | `ai.chatprompt-basics-ts.md` | `ai.citations-feedback-ts.md` | `ai.function-calling-design-ts.md` | `ai.function-calling-implementation-ts.md` | `ai.memory-localmemory-ts.md` | `ai.model-setup-ts.md` | `ai.rag-retrieval-ts.md` | `ai.rag-vectorstores-ts.md` | `ai.streaming-ts.md` | `auth.oauth-sso-ts.md` | `compat.botbuilder-interop-ts.md` | `dev.debug-test-ts.md` | `graph.usergraph-appgraph-ts.md` | `mcp.client-basics-ts.md` | `mcp.expose-chatprompt-tools-ts.md` | `mcp.security-ts.md` | `mcp.server-basics-ts.md` | `project.scaffold-files-ts.md` | `runtime.app-init-ts.md` | `runtime.manifest-ts.md` | `runtime.proactive-messaging-ts.md` | `runtime.routing-handlers-ts.md` | `state.storage-patterns-ts.md` | `teams-dotnet.md` | `teams-python.md` | `ui.adaptive-cards-ts.md` | `ui.dialogs-task-modules-ts.md` | `ui.message-extensions-ts.md`

<!-- ATK toolkit experts (lifecycle-cli, environments, playground, publish) moved to parent skill references/ -->
