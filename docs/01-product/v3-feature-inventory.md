# v3 feature inventory

> **Purpose.** A complete enumeration of what the v3 toolkit can build today, derived from the shipped templates, the [`.dev/features.json`](../../.dev/features.json) feature registry, the [`vscode-extension`](../../packages/vscode-extension/) `package.json` contributions, and the v3 CLI command surface.
>
> **v4 design status.** This page **is** an allowed input to v4 design. It describes user-visible product capabilities — *what* gets built — without committing to *how* v3 implements it internally. v4 may use this list to scope its own template registry. For the corresponding *internal* extractions (FxCore, drivers, generators), see [`_v3-reference/`](../_v3-reference/README.md) — **forbidden** as v4 design input.

This page complements [capabilities-matrix.md](capabilities-matrix.md), which is the high-level cross-engine table. This page goes deeper: per-template description, languages, generated files, lifecycle stages, optional features, and the integration surface.

## Headline product surface

A v3 user can:

- **Create** a new project from one of 29 shipped VS Code templates across 4 languages (TypeScript, JavaScript, Python, C#) and `common` (for code-less templates), plus separate VS-only `templates/vs/csharp/` C# templates.
- **Run / debug** locally with M365 Agents Playground (Test Tool), Teams sideload, or dev tunnels.
- **Provision** Azure resources via declarative IaC (Bicep) for compute, identity, storage, search, SQL, Key Vault, Static Web Apps, and Bot Framework registration.
- **Deploy** code to App Service, Functions, or Static Web Apps via Kudu zip-deploy.
- **Publish** to a tenant app catalog (via Microsoft Graph `/beta/appCatalogs/teamsApps`) or sideload to M365 (V1 classic + V2 declarative agent paths via M365 PackageService).
- **Validate** the manifest and full app package against the Teams Developer Portal.
- **Share** an app with collaborators by tenant scope or specific user list.
- **Manage permissions** — grant/list collaborators on the Teams app and the associated Entra ID app.
- **Add capabilities post-scaffold** — actions, knowledge sources, MCP server actions, OAuth/API key auth configs, sensitivity labels, SPFx web parts.
- **Switch tenants** — both M365 and Azure tenant switching from the toolkit UI.
- **Sideload / uninstall** apps directly via the CLI (`atk install` / `atk uninstall`).
- **Diagnose** the local environment (`atk doctor` checks Node.js, Functions tools, SSL cert, M365 account).

> **Lifecycle stages.** v3 ships **4 driver-run stages** in `m365agents.yml`: `provision`, `deploy`, `publish`, `share`. Plus two *non-driver* stages users perceive as part of the lifecycle: `scaffold` (one-shot at create time) and `run/debug` (handled by `m365agents.local.yml` + dev-tunnel/Test Tool tooling). The capabilities matrix lists these as 4 lifecycles per template (matching `features.json`).

## Capability domains

| Domain | Templates | Languages |
|--------|----------|-----------|
| **Bots** | `default-bot` (Echo Bot) | TS · JS · Python · C# |
| **Custom Engine Agents (CEA)** | `basic-custom-engine-agent`, `weather-agent`, `teams-collaborator-agent`, `teams-agent-with-data-custom-api-v2` | TS · JS · Python (collaborator: TS only; teams-agent-with-data: Python only) |
| **AI Agents (Teams AI Library)** | `custom-copilot-basic` (chat), `custom-copilot-rag-azure-ai-search`, `custom-copilot-rag-custom-api`, `custom-copilot-rag-customize` | TS · JS · Python |
| **Tabs** | `basic-tab` (registered in `features.json` as `non-sso-tab`) | TS · C# |
| **Message Extensions** | `message-extension-v2`, `message-extension-with-existing-api` | TS · Python (v2); common (existing-api) |
| **Declarative Agents (DA) — code-less** | `declarative-agent-basic`, `declarative-agent-meta-os-upgrade-project`, `declarative-agent-typespec`, `declarative-agent-with-action-from-existing-api`, `declarative-agent-with-action-from-mcp` | common |
| **Declarative Agents (DA) — with backend** | `declarative-agent-with-action-from-scratch` (no auth), `declarative-agent-with-action-from-scratch-bearer`, `declarative-agent-with-action-from-scratch-oauth` | TS · JS |
| **Declarative Agents (DA) — Foundry** | `foundry-agent-to-m365` | TS |
| **Declarative Agents (DA) — MetaOS** | `declarative-agent-meta-os-new-project` | common |
| **Graph Connectors** | `graph-connector` | TS |
| **Office Add-ins (taskpane)** | `office-addin-config`, `office-addin-excel-cfshortcut`, `office-addin-outlook-taskpane`, `office-addin-wxpo-taskpane`, `office-addin` (common), `office-xml-addin-common` | TS · common |
| **Samples (separate from templates)** | many — listed via `atk list samples` | varies |

## Template-level details (shipped)

> **Languages column reflects the `features.json` `languages` field where the template has a feature-registry entry, otherwise the directories present in `templates/vsc/`. The two sources can diverge — e.g. `weather-agent` ships a Python folder (`templates/vsc/python/weather-agent/`) but `features.json` only lists TS+JS, meaning E2E coverage targets only TS/JS even though Python scaffolds work.**

| Template | What it builds | Languages | Has Azure infra | Notes |
|----------|----------------|-----------|-----------------|-------|
| `default-bot` | Echo Bot on Microsoft Agents SDK | TS · JS · Python · C# | Yes (Bot/AppService+Identity) | Foundational bot scaffold |
| `basic-tab` (feat-registry id: `non-sso-tab`) | Personal/configurable tab on a Web App | TS · C# | Yes (Tab/AppService) | No Entra ID app required by default |
| `basic-custom-engine-agent` | Custom-engine agent (Microsoft Agents SDK) | TS · JS · Python | Yes (Bot/AppService+Identity) | Echo-style starting point |
| `weather-agent` | Function-calling weather CEA | TS · JS (E2E); Python folder ships but not in feat registry | Yes (Bot/AppService+Identity) | Demonstrates tool use |
| `teams-collaborator-agent` | Multi-step collaboration agent | TS only | Yes (Bot/AppService+SQL) | Adds Azure SQL DB + AOAI integration |
| `teams-agent-with-data-custom-api-v2` | Python CEA with custom-API RAG | Python only | Yes (Bot/AppService+Identity) | Not in feat registry; lives in `templates/vsc/python/` only |
| `custom-copilot-basic` | AI chat bot (Teams AI Library) | TS · JS · Python | Yes (Bot/AppService+Identity) | LLM provider configurable (OpenAI / AzureOpenAI / Foundry) |
| `custom-copilot-rag-azure-ai-search` | RAG agent backed by Azure AI Search | TS · JS · Python | Yes (Bot/AppService+Identity+Search) | |
| `custom-copilot-rag-custom-api` | RAG agent backed by user-supplied API | TS · JS · Python | Yes (Bot/AppService+Identity) | API spec input required |
| `custom-copilot-rag-customize` | RAG agent with customised data layer | TS · JS · Python | Yes (Bot/AppService+Identity) | |
| `message-extension-v2` | Search-based message extension | TS · Python | Yes (Bot/AppService+Identity) | |
| `message-extension-with-existing-api` | ME backed by an existing OpenAPI spec | common | No (no backend infra; spec drives it) | Interactive `apiSpecPath` input |
| `declarative-agent-basic` | Pure DA (no code, no backend) | common | No | Just `appPackage/declarativeAgent.json` |
| `declarative-agent-with-action-from-existing-api` | DA + API plugin from an existing OpenAPI spec | common | No (uses external API) | Interactive spec input |
| `declarative-agent-with-action-from-scratch` | DA + new API on Functions (no auth) | TS · JS | Yes (Functions/API) | |
| `declarative-agent-with-action-from-scratch-bearer` | DA + new API on Functions (bearer/API key) | TS · JS | Yes (Functions/API) | |
| `declarative-agent-with-action-from-scratch-oauth` | DA + new API on App Service (OAuth) | TS · JS | Yes (Bot/AppService+Identity for OAuth) | |
| `declarative-agent-with-action-from-mcp` | DA + action backed by an MCP server | common | No (uses external MCP server) | Behind feature flag |
| `declarative-agent-typespec` | DA + API plugin defined via TypeSpec | common | No | Compilation step required (`typeSpec/compile` driver) |
| `declarative-agent-meta-os-new-project` | New MetaOS-shaped DA project | common | Yes (Static Web App) | |
| `declarative-agent-meta-os-upgrade-project` | Convert existing MetaOS add-in into DA | common | No | Inline `scaffoldFn` for upgrade flow |
| `foundry-agent-to-m365` | Bridge a Microsoft Foundry agent into M365 | TS | Yes (Bot/AppService+Identity) | Inputs: `foundryEndpoint`, `foundryAgentId` |
| `graph-connector` | External-content connector → Microsoft Graph | TS | Yes (Functions/Connector + Storage + Key Vault + App Insights + Log Analytics) | Heaviest infra footprint |
| `office-addin-config` | Office Add-in scaffolding/config base | TS · common | Yes (Static Web App) | |
| `office-addin-excel-cfshortcut` | Excel custom function + shortcut add-in | TS | Yes (Static Web App) | |
| `office-addin-outlook-taskpane` | Outlook task-pane add-in | TS | Yes (Static Web App) | |
| `office-addin-wxpo-taskpane` | Word/Excel/PowerPoint/Outlook task-pane | TS | Yes (Static Web App) | |
| `office-addin` | Common Office Add-in scaffold (lives in `templates/vsc/common/office-addin/`) | common | Yes (Static Web App) | |
| `office-xml-addin-common` | XML-manifest Office add-in (legacy shape) | common | No | |

## Template-driven question-tree branches (user-visible)

The v3 question model exposes these inputs across templates (full enum in [`_v3-reference/data-model/question-names.md`](../_v3-reference/data-model/question-names.md), but the user-facing branches are):

- **App name + folder** — universal
- **Programming language** — TS / JS / Python / C# (per template's allowed set)
- **Project type** — Bot · Tab · CEA · AI Agent · ME · DA · Connector · Office Add-in
- **LLM provider** — OpenAI · Azure OpenAI · Foundry (for AI agents)
- **API spec path** — for plugin/RAG-from-spec / ME-from-spec templates
- **Authentication** — None · Bearer · OAuth · API Key (for plugin templates with backends)
- **Foundry endpoint + agent ID** — for `foundry-agent-to-m365`
- **Graph connector config** — connector ID / name / description / repos (for `graph-connector`)
- **Office add-in host** — Word / Excel / PowerPoint / Outlook (for cross-host templates)
- **Custom data variant** — `azure-ai-search` / `custom-api` / `customize` (for RAG)

## Optional / post-scaffold features

| Feature | Trigger | What it adds |
|---------|---------|--------------|
| **Add Action** (`atk add action`, VS Code "Add Action") | DA project | Adds a new API plugin action |
| **Add Action from MCP** | DA project, feature-flagged | Adds an MCP-backed action |
| **Add Auth Configuration** (`atk add auth-config`) | DA project with action | Adds OAuth / API-key registration |
| **Add Capability** (`atk add capability`, VS Code "Add Capability") | DA project | Adds a knowledge source (web search · OneDrive/SharePoint · Graph connector · embedded knowledge) |
| **Add SPFx Web Part** (`atk add spfx-web-part`) | SPFx project | Adds another web part to an SPFx solution |
| **Regenerate Action** (`atk regenerate action`) | DA project | Re-generates plugin manifest from a changed spec |
| **Set Sensitivity Label** | DA project, feature-flagged | Annotates DA with a sensitivity label |
| **Sync Manifest** | All projects, feature-flagged | Re-syncs `manifest.json` with TDP-side changes |
| **Update Action with MCP** | DA project | Replaces an action with an MCP-backed one |
| **Add Configurations to Support Actions with Authentication** | DA project | Wires auth boilerplate for actions |

## Integration surface

| Surface | What v3 integrates with |
|---------|-------------------------|
| Microsoft 365 Copilot | Declarative agents, API plugins, MCP actions, knowledge sources, sensitivity labels |
| Microsoft Teams | Bots, message extensions, tabs, app sideload + tenant catalog publish |
| Microsoft Outlook | Add-ins, message extensions (cross-surface) |
| Microsoft 365 app | Add-ins, sideloaded apps |
| Microsoft Foundry | `foundry-agent-to-m365` template; Foundry Agent template |
| Microsoft Agents SDK | Bots and CEAs |
| Teams AI Library | AI Agent templates |
| Microsoft Graph | App catalog publishing, Entra ID app registration, Graph connectors, sensitivity labels, sharing/permissions |
| Teams Developer Portal | Manifest validation, package validation, OAuth/API key registration |
| Bot Framework | Channel registration via ARM |
| Azure | App Service, Functions, Static Web Apps, Bot Service, SQL, AI Search, Key Vault, Storage, Log Analytics, App Insights, Managed Identity |
| Azure DevOps + GitHub Actions | Sample CI/CD pipelines emitted with templates |
| GitHub Copilot Chat (`@m365agents`) | External companion extension; this extension constructs pre-filled queries and link buttons. See [01-product/ux/surfaces/copilot-chat-participant.md](../01-product/ux/surfaces/copilot-chat-participant.md). |
| MCP server (`@microsoft/m365agentstoolkit-mcp`) | Auto-generated `.vscode/mcp.json` references the toolkit's standalone MCP server package; lets MCP-aware tools (Claude, Cursor, etc.) drive the toolkit. See [`packages/vscode-extension/src/utils/mcpUtils.ts`](../../packages/vscode-extension/src/utils/mcpUtils.ts). |
| Sovereign clouds | Setting `M365AgentsToolkit.sovereignCloudEnvironment` switches to GCC M / GCC H / DoD. See [01-product/ux/flows/tenant-and-sovereign-cloud.md](../01-product/ux/flows/tenant-and-sovereign-cloud.md). |
| Dev tunnels | `dev-tunnel` task type for local public-https endpoints. See [01-product/ux/surfaces/local-debug-and-prereqs.md](../01-product/ux/surfaces/local-debug-and-prereqs.md). |
| M365 Agents Playground (Test Tool) | Local emulator for bots / message extensions; auto-detected via `testtool` env. |

## Observability surface

- VS Code output channel + `M365AgentsToolkit.logLevel` setting (Info / Verbose / Debug).
- Telemetry via Application Insights (auto-injected `correlationId`, `component`, `commandName`, `projectId`, `runFrom`).
- Structured `FxError` shape: `source`, `name`, `message`, `displayMessage`, `helpLink`.

## Localisation surface

13 locales: `cs`, `de`, `es`, `fr`, `it`, `ja`, `ko`, `pl`, `pt-BR`, `ru`, `tr`, plus `zh-cn`, `zh-Hans`, `zh-Hant`, `zh-tw`.

Strings keyed in `package.nls.json`, translations in `package.nls.{locale}.json`. Three sub-systems:

1. VS Code extension (`packages/vscode-extension/package.nls*.json`).
2. CLI (`packages/cli/src/resource/`).
3. fx-core (`packages/fx-core/resource/package.nls.*.json`).

## CI/CD surface

Templates ship sample pipelines under `.github/workflows/` (GitHub Actions) and/or `.ado/` (Azure DevOps). They invoke the CLI in non-interactive mode (`--interactive false`).

## Coverage in `features.json`

[`features.json`](../../.dev/features.json) is the source-of-truth for **automated lifecycle test coverage**, not for the full template catalog. As of extraction:

- **16 testable features** + **2 tracked-only features** = 18 entries.
- The remaining 11 shipped templates (Office add-in family, several DA variants, `message-extension-with-existing-api`, `teams-agent-with-data-custom-api-v2`) are **not in the registry** — they ship and scaffold, but lack automated lifecycle coverage.
- Some features-registry entries diverge from folder names: e.g. `basic-tab` is registered as `non-sso-tab` (the original v2 name).
- Some features-registry entries declare a narrower language set than the folders on disk: e.g. `weather-agent` registers TS+JS but ships a Python folder too.

## Headline gaps in v3 (informing v4 priorities)

- **Implicit generator activation order** — first-activated-wins; collisions are silent.
- **`TOOLS` global singleton** — makes parallel lifecycles, isolated tests, and fan-out impossible.
- **Ad-hoc YAML driver dispatch** — driver inputs are not pre-validated; typos surface as `undefined` accesses far from the YAML site.
- **Spread-out manifest manipulation** — multiple drivers edit raw manifest JSON.
- **Heavy CLI startup** — webpack bundle, eager `applicationinsights` load, `node_modules` ~4 GB heap requirement.
- **Coupled `api` + `manifest` packages** — every `api` change rebuilds everything downstream.

These gaps are **product-visible inputs** to v4 design (faster cold start, parallel-safe operations, schema-validated YAML, single-pass manifest layer). They do **not** mean v4 should reuse v3 internals.

## See also

- [capabilities-matrix.md](capabilities-matrix.md) — the v3 + v4 cross-engine summary
- [scenarios/](scenarios/README.md) — end-to-end user scenarios
- [.dev/features.json](../../.dev/features.json) — machine-readable feature registry
- [`_v3-reference/`](../_v3-reference/README.md) — the *internal* v3 extractions (forbidden as v4 design input)
