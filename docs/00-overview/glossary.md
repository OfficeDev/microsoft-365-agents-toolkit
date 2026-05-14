# Glossary

Canonical terminology used across this site, the codebase, and instructions. Where a term has a v3 / v4 split, both are listed.

## Product

| Term | Meaning |
|------|---------|
| **M365 Agents Toolkit** / **ATK** | This product. Successor naming for Teams Toolkit. CLI binary: `atk`. |
| **Declarative Agent (DA)** | An agent for Microsoft 365 Copilot defined declaratively in a manifest, optionally with knowledge sources, actions (API plugins / MCP), and capabilities (e.g. sensitivity labels). |
| **Custom Engine Agent (CEA)** | A code-first agent built on the Microsoft Agents SDK / Teams AI Library. |
| **Teams Agent** | A bot-based agent built on the Bot Framework and Teams AI Library, deployed as a Teams app. |
| **API Plugin** | A capability that lets an agent call an OpenAPI-described HTTP API. |
| **MCP Action** | A capability backed by a Model Context Protocol server (added via `add-action-from-mcp`). |
| **Graph Connector** | A connector that ingests external content into the Microsoft Graph. |
| **M365 Agents Playground** | A local Teams simulator for chat-based apps (bots, message extensions, AI agents, CEA). No login required. Does not support tabs or connectors. Installed via `devTool/install` action; launched from VS Code's "Debug in Microsoft 365 Agents Playground" configuration. |

## Lifecycle

| Term | Meaning |
|------|---------|
| **Scaffold** | Generate a new project from a template. |
| **Provision** | Register the M365 resources (Teams App, Entra App, OAuth connection, API Key) and optionally Azure resources needed to run the app. |
| **Deploy** | Push code (zip, container, etc.) to the provisioned resources. |
| **Publish** | Submit the Teams app package to a tenant catalog or the M365 store. |
| **Lifecycle YAML** | `m365agents.yml` — declarative list of drivers executed per stage (provision, deploy, publish). |
| **Driver** | The unit that implements one YAML action (e.g. `arm/deploy`, `teamsApp/publishAppPackage`). Takes a Zod-validated `with:` block; writes outputs to `writeToEnvironmentFile`. |
| **`writeToEnvironmentFile`** | YAML key on a driver step that maps driver output keys to env var names in `.env.{envName}`. The mechanism by which provision state flows to deploy and publish steps. |

## Specs

| Term | Meaning |
|------|---------|
| **Domain** | One of the 7 capability areas that partition toolkit behaviour: Scaffolding, Lifecycle, Identity, Manifest, Extensibility, Environment, Collaboration. Each has a domain spec. |
| **Domain Spec** | Defines a domain's boundary (in-scope / out-of-scope), inter-domain interfaces, and invariants. Lives in `docs/04-specs/domains/`. |
| **Operation Spec** | Defines exactly what one operation does: inputs, outputs, Acceptance Criteria table, boundary, invariants, and error cases. Lives in `docs/04-specs/operations/<domain>/`. One spec → one or more test files. |
| **Acceptance Criteria (AC)** | ID-based table in an operation spec (`AC-01`, `AC-02`, …). Each row maps 1:1 to a test case. Test names include the AC ID. |
| **Invariant** | A constraint in a spec that must never be violated by any implementation of that operation or domain (e.g. "clientSecret must never appear in logs"). |
| **Human Gate** | A mandatory review checkpoint in the vibe-coding workflow. Gate 1: resolve AI-surfaced ambiguities. Gate 2: approve the AC table. No code is written before Gate 2. |

## Engine v3

| Term | Meaning |
|------|---------|
| `FxCore` | Class entry point in `packages/fx-core`. Methods orchestrate scaffolding, provision, deploy, publish. |
| `TOOLS` | Module-scoped singleton holding `LogProvider`, `TelemetryReporter`, `TokenProvider`, `UserInteraction`. |
| `Generator` | A `DefaultTemplateGenerator` subclass; first one whose `activate()` returns `true` wins. |
| **Question Model (QM)** | Tree of `UserInputQuestion` nodes evaluated per platform (VS Code, VS, CLI). |
| `FxError` | Base error interface in `@microsoft/teamsfx-api`. Concrete: `UserError`, `SystemError`. |

## Engine v4 (`core-next`)

| Term | Meaning |
|------|---------|
| `AtkContext` | Injected context replacing `TOOLS`. Holds `logger`, `telemetry`, `ui`, `auth`, `correlationId`. |
| `AtkError` | v4 error type; extends the same `FxError` shape. Stable `name` string is the telemetry partition key. |
| `Operation` | Record produced by `defineOperation(name, schema, fn)`. Executed via `runOperation()`. |
| `TemplateDescriptor` | Record registered with `TemplateRegistry`. Drives both scaffold and CLI command generation. |
| `DriverDescriptor` | Record produced by `createDriver({ id, name, inputSchema, execute, rollback? })`. Registered with `DriverRegistry`. |
| **Lifecycle engine** | `parseProjectYaml` → `resolveLifecycle` → `executeLifecycle`, with composable prerequisites and `LifecycleProgress`. |
| `PostAction` | `{ type: "openUrl" \| "showMessage", ... }` returned by operations for the consumer to render. |
| **Built-in fallback ZIPs** | `packages/core-next/templates/fallback/*.zip` shipped for offline scaffolding. |

## Cross-cutting

| Term | Meaning |
|------|---------|
| `Result<T, FxError>` | `neverthrow` Result; the universal return type for fallible operations. Never throw for expected failures. |
| **Correlation ID** | UUID propagated via `Correlator.run()` (v3) or `AsyncLocalStorage` (v4) for distributed tracing. |
| **Feature flag** | Boolean env-driven gate. v3: `featureFlagManager` singleton. v4: injectable `FeatureFlagRegistry`. Key flag: `TEAMSFX_V4_CORE`. |
| **Secret masker** | Module that redacts credentials in logs/telemetry. v3: SVM + BloomFilter + keywords. v4: keyword-only regex. |
| **EAFP** | "Easier to Ask Forgiveness than Permission" — the required pattern for filesystem access (avoids TOCTOU). |
| **Zip Slip** | Path-traversal attack on archive extraction. Mitigated by entry-name validation in download/unzip. |

## Files in a project

| File | Purpose |
|------|---------|
| `m365agents.yml` | Lifecycle action definitions (provision, deploy, publish). |
| `m365agents.local.yml` | Local-debug-only actions. |
| `appPackage/manifest.json` | Teams app manifest (schema 1.x → 2.4). |
| `appPackage/declarativeAgent.json` | Declarative Agent definition. |
| `appPackage/aiPlugin.json` | API plugin manifest (v2.4: `mcp_tool_description`, `auth`, `namespace`). |
| `infra/*.bicep` | Azure infrastructure as code (Bicep templates). |
| `env/.env.{envName}` | Per-environment variables output by provision drivers; secrets in `.env.{envName}.user`. |
| `.fx/` | Legacy v3 state directory (per-project). |
| `~/.fx/account/` | Shared MSAL token cache (v3 + v4). |
