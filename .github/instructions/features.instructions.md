---
description: "Use when exploring toolkit capabilities, adding new templates, or writing integration tests. Provides the canonical list of features and templates with rich metadata."
applyTo: "packages/fx-core/**/*.ts, templates/**, .dev/features.json"
---

# Feature Registry — Toolkit Capabilities Reference

The **single source of truth** for all testable features lives in two files:

```
.dev/features.json                                    ← Human/machine readable (edit this)
packages/fx-core/tests/integration/featureRegistry.ts ← Typed wrapper (loads JSON, adds helpers)
```

> **v4 Note:** `packages/core-next/src/templates/` contains a `TemplateRegistry` with
> `TemplateDescriptor` records. **26 descriptors** are registered across 9 files in
> `src/templates/descriptors/`: DA (12), Bot (1), Tab (1), AI Agent (3), Engine Agent (3),
> Connector (1), Message Extension (1), OpenAPI (3), Foundry (1). Registered via `registerBuiltinTemplates()`.
> Descriptors support `questions?: QuestionSpec[]` for template-specific prompts (e.g., LLM
> provider selection for AI agents, graph connector config for connectors).
> Descriptors support `testable?: boolean` (defaults to `true`). Set `testable: false` for
> templates that cannot be E2E-tested automatically (e.g., require interactive input or lack
> required artifacts). Currently 4 descriptors are `testable: false`:
> - `da/api-plugin-from-spec`, `ai-agent/rag-from-spec`, `me/from-spec` — require interactive `apiSpecPath` input
> - `da/graph-connector` — template artifact lacks `appPackage/manifest.json`
> OpenAPI descriptors use `makeOpenApiScaffoldFn()` with a `RealSpecParserAdapter` backed
> by the inline `specParser/` module (parse, validate, filter, optimize OpenAPI specs).
> Dependencies: `@apidevtools/swagger-parser ^10.1.1`, `swagger2openapi 7.0.8` (exact pin).
>
> **Scaffold pipeline is E2E-verified.** The scaffold system now includes:
> - `resolveFallbackDir()` — auto-resolves local fallback ZIPs (explicit param → `TEMPLATE_FALLBACK_DIR` env → bundled `templates/fallback/`)
> - Auto-filter by `{templateName}/` prefix in language-level ZIPs + prefix stripping before file write
> - `convertToLangKey()` — maps `\"typescript\"` → `\"ts\"`, `\"javascript\"` → `\"js\"`, etc.
> - `getTemplatesFolder()` in `src/folder.ts` — resolves bundled templates dir from compiled output
> - Bundled fallback ZIPs shipped in `packages/core-next/templates/fallback/` (common, ts, js, python, csharp)
> - **9/9 E2E scaffolds verified** via cli-next: bot (TS/JS/Python), DA basic, AI chat, CEA basic, CEA weather (Python), tab basic, connector graph
>
> **Template name constants must match actual folder names** in `templates/vsc/{lang}/`, not legacy
> display names. Example: `DATemplateNames.Basic = \"declarative-agent-basic\"` (not `\"copilot-gpt-basic\"`).
>
> The question model in `src/questions/` provides `buildQuestionTree(registry)` to auto-generate
> `IQTreeNode` trees from registry metadata, and `traverseQuestionTree(tree, ui, inputs)` for
> interactive traversal with back navigation. `createProjectInteractive(ctx, inputs)` combines
> both for question-driven project creation.
>
> The driver system has a `createDriver()` factory and **22 built-in driver implementations**
> registered via `registerBuiltinDrivers()`. The `publishAppPackage` driver now uses
> `GraphApiClient` (Graph `/beta/appCatalogs/teamsApps`) instead of the legacy TDP API.
> Service clients in `src/clients/` provide authenticated access to Teams Developer Portal,
> Microsoft Graph (Entra ID), Azure ARM APIs, and M365 PackageService (sideloading).
>
> The lifecycle engine in `src/lifecycle/` provides:
> - **Parser + executor**: `parseProjectYaml()` → `executeLifecycle()` with optional `LifecycleProgress` callbacks
> - **Driver introspection**: `analyzeSteps()` determines M365/Azure prerequisites from step driver IDs
> - **Composable prerequisites**: `ensureM365Auth`, `ensureAzureAuth`, `ensureSubscription`, `ensureResourceGroup`, `confirmProvision`, `confirmDeploy`
> - **Operations**: `provisionOp`, `deployOp`, `publishOp` — complete orchestration via `defineOperation()`
> - **Progress**: `createProgressAdapter(ui)` bridges to platform progress bar; `silentProgress` for CI
>
> **v4 Testing:** Both `packages/core-next/` and `packages/cli-next/` have full test
> infrastructure (Mocha + Chai + Sinon + NYC). core-next has **606 unit tests**;
> cli-next has **87 unit + 81 integration tests**; core-next has **48 integration tests**
> (includes specParser pipeline and adapter tests).
> Plus **9 E2E scaffold tests** verified via cli-next + **9 OpenAPI spec-parser E2E tests**.
> Run with `npm run test:unit` / `npm run test:integration`.
> CI: `.github/workflows/ci-next.yml` (build → lint, format-check, unit-test, integration-test).
> ESLint flat config with `shared` + `header`; Prettier shared config; 80% coverage gate.
>
> **v4 Lifecycle coverage:** 38 of 39 templates have all drivers registered for full
> provision/deploy/publish support. The only gap is `typeSpec/compile` for the
> `declarative-agent-typespec` template. Drivers added in this pass:
> - `teamsApp/update` — alias of configure, referenced by all templates
> - `teamsApp/extendToM365` — sideloads to M365 ecosystem (V1 classic, V2 DA)
> - `cli/runNpmCommand` — runs npm install/build for deploy lifecycle
> Auth providers are real MSAL-based implementations (not stubs).
>
> The `atk new` command tree is generated from `TemplateRegistry` via `buildNewCommands()` in
> `src/commands/factory.ts`. Adding a new `TemplateDescriptor` to the registry automatically
> creates a new CLI subcommand (e.g., `atk new da <template>`, `atk new bot <template>`).
> Category slugs: da, cea, ai, me, bot, tab, connector, addin.

**To add or update a feature, edit `.dev/features.json`.** The TypeScript wrapper
validates `templateName` values against the `TemplateNames` enum at load time and
will throw a clear error if a name doesn't match.

## What it contains

Each feature entry in `.dev/features.json` has:

| Field | Purpose |
|-------|---------|
| `id` | Unique identifier (e.g., `"default-bot"`) |
| `name` | Human-readable name |
| `description` | Full description of what the template produces |
| `category` | Grouping: Bot, Tab, Messaging Extension, Declarative Agent, Custom Engine Agent, AI Agent, Connector, Office Add-in |
| `templateName` | String value from the `TemplateNames` enum — validated at load time |
| `languages` | Supported languages (e.g., `["typescript", "javascript", "python"]`) |
| `lifecycles` | Testable lifecycle stages: scaffold, provision, deploy, publish |
| `projectType` | The `inputs[QuestionNames.ProjectType]` value (metadata only — ignored when `templateName` is set) |
| `capabilities` | Semantic tags (e.g., `"ai-powered"`, `"function-calling"`, `"conversational"`) |
| `generatedFiles` | Key files in the scaffolded project |
| `entryPoints` | Which entry points can exercise it: fx-core, cli, server, vscode |
| `adoSuiteId` | Links to ADO E2E test suite in plan 24569079 for traceability |
| `adoTestCaseCount` | Number of manual E2E test cases in the linked ADO suite |

The JSON also has a `trackedOnly` array for features that can't be scaffold-tested
locally but are tracked for ADO traceability.

## ADO traceability

Each feature links back to its corresponding E2E test suite in the
**TeamsFx Test Plan (TTK&CLI)** (ADO plan ID 24569079). This enables:

- Cross-referencing integration test coverage with manual E2E test plans
- Coverage gap analysis between automated and manual testing
- The HTML coverage report includes an ADO traceability table

Use `getAdoTraceability()` to get all features with ADO suite links.
Use `getFeatureByAdoSuiteId(suiteId)` to look up a feature by its ADO suite.

## How to use it

### Finding features by capability

```typescript
import { getFeatures, FeatureCategory, Lifecycle } from "./featureRegistry";

// All bot features
getFeatures({ category: FeatureCategory.Bot });

// All features that support provisioning
getFeatures({ lifecycle: Lifecycle.Provision });

// AI-powered features (search capabilities array)
getFeatures().filter(f => f.capabilities.includes("ai-powered"));
```

### Finding a specific feature

```typescript
import { getFeatureById } from "./featureRegistry";

const bot = getFeatureById("default-bot");
// → { name: "Echo Bot", templateName: "default-bot", languages: ["typescript", "javascript", "python"], ... }
```

### Writing integration tests

Tests are generated automatically from the registry. To add coverage for a new feature:

1. Add an entry to `.dev/features.json`
2. Run `npm run test:integration` — the coverage report will flag it as untested
3. The `scaffoldTest()` and `provisionTest()` builders in `testBuilders.ts` handle the rest

### Coverage reports

After running integration tests, reports are generated at:
- `coverage/integration/integration-coverage.json` — machine-readable
- `coverage/integration/integration-coverage.html` — visual matrix (feature × lifecycle × entry point) + ADO traceability table

## Current features (14 testable, 2 tracked-only)

| Category | Template | Languages | ADO Suite | v4 E2E |
|----------|----------|-----------|-----------|--------|
| Bot | Echo Bot (`default-bot`) | TS, JS, Python | 24569101 | **Yes** |
| Tab | Basic Tab (`basic-tab`) | TS | 24569106 | **Yes** |
| Custom Engine Agent | Basic Custom Engine Agent | TS, JS, Python | 34834051 | **Yes** |
| Custom Engine Agent | Weather Agent (function-calling) | TS, JS, Python | 34648283 | **Yes** |
| Custom Engine Agent | Teams Collaborator Agent | TS | 35527236 | — |
| AI Agent | AI Chat Bot (`custom-copilot-basic`) | TS, JS, Python | 27042287 | **Yes** |
| AI Agent | AI Agent with AI Search (RAG) | TS, JS, Python | 27689412 | — |
| AI Agent | AI Agent with Custom Data (RAG) | TS, JS, Python | 27689419 | — |
| Declarative Agent | DA Basic (`declarative-agent-basic`) | common | 27971458 | **Yes** |
| Declarative Agent | DA with API Plugin (No Auth) | TS, JS | 27971458 | — |
| Declarative Agent | DA with API Plugin (OAuth) | TS, JS | 27971458 | — |
| Declarative Agent | DA with API Plugin (Bearer) | TS, JS | 27971458 | — |
| Connector | Graph Connector (`graph-connector`) | TS | 32019603 | **Yes** |
| Messaging Extension | Message Extension v2 (`message-extension-v2`) | TS, Python | 34869329 | — |
| *AI Agent* | *RAG Custom API* (needs API spec) | — | *27588348* | — |
| *AI Agent* | *Foundry Proxy Agent* (C# only) | — | *36750068* | — |

*Italic entries are tracked for ADO traceability only — not scaffold-testable locally.*

## Adding a new feature

1. **Add local template** in `templates/vsc/{lang}/{templateDir}/`
2. **Add entry to `.dev/features.json`** with full metadata
3. **Set `adoSuiteId`** if an ADO E2E suite exists for traceability
4. **Run integration tests** — validates `templateName` against enum and verifies scaffold
5. **Update this file** if adding a new category
