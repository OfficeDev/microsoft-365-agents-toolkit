---
description: "Use when editing or creating generators, drivers, coordinators, question models, feature flags, or error types in fx-core or core-next. Covers the core engine architecture and patterns."
applyTo: "packages/fx-core/**/*.ts, packages/core-next/**/*.ts"
---

# FxCore Package Conventions

> **v4 (`core-next`) is design-first.** Non-trivial changes (new operation, driver, descriptor, contract, lifecycle stage) MUST start with a docs change under [`docs/02-architecture/`](../../docs/02-architecture/README.md) or [`docs/04-specs/data-model/`](../../docs/04-specs/data-model/README.md), and an ADR under [`docs/02-architecture/09-architecture-decisions/`](../../docs/02-architecture/09-architecture-decisions/README.md) for architectural decisions. The design page is authoritative until the code lands. See [codebase.instructions.md](codebase.instructions.md) §"Source-of-Truth Workflow".
>
> **v3 (`fx-core`) is code-first.** The shipping engine is the truth; docs catch up to it.

## Architecture

```
FxCore (entry point — src/core/FxCore.ts)
  → Coordinator (orchestrates multi-step workflows)
    → Generators (scaffold project templates)
    → Drivers (interact with external services)
    → Question Model (user interaction trees)
  → EnvironmentManager (multi-env support)
  → FeatureFlagManager (gated rollouts)
```

## Generators

Located in `src/component/generator/`. Each scaffolds project templates.

**Class pattern:**
```typescript
export class MyGenerator extends DefaultTemplateGenerator {
  componentName = "my-generator";

  public override activate(context: Context, inputs: Inputs): boolean {
    return inputs[QuestionNames.TemplateName] === TemplateNames.MyTemplate;
  }

  public override async getTemplateInfos(...): Promise<Result<TemplateInfo[], FxError>> {
    // Build replacement map, return template info
  }

  public override async post(...): Promise<Result<GeneratorResult, FxError>> {
    // Post-processing after scaffolding
  }
}
```

- Register in `generatorProvider.ts` — **order matters** (first activated wins)
- `activate()` determines if this generator handles the current input
- `getTemplateInfos()` returns template names, languages, and Mustache replacement maps
- `post()` runs after template files are written (install deps, generate files, etc.)

## Drivers

Encapsulate external service interactions in `src/component/driver/`:

| Driver | Purpose |
|--------|---------|
| `aad/` | Entra ID (Azure AD) app registration |
| `arm/` | Azure Resource Manager deployments |
| `teamsApp/` | Teams app packaging and publishing |
| `deploy/` | Deployment to Azure services |
| `apiKey/` | API key registration |
| `oauth/` | OAuth configuration |

- Drivers return `Result<T, FxError>` — never throw
- Use `getLocalizedString()` for all user-facing messages
- Log with the shared `LogProvider` from context, not `console`
- **Filesystem EAFP**: Use try/catch on file ops, catch `ENOENT` — never `existsSync()` before
  `readFile()` (TOCTOU). See `codebase.instructions.md` > Security > EAFP Pattern.
- **Archive security**: ZIP extraction must guard against Zip Slip (`download.ts`); ZIP uploads
  must validate magic bytes (`teamsDevPortal/client.ts`, `zipDeploy.ts`).

## Question Model

Each platform has its own question tree:

```
src/question/scaffold/
  ├── vsc/createRootNode.ts    (VS Code)
  ├── vs/createRootNode.ts     (Visual Studio)
  └── cli/                     (CLI)
```

- Questions are tree nodes; children activate based on parent answers
- Use `DynamicOptions` (`LocalFunc<StaticOptions>`) for options computed at runtime
- Add options as `{ id: string, label: string, data: TemplateName }` objects

## Error Definitions

Organized by domain in `src/error/`:

```
error/
  ├── common.ts     # Cross-cutting errors
  ├── arm.ts        # ARM deployment errors
  ├── azure.ts      # Azure service errors
  ├── teamsApp.ts   # Teams app errors
  ├── yml.ts        # YAML parsing errors
  └── ...
```

- Each file exports error factory functions or class constructors
- Always include `source` (component name) and stable `name` (for telemetry)
- Use `getDefaultString()` for `message` (English, for logs)
- Use `getLocalizedString()` for `displayMessage` (localized, for users)

## Feature Flags

```typescript
import { featureFlagManager, FeatureFlagName } from "../common/featureFlags";

if (featureFlagManager.getBooleanValue(FeatureFlagName.MyFeature)) {
  // new behavior
} else {
  // old behavior
}
```

- Define new flags in `FeatureFlagName` enum
- Read from environment variables at runtime
- Never hardcode flag values

### Key Feature Flag: `TEAMSFX_V4_CORE`

Gates the migration from `fx-core` (v3) to `core-next` (v4). When enabled, consumers
can switch to the v4 pipeline. Default: `false`. Defined in `featureFlags.ts`.

## core-next (v4 Successor)

`packages/core-next/` is the next-generation replacement for both `packages/api/` and
`packages/fx-core/`. It publishes as `@microsoft/teamsfx-core` v4.0.0.

### What's Different

| Aspect | fx-core (v3) | core-next (v4) |
|--------|-------------|----------------|
| API contracts | Separate `@microsoft/teamsfx-api` package | Merged into `src/api/` |
| Context | `TOOLS` global singleton | `AtkContext` interface (injected) |
| Operations | `FxCore` class methods | `Operation` pipeline with `runOperation()` / `defineOperation()` |
| Templates | Inline in generators | `TemplateRegistry` + scaffold system (download, render, unzip) |
| Drivers | Implicit in `src/component/driver/` | `DriverRegistry` + `createDriver()` factory + built-in drivers (file, script) |
| Errors | `FxError` from api | `AtkError` (extends `FxError`) |
| Declarative Agent | Spread across generators + drivers | Dedicated `declarativeAgent/` module (knowledge, actions, auth, capabilities) |
| Lifecycle | YAML actions dispatched ad-hoc | `lifecycle/` engine (parser → resolver → executor) |
| Project creation | `FxCore.createProject()` monolith | `project/create.ts` operation (template lookup → scaffold) |
| Environment | `EnvironmentManager` class | `environment/envManager.ts` pure functions |
| Telemetry helpers | `TOOLS.telemetry` + utility functions scattered across files | `telemetry/` module — `sendStartEvent`, `instrumentOperation`, correlation via AtkContext |
| Secret masking | SVM + BloomFilter + keyword detection in `common/secretmasker/` | `secretMasker/` — keyword-only regex, pure functions |
| Feature flags | `FeatureFlagManager` singleton | `featureFlags/` — injectable `FeatureFlagRegistry` with `FeatureFlagSource` |
| Localization | `getLocalizedString()` in `common/localizeUtils.ts` singleton | `localization/Localizer` class (injectable, reuses package.nls.json) |
| HTTP client | `wrappedAxiosClient` + `requestUtils` (TOOLS-dependent) | `http/` — `createHttpClient(ctx)` with telemetry interceptors, `sendWithRetry/Timeout` |

### Package Layout

```
packages/core-next/src/
  api/               — Merged contracts (error, types, context, question model, UI, etc.)
  core/              — AtkContext, Operation, AtkError, defineOperation()
  folder.ts          — getTemplatesFolder() — resolves bundled templates/ dir from compiled output
  declarativeAgent/  — Full DA module:
    ├── manifest/        — resolver (wraps @microsoft/app-manifest TeamsManifestWrapper)
    ├── knowledge/       — webSearch, oneDriveSharePoint, graphConnector, embeddedKnowledge, addKnowledge
    ├── actions/         — addAction, addActionFromMCP, removeAction
    ├── auth/            — authInjector
    ├── capabilities/    — sensitivityLabel
    ├── operations.ts    — DA-specific operations (addKnowledgeOp, addActionOp, addMCPActionOp, etc.)
    └── types.ts         — DA domain types (AuthScheme, MCPToolDefinition, AddMCPActionInput, etc.)

> **v2.4 Schema Notes (APIPluginManifestWrapper):**
> - `mcp_tool_description` must be `{ file: "mcp-tools.json" }` (not a bare string) — the v2.4 typed converter expects `MCPTool` object
> - `auth` is **always required** (even `{ type: "None" }` for unauthenticated) — `RuntimeAuthenticationObject` is non-optional in v2.4
> - `namespace` is required in the top-level manifest
> - `schema_version` must be `"v2.4"` to support `RemoteMCPServer` runtime type
  lifecycle/         — YAML lifecycle engine + operations:
    ├── parser.ts        — parseProjectYaml() — reads m365agents.yml into RawProjectModel
    ├── resolver.ts      — resolveLifecycle() — matches actions to registered drivers
    ├── executor.ts      — executeLifecycle() — runs resolved actions in sequence; optional LifecycleProgress callbacks
    ├── types.ts         — DriverStep, RawProjectModel, LifecycleResult, LifecycleProgress, LifecycleAnalysis, LifecycleOperationResult, PostAction, M365TenantInfo, AzureAccountInfo, ResourceGroupInfo
    ├── analyze.ts       — analyzeSteps() — driver introspection (needsM365? needsAzure? unresolvedVars?)
    ├── progress.ts      — createProgressAdapter(ui) — maps LifecycleProgress to platform progress bar; silentProgress for CI
    ├── prerequisites.ts — composable auth gates: ensureM365Auth, ensureAzureAuth, ensureSubscription, ensureResourceGroup, confirmProvision, confirmDeploy
    ├── operations.ts    — provisionOp, deployOp, publishOp — defineOperation wrappers composing prerequisites + executor + env persistence
    └── index.ts         — barrel exports
  project/           — createProjectOp (direct) + createProjectInteractive (question tree → scaffold → tracking ID)
  questions/         — Question tree infrastructure:
    ├── questionNames.ts    — QuestionNames constants (22 canonical question names)
    ├── commonQuestions.ts  — 18 reusable question factory functions
    ├── treeBuilder.ts      — buildQuestionTree(registry) — auto-generates IQTreeNode tree from TemplateRegistry
    ├── traverse.ts         — traverseQuestionTree(tree, ui, inputs) — iterative DFS with back-stack, subtree skipping
    └── index.ts            — Barrel exports
  environment/       — listEnvironments, readEnvFile, writeEnvFile, addEnvironment, resetEnvironment
  teamsApp/          — validateManifestOp, packageAppOp, publishAppOp
  templates/         — Template system:
    ├── registry.ts      — TemplateRegistry, TemplateDescriptor
    ├── types.ts         — TemplateDescriptor type
    ├── scaffold/        — Scaffold pipeline:
    │   ├── scaffolder.ts    — scaffoldTemplates() entry point (remote → fallback → unzip; auto-filter + prefix-strip)
    │   ├── download.ts      — resolveTemplateUrl(), fetchZip(), loadLocalFallback(), unzipWithTransform()
    │   ├── render.ts        — Mustache rendering (.tpl files, preserves undefined vars)
    │   ├── replaceMap.ts    — getTemplateReplaceMap() (appName, safeProjectName, etc.)
    │   └── types.ts         — TemplateInfo, ScaffoldContext, TemplateConfig, convertToLangKey()
    ├── descriptors/     — Built-in template registrations (26 descriptors):
    │   ├── declarativeAgent.ts — 12 DA descriptors (da/* IDs, includes metaos-upgrade with inline scaffoldFn)
    │   ├── bot.ts              — 1 bot descriptor (echo only)
    │   ├── tab.ts              — 1 tab descriptor (basic only)
    │   ├── aiAgent.ts          — 3 AI agent descriptors (with LLM questions)
    │   ├── engineAgent.ts      — 3 custom engine agent descriptors
    │   ├── connector.ts        — 1 connector descriptor (with graph connector questions)
    │   ├── messageExtension.ts — 1 message extension descriptor (search-based)
    │   ├── openApi.ts          — 3 OpenAPI-backed descriptors (da, ai-agent, me)
    │   ├── foundry.ts          — 1 Foundry Agent descriptor (questions: foundryEndpoint, foundryAgentId)
    │   └── index.ts            — registerBuiltinTemplates() + barrel exports
    └── openApi/         — OpenAPI scaffolding support:
        ├── specParserAdapter.ts — SpecParserAdapter interface + StubSpecParserAdapter + createSpecParserAdapter() factory
        ├── realSpecParserAdapter.ts — RealSpecParserAdapter backed by inline specParser module
        ├── scaffoldFn.ts        — makeOpenApiScaffoldFn() factory (validate → scaffold → parse → generate → write)
        └── index.ts             — Barrel exports
  helpers/           — Template helper utilities:
    └── metaOSHelper.ts  — MetaOS upgrade functions: copyExistMetaOSProject(), extendToDA(), unifyProjectID() (native fs, no fs-extra)
  specParser/          — Inline OpenAPI spec parser (merged from @microsoft/m365-spec-parser):
    ├── types.ts         — ParsedSpec, ValidationResult, ErrorType, WarningType, ProjectType, ParseOptions, AuthInfo, etc.
    ├── constants.ts     — SpecParserMessages, HTTPMethods, WellKnownNames, Limits, AdaptiveCardConstants
    ├── parser.ts        — parseSpec() → Result<ParsedSpec>, resolveEnvVars(), hasCircularRefs() (wraps swagger-parser + swagger2openapi)
    ├── utils.ts         — Auth helpers, schema helpers, URL/server validation, parameter generation, naming utils
    ├── validator.ts     — Abstract Validator + CopilotValidator, SMEValidator, TeamsAIValidator, createValidator() factory
    ├── filter.ts        — filterSpec() — filter to selected operations + optimize
    ├── optimizer.ts     — optimizeSpec() — remove unused components/tags/security/vendor extensions
    └── index.ts         — Barrel exports
  drivers/           — DriverRegistry, DriverDescriptor, createDriver() factory:
    ├── types.ts         — DriverDescriptor, DriverConfig, DriverOutput
    ├── registry.ts      — DriverRegistry class + driverRegistry singleton
    ├── createDriver.ts  — createDriver() factory: Zod validation, telemetry, error normalization
    └── builtin/         — Built-in driver implementations:
        ├── index.ts         — registerBuiltinDrivers() + builtinDrivers array (22 drivers)
        ├── file/
        │   ├── createOrUpdateEnvironmentFile.ts — .env file merge driver
        │   └── createOrUpdateJsonFile.ts       — JSON file deep-merge driver
        ├── script/
        │   └── run.ts   — Cross-platform shell execution driver with ::set-output parsing
        ├── cli/
        │   └── runNpmCommand.ts — Runs npm commands (install, build) for deploy lifecycle
        ├── teamsApp/
        │   ├── zipAppPackage.ts     — Zip manifest + icons into Teams app package
        │   ├── validateManifest.ts  — Schema validation via TeamsManifestWrapper
        │   ├── validateAppPackage.ts — Package validation via TDP API
        │   ├── create.ts            — Create/import Teams app in TDP
        │   ├── configure.ts         — Update Teams app config in TDP
        │   ├── update.ts            — Update Teams app in TDP (alias of configure)
        │   ├── publishAppPackage.ts — Publish to org app catalog via Graph API (/beta/appCatalogs/teamsApps)
        │   └── extendToM365.ts      — Sideload app to M365 ecosystem (Outlook, Microsoft 365 app)
        ├── aadApp/
        │   ├── create.ts — Create Entra ID app via MS Graph
        │   └── update.ts — Update Entra ID app properties
        ├── botAadApp/
        │   └── create.ts — Create bot Entra ID app with password
        ├── botFramework/
        │   └── create.ts — Register bot in Bot Framework via ARM
        ├── arm/
        │   └── deploy.ts — ARM/Bicep template deployment
        ├── azureAppService/
        │   └── zipDeploy.ts — Zip deploy to App Service via Kudu
        ├── azureFunctions/
        │   └── zipDeploy.ts — Zip deploy to Functions via Kudu
        ├── oauth/
        │   └── register.ts — Register OAuth config in TDP
        └── apiKey/
            └── register.ts — Register API key secret in TDP
  telemetry/         — Telemetry helpers, error property extraction, instrumentOperation, correlation
  secretMasker/      — Keyword-based credential masking (pure functions)
  featureFlags/      — Injectable FeatureFlagRegistry with built-in flags
  localization/      — Localizer class (package.nls.json bundles)
  http/              — createHttpClient (Axios + telemetry interceptors), retry/timeout helpers
  clients/           — Authenticated service clients:
    ├── teamsDevPortal/  — Teams Developer Portal API (app CRUD, validation, publishing, OAuth, API keys)
    ├── graphApi/        — Microsoft Graph API (Entra ID app registration, updates, passwords, app catalog publish)
    ├── azure/           — Azure ARM (deployments, Kudu zip deploy)
    ├── m365/            — M365 PackageService (sideloading V1/V2 for classic and DA apps; AtkContext+token constructor, Result returns)
    └── index.ts         — Barrel re-export
  index.ts           — Public barrel exports (all modules)
```

### Declarative Agent Module

The DA module in `src/declarativeAgent/` encapsulates all Declarative Agent operations,
re-exporting `TeamsManifestWrapper` from `@microsoft/app-manifest` as the underlying
manifest manipulation layer.

Key patterns:
- Each knowledge source / action function is a pure function returning `Result<void, AtkError>`
- `operations.ts` wraps domain functions as `Operation` records via `defineOperation(name, schema, fn)`
- `defineOperation()` takes 3 positional args: `(name, inputSchema, executeFn)` — NOT an object

### Lifecycle Engine

The lifecycle engine in `src/lifecycle/` replaces the ad-hoc YAML action dispatch in fx-core:
1. `parseProjectYaml()` — reads `m365agents.yml` into a `RawProjectModel`
2. `resolveLifecycle()` — matches each action to a registered driver via `DriverRegistry`
3. `executeLifecycle()` — runs resolved actions in sequence, collecting results; accepts optional `LifecycleProgress` callbacks. Auto-injects `ctx.projectPath` into envMap as `PROJECT_PATH` if not already present (needed by drivers like `teamsApp/zipAppPackage`). **Before each driver call**, temporarily syncs envMap entries into `process.env` so drivers loading external files (ARM parameter JSON, AAD manifest templates) can resolve `${{VAR}}` placeholders produced by earlier steps. Injected vars are cleaned up in a `finally` block to avoid leaking state between steps.

### Lifecycle Operations

The lifecycle operations in `src/lifecycle/operations.ts` compose prerequisites + executor into
complete orchestration functions exposed as `Operation` records via `defineOperation()`:

| Operation | Input | Pipeline |
|-----------|-------|----------|
| `provisionOp` | `{ projectPath, envName, skipConsent? }` | loadEnv → parseYAML → analyzeSteps → ensureM365Auth → ensureAzureAuth → ensureSubscription → ensureResourceGroup → confirmProvision → executeLifecycle → persistEnv |
| `deployOp` | `{ projectPath, envName, skipConsent? }` | loadEnv → parseYAML → confirmDeploy → executeLifecycle → persistEnv |
| `publishOp` | `{ projectPath, envName }` | loadEnv → parseYAML → executeLifecycle → persistEnv |

**Composable prerequisites** (`src/lifecycle/prerequisites.ts`):
- `ensureM365Auth(ctx)` → acquires M365 token, extracts tenant ID from JWT claims
- `ensureAzureAuth(ctx)` → triggers Azure login via `getIdentityCredentialAsync`
- `ensureSubscription(ctx, envMap)` → auto-selects single sub or prompts for multiple
- `ensureResourceGroup(ctx, envMap, subscriptionId, projectName, envName)` → prompts with `rg-{safeName}{suffix}-{envName}` default; also triggered when `AZURE_RESOURCE_GROUP_NAME` is present but empty
- `ensureResourceSuffix(envMap)` → generates/reuses 6-char random suffix
- `confirmProvision(ctx, envName, m365Info?, azureInfo?)` → consent dialog with context details
- `confirmDeploy(ctx, envName)` → skipped for local/testtool/playground/sandbox envs

**Driver introspection** (`src/lifecycle/analyze.ts`):
- `analyzeSteps(steps, envMap?)` returns `LifecycleAnalysis { needsM365, needsAzure, driverIds, unresolvedVars }`
- Uses `M365_DRIVERS` set (14 driver IDs) and `AZURE_DRIVERS` set (5 driver IDs)
- Collects `${{VAR}}` placeholders, checks them against envMap for resolution status

**Progress** (`src/lifecycle/progress.ts`):
- `createProgressAdapter(ui, title?)` — bridges `LifecycleProgress` callbacks to platform `createProgressBar`
- `silentProgress` — no-op implementation for CI/testing

**Post-actions** — operations return `PostAction[]` for consumer to render:
- `{ type: "openUrl", message, url }` — e.g., Azure portal link
- `{ type: "showMessage", message }` — e.g., completion summary

### Scaffolding System

The scaffold system in `src/templates/scaffold/` implements template download and rendering:
- `scaffoldTemplates()` — main entry point; downloads ZIP → unzips with transform → renders Mustache
- `resolveFallbackDir()` — auto-resolves local fallback ZIP directory: explicit param → `TEMPLATE_FALLBACK_DIR` env var → bundled `templates/fallback/`
- Auto-filter: each template entry is filtered by `{templateName}/` prefix from the language-level ZIP, and the prefix is stripped before writing
- `convertToLangKey()` — maps full language names (`"typescript"` → `"ts"`, `"javascript"` → `"js"`, `"csharp"`, `"python"`, `"common"`); used by all descriptors
- `getTemplatesFolder()` — resolves bundled `templates/` directory from compiled output (`src/folder.ts`)
- `getTemplateReplaceMap()` — returns the standard placeholder map (`appName`, `SafeProjectName`, etc.)

**Template ZIP structure:** Each language ZIP (`ts.zip`, `js.zip`, etc.) contains ALL templates for that language
as subdirectories (e.g., `default-bot/`, `custom-copilot-basic/`). The scaffold pipeline auto-filters to the
requested template folder and strips the prefix before writing files to the destination.

**Template name constants in descriptors must match actual folder names** in `templates/vsc/{lang}/`,
not legacy display names. For example, `DATemplateNames.Basic = "declarative-agent-basic"` (not `"copilot-gpt-basic"`).

Built-in descriptors in `src/templates/descriptors/` register **24 templates** across 8 files:
- DA (11), Bot (1), Tab (1), AI Agent (3), Engine Agent (3), Connector (1), Message Extension (1), OpenAPI (3)
- All registered via `registerBuiltinTemplates()` in `descriptors/index.ts`
- OpenAPI descriptors use `makeOpenApiScaffoldFn()` from `templates/openApi/` with a pluggable `SpecParserAdapter`
- `createSpecParserAdapter()` returns `RealSpecParserAdapter` backed by the inline `specParser/` module
- `specParser/` provides: parse (swagger-parser + swagger2openapi), validate (per-project-type), filter, optimize
- Dependencies: `@apidevtools/swagger-parser ^10.1.1`, `swagger2openapi 7.0.8` (exact pin), `openapi-types ^12.1.3` (devDep)

**Bundled fallback ZIPs:** `packages/core-next/templates/fallback/` ships `common.zip`, `ts.zip`, `js.zip`,
`python.zip`, `csharp.zip` — copied from `templates/build/fallback/`. Listed in `package.json` `files` field.

### Question Model Module

`src/questions/` provides the question tree infrastructure for interactive project creation:
- `QuestionNames` — 20 canonical question name constants
- `commonQuestions.ts` — 15 reusable factory functions (`projectNameQuestion()`, `languageQuestion()`, `llmProviderQuestion()`, etc.)
- `buildQuestionTree(registry)` — auto-generates an `IQTreeNode` tree from `TemplateRegistry` metadata (category → template → language → extra questions)
- `traverseQuestionTree(tree, ui, inputs)` — iterative DFS with back-stack, condition evaluation (`equals`, `enum`, `contains`, `ConditionFunc`), subtree-aware skipping, pre-filled input bypass
- `createProjectInteractive(ctx, inputs)` — builds tree + traverses + delegates to `createProjectOp`

### CLI Integration Layer (cli-next)

`packages/cli-next/` consumes core-next operations through a three-layer architecture:

1. **Context** (`src/context.ts`) — `createCliContext(projectPath?)` creates an `AtkContext` wiring
   CLI logger, telemetry adapter, and `CLIUserInteraction` to core-next's DI interface.

2. **Command Factory** (`src/commands/factory.ts`) — `buildNewCommands(parent, registry)` reads
   `TemplateRegistry` and generates Commander subcommands per category. The `Question` interface's
   `cliName`, `cliShortName`, `cliDescription`, `isBoolean` fields are used to generate CLI options
   via `mapQuestionToOption()`. Category slugs are configurable (default: da, cea, ai, me, bot, tab, connector, addin).

3. **Actions** (`src/actions/`) — Pure async functions that bridge CLI to core-next:
   - `createProjectAction` → `runOperation(createProjectOp)`
   - `provisionAction`/`deployAction`/`publishAction` → `runOperation(provisionOp/deployOp/publishOp)`
   - `envListAction`/`envAddAction`/`envResetAction` → `environment.*` functions
   - `validateAction`/`packageAction` → `runOperation(validateManifestOp/packageAppOp)`
   - `listTemplatesAction` → `registry.list()` formatted as table rows

**Handler pattern:**
- `wrapHandler(name, handler)` — telemetry + error handling (no context)
- `wrapHandlerWithContext(name, handler)` — creates `AtkContext` + delegates (for core-next ops)
- `renderPostActions(actions)` — displays `PostAction[]` messages/URLs

### Telemetry Module

`src/telemetry/` provides DI-first telemetry helpers that replace the fx-core TOOLS-based pattern:
- `sendStartEvent()`, `sendSuccessEvent()`, `sendErrorEvent()` — all take `AtkContext` as first arg
- `extractErrorProperties()` — pure function that converts `AtkError` to telemetry-safe properties
- `instrumentOperation()` — wraps any async function with start/end telemetry + timing
- `correlationScope()` / `getCurrentCorrelationId()` — AsyncLocalStorage-based correlation ID propagation
- `TelemetryEvent` enum — fresh events covering only implemented core-next operations
- Zero globals — all functions use `ctx.telemetry` and `ctx.correlationId`

### Secret Masker

`src/secretMasker/` provides keyword-based credential masking (simplified from fx-core's SVM/BloomFilter approach):
- `CREDENTIAL_KEYWORDS` — 100+ expanded credential keyword suffixes
- `matchesCredentialKeyword()` — checks if a key ends with a credential suffix
- `maskSecret()` — masks credential-like values in free-form text
- `maskSecretValues()` — masks values in key-value records where keys match credential keywords
- All pure functions, no singleton state

### Feature Flags Module

`src/featureFlags/` provides an injectable registry replacing `FeatureFlagManager` from fx-core:
- `FeatureFlagRegistry` — injectable class with `register()`, `isEnabled()`, `list()`, `listEnabled()`
- `FeatureFlagSource` interface — abstracts env var reading (defaults to `process.env`, swappable for tests)
- `builtinFlags` — only flags relevant to core-next (V4Core, MCPForDA, SensitivityLabel, DAMetaOS, etc.)
- `createDefaultRegistry()` — factory that returns a registry pre-populated with built-in flags

### Localization Module

`src/localization/` provides locale-aware string lookup reusing the existing `package.nls.json` format:
- `Localizer` class — `loadBundle()`, `getString()`, `getDefaultString()`
- `createLocalizer()` — factory with sensible defaults
- Supports `util.format()` parameter substitution, same as fx-core's `getLocalizedString()`

### HTTP Client Module

`src/http/` provides an instrumented Axios client and retry helpers:
- `createHttpClient()` — returns an `AxiosInstance` with telemetry interceptors (request/response logging, secret masking of URLs)
- `sendWithRetry()` — retries on 5xx errors with exponential backoff
- `sendWithTimeout()` — AbortController-based timeout (replaces old CancelToken pattern)
- Pure functions, no TOOLS dependency

### Service Clients Module

`src/clients/` provides authenticated HTTP clients for external service APIs:

| Client | Directory | Purpose |
|--------|-----------|---------|
| `TeamsDevPortalClient` | `clients/teamsDevPortal/` | Teams Developer Portal REST API — app CRUD, manifest validation, publishing, OAuth registration, API key registration |
| `GraphApiClient` | `clients/graphApi/` | Microsoft Graph API — Entra ID (Azure AD) app registration, updates, password credentials |
| `AzureArmClient` | `clients/azure/` | Azure Resource Manager — ARM/Bicep deployment, deployment status, SCM/Kudu zip deploy |
| `M365PackageService` | `clients/m365/` | M365 sideloading (MOS PackageService API — V1 classic apps, V2 declarative agents) |

Key patterns:
- Each client take `(ctx: AtkContext)` — uses `ctx.auth.m365TokenProvider` or `ctx.auth.azureAccountProvider` for tokens
- HTTP calls via `sendWithRetry()` from `src/http/` — automatic retry + telemetry
- 404 responses return `undefined` (not errors) for idempotent check-before-create patterns
- Types co-located in each client directory (`types.ts`)
- Barrel re-export via `clients/index.ts`

### Driver System

`src/drivers/` provides the driver infrastructure and built-in implementations.

**`createDriver()` factory** — `src/drivers/createDriver.ts`
- Creates a `DriverDescriptor` with automatic Zod pre-validation, telemetry, and error normalization
- Signature: `createDriver<TConfig>({ id, name, inputSchema, execute, rollback? })`
- Validates config against the Zod schema before calling `execute()` — returns `InvalidDriverInput` error on failure
- Wraps unexpected thrown exceptions into `DriverExecutionError` system errors; recognizes `AtkError` plain objects (with `code`/`message`/`kind` properties) and returns them directly instead of wrapping — prevents `[object Object]` serialization
- Sends `driver-start` / `driver-end` telemetry events with duration measurement
- Also generates `validateFn` for preflight validation without execution

**Built-in drivers** — `src/drivers/builtin/`

| Driver ID | File | Purpose |
|-----------|------|---------|
| `file/createOrUpdateEnvironmentFile` | `builtin/file/createOrUpdateEnvironmentFile.ts` | Merges key-value pairs into `.env` files; outputs written env vars |
| `file/createOrUpdateJsonFile` | `builtin/file/createOrUpdateJsonFile.ts` | Deep-merges content into JSON files; supports `content` and `appsettings` aliases |
| `script` | `builtin/script/run.ts` | Cross-platform shell execution; parses `::set-output` / `::set-teamsfx-env` directives from stdout |
| `teamsApp/zipAppPackage` | `builtin/teamsApp/zipAppPackage.ts` | Zips app manifest + icons into Teams app package |
| `teamsApp/validateManifest` | `builtin/teamsApp/validateManifest.ts` | Validates manifest schema via TeamsManifestWrapper |
| `teamsApp/validateAppPackage` | `builtin/teamsApp/validateAppPackage.ts` | Validates a packaged .zip via TDP API |
| `teamsApp/create` | `builtin/teamsApp/create.ts` | Creates/imports Teams app in TDP (idempotent via existingTeamsAppId) |
| `teamsApp/configure` | `builtin/teamsApp/configure.ts` | Updates Teams app configuration in TDP |
| `teamsApp/update` | `builtin/teamsApp/update.ts` | Updates Teams app in TDP (alias of configure, referenced by templates) |
| `teamsApp/publishAppPackage` | `builtin/teamsApp/publishAppPackage.ts` | Publishes/updates app in org app catalog |
| `teamsApp/extendToM365` | `builtin/teamsApp/extendToM365.ts` | Sideloads app to M365 ecosystem via PackageService (V1 classic, V2 DA) |
| `aadApp/create` | `builtin/aadApp/create.ts` | Creates Entra ID (Azure AD) app registration via MS Graph |
| `aadApp/update` | `builtin/aadApp/update.ts` | Updates Entra ID app properties (redirect URIs, identifier URIs, etc.); resolves `${{VAR}}` env placeholders in AAD manifest before parsing |
| `botAadApp/create` | `builtin/botAadApp/create.ts` | Creates bot-specific Entra ID app with password credential |
| `botFramework/create` | `builtin/botFramework/create.ts` | Registers bot channel in Bot Framework via ARM |
| `arm/deploy` | `builtin/arm/deploy.ts` | Deploys ARM/Bicep templates via Azure Resource Manager |
| `azureAppService/zipDeploy` | `builtin/azureAppService/zipDeploy.ts` | Zip-deploys to Azure App Service via Kudu API |
| `azureFunctions/zipDeploy` | `builtin/azureFunctions/zipDeploy.ts` | Zip-deploys to Azure Functions via Kudu API |
| `oauth/register` | `builtin/oauth/register.ts` | Registers OAuth configuration in TDP (Custom + MicrosoftEntra providers) |
| `apiKey/register` | `builtin/apiKey/register.ts` | Registers API key secret in TDP |
| `cli/runNpmCommand` | `builtin/cli/runNpmCommand.ts` | Runs npm commands (install, build) for deploy lifecycle |
| `cli/runDotnetCommand` | `builtin/cli/runDotnetCommand.ts` | Runs dotnet commands (publish, build) for csharp deploy lifecycle |

**Registration:** `registerBuiltinDrivers()` in `builtin/index.ts` registers all built-in drivers with the global `driverRegistry`. Idempotent — safe to call multiple times. Call once at startup before executing any lifecycle.

**Key design decisions vs fx-core:**
- Zod pre-validation before execution (fx-core validated late or not at all)
- No TypeDI `@Service()` decorators — plain `createDriver()` factory + `driverRegistry.register()`
- Controlled `process.env` sync — executor temporarily injects envMap into `process.env` before each driver call (for drivers loading external files with `${{VAR}}` placeholders), then cleans up in `finally`; drivers themselves return outputs for the executor to write
- No external deps for file I/O — uses Node.js built-ins instead of `dotenv`, `fs-extra`, `comment-json`
- Driver IDs match fx-core naming (`file/createOrUpdateEnvironmentFile`, `script`) for YAML backward compatibility

### Guidelines for core-next

- All conventions from fx-core apply (Result pattern, localization, testing)
- Import from `@microsoft/teamsfx-core` — no separate api package import
- Use `AtkContext` instead of the `TOOLS` global singleton
- Register templates via `TemplateRegistry.register()` instead of generator activation order
- Register drivers via `DriverRegistry.register()` instead of implicit module loading
- Define drivers using `createDriver({ id, name, inputSchema, execute, rollback? })` — returns a `DriverDescriptor`
- Built-in drivers live in `src/drivers/builtin/` and are registered via `registerBuiltinDrivers()`
- Use `defineOperation(name, inputSchema, executeFn)` — 3 positional args, not an object
- DA operations re-export from `@microsoft/app-manifest` — don't duplicate manifest logic
- Pure functions preferred — side effects only at the operation boundary

### Testing in core-next

Test infrastructure mirrors the fx-core/cli pattern (Mocha + Chai + Sinon + NYC):

| Config | Purpose |
|--------|---------|
| `.mocharc.js` | ts-node/register, spec reporter, no-experimental-strip-types |
| `.nycrc` | 50% coverage threshold, excludes `src/api/**/*` |

**Test layout:**

```
tests/
  unit/
    testHelper.ts        — createMockContext() shared across all unit tests
    core/                — error.test.ts, context.test.ts, operation.test.ts
    telemetry/           — helpers, instrumentOperation, correlation, errorProperties
    secretMasker/        — masker.test.ts
    featureFlags/        — registry.test.ts
    templates/
      descriptors.test.ts    — 43 descriptor registration + metadata tests
      openApi.test.ts        — OpenAPI descriptor + scaffoldFn tests
      registry.test.ts       — TemplateRegistry CRUD tests
    questions/
      commonQuestions.test.ts — 15 question factory tests
      treeBuilder.test.ts    — buildQuestionTree tests
      traverse.test.ts       — traverseQuestionTree unit tests (20 tests)
      traverseIntegration.test.ts — build+traverse integration tests (3 tests)
    lifecycle/
      analyze.test.ts          — 9 tests: M365/Azure detection, placeholder resolution, dedup, nesting
      progress.test.ts         — 6 tests: createProgressAdapter (start, title, step, end) + silentProgress
      prerequisites.test.ts    — 29 tests: ensureM365Auth, ensureAzureAuth, ensureSubscription, ensureResourceGroup, ensureResourceSuffix, confirmProvision, confirmDeploy
      operations.test.ts       — 16 tests: provisionOp, deployOp, publishOp (auth gates, consent, env persistence, post-actions, validation)
    drivers/             — registry.test.ts
      builtin/           — fileDrivers, scriptDriver, registration, teamsAppDrivers,
                           teamsAppPlatform, entraAndBotDrivers, azureInfraDrivers,
                           authPluginDrivers
    clients/             — teamsDevPortor, graphApi, azure
  integration/
    operationPipeline.test.ts  — runOperation → telemetry sequence, correlation propagation
    crossCutting.test.ts       — feature flags + masking, localizer, secret masking realistic
    daPackaging.test.ts        — full DA packaging pipeline with plugin, spec, knowledge
    daOperations.test.ts       — DA operations on real filesystem: addKnowledge (web-search,
                                  embedded), addExistingPlugin, setSensitivityLabel,
                                  setConversationStarters, extendToM365Op validation
    lifecycleExecution.test.ts — YAML parse → driver execution → env-var chaining between steps
```

**Key test helper:** `createMockContext()` in `tests/unit/testHelper.ts` returns a fully-stubbed
`AtkContext` with sinon stubs for `telemetry.sendEvent`, `telemetry.sendErrorEvent`, `logger`,
`ui`, and `auth`. Use this for every core-next unit test.

**Test commands:**

```bash
cd packages/core-next
npm run test:unit           # 492 unit tests (with NYC coverage)
npm run test:integration    # 24 integration tests
npm run test                # Alias for test:unit
npm run lint                # ESLint check (0 errors required)
npm run format:check        # Prettier check (CI gate)
```

**CI:** core-next is tested by `.github/workflows/ci-next.yml` (separate from old `unit-test.yml`).
Jobs: build → lint, format-check, unit-test (80% coverage gate via nyc), integration-test
(both core-next and cli-next integration tests run in the integration-test job).
Uses `pnpm --filter ./packages/core-next` (directory path, not package name — required because
old fx-core shares the same npm package name `@microsoft/teamsfx-core`).

**E2E:** `.github/workflows/e2e-test-next.yml` runs cli-next E2E tests (daily schedule + PR + manual).
Failures produce a GitHub Step Summary with stats, failed test table, and log tail.

## Global Variables

`src/common/globalVars.ts` — shared state:
- `TOOLS` — singleton with `LogProvider`, `TelemetryReporter`, `TokenProvider`, `UserInteraction`
- Set once during initialization; access via import anywhere in fx-core

## Exports

`index.ts` uses **selective named exports** (not barrel `export *`):

```typescript
export { FxCore } from "./core/FxCore";
export { featureFlagManager, FeatureFlagName } from "./common/featureFlags";
export { AppStudioScopes } from "./common/constants";
```

- Only export what CLI, VS Code extension, and server need
- Keep internal utilities unexported to maintain encapsulation

## Localization

```typescript
import { getLocalizedString, getDefaultString } from "../common/localizeUtils";

// For user-facing display
const msg = getLocalizedString("teamsfx.myFeature.description", param1);

// For logs and telemetry (always English)
const log = getDefaultString("teamsfx.myFeature.description", param1);
```

- String keys support `util.format()` parameter substitution
- Never concatenate translated strings — use parameterized keys
- Add keys to `package.nls.json` (English) with translations in locale-specific files

## Testing

- Tests in `tests/` mirror `src/` structure
- 50+ granular test scripts: `test:core`, `test:component`, `test:bot`, etc.
- Mock I/O boundaries (HTTP clients, file system, Azure SDK) with Sinon stubs
- Test both `ok` and `err` paths for every `Result`-returning function
- Restore stubs in `afterEach` — always call `sinon.restore()`

### Integration Tests

Integration tests in `tests/integration/` exercise the full FxCore pipeline (scaffold → provision)
with only external HTTP boundaries mocked. They are driven by a **feature registry** — see
[features.instructions.md](features.instructions.md) for the canonical feature list and how to
add new test coverage.

Key files:
- `featureRegistry.ts` — single source of truth for all testable features
- `testBuilders.ts` — reusable `scaffoldTest()` / `provisionTest()` factories
- `coverageTracker.ts` — runtime tracking of tested (feature × lifecycle × entryPoint)
- `reportGenerator.ts` — generates JSON + HTML coverage reports

Run: `npm run test:integration`
