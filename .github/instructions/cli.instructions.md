---
description: "Use when editing or creating CLI commands, prompts, argument parsing, login flows, telemetry, or error handling in the CLI package (v3 or v4)."
applyTo: "packages/cli/**/*.ts, packages/cli-next/**/*.ts"
---

# CLI Package Conventions

> **v4 (`cli-next`) is design-first.** Non-trivial changes (new command group, new action, new auth provider, telemetry shape change, command-factory change) MUST start with a docs change under [`docs/01-product/ux/`](../../docs/01-product/ux/README.md) (surfaces / flows) or [`docs/05-engineering/package-reference/cli-next.md`](../../docs/05-engineering/package-reference/cli-next.md). For architectural changes, file an ADR. The design is authoritative until the code lands. See [codebase.instructions.md](codebase.instructions.md) §"Source-of-Truth Workflow".
>
> **v3 (`cli`) is code-first.** The shipping CLI is the truth; docs catch up to it.

## Architecture

```
cli.js (entry)
  → activate() in index.ts
    → CLIEngine.start(rootCommand)
      → findCommand() → parseArgs() → validateOptions() → handler()
```

## Command Model

Commands are declarative `CLICommand` objects in `src/commands/models/`:

```typescript
export const myCommand: CLICommand = {
  name: "my-command",
  description: "What this command does",
  options: [...MyCommandOptions],  // from fx-core or local definitions
  handler: async (ctx: CLIContext): Promise<Result<undefined, FxError>> => {
    // delegate to fx-core
    const result = await ctx.core.myOperation(ctx.inputs);
    return result;
  },
  telemetry: { event: TelemetryEvent.MyCommand },
};
```

### Adding a New Command

1. Create `src/commands/models/mycommand.ts` with a `CLICommand` object
2. Export from `src/commands/models/index.ts`
3. Add to `commands` array in `src/commands/models/root.ts`
4. Add localized strings in `src/resource/`

## Command Engine (7 Phases)

1. **Find command** — tree traversal with alias support
2. **Parse args** — custom parser: `--key=value`, `--key value`, `-k value`
3. **Version check** — `--version` flag
4. **Help display** — `--help` flag
5. **Validate options** — skipped in interactive mode (reserved options only)
6. **Version compatibility** — project version check + phantom migration
7. **Execute handler** — wrapped in `Correlator.run()` for tracing

## Interactive vs Non-Interactive

- **Interactive** (default): CLI options are discarded; question model drives prompts via `CLIUserInteraction`
- **Non-interactive** (`CI_ENABLED=true`): All options must be passed as CLI args; prompts auto-return defaults
- `reservedOptionNamesInInteractiveMode` — options always parsed even in interactive mode

## User Interaction

`CLIUserInteraction` wraps `@inquirer/prompts`:

| Method | Prompt |
|--------|--------|
| `singleSelect()` | Custom list prompt (`customizedListPrompt.ts`) |
| `multiSelect()` | Custom checkbox prompt (`customizedCheckboxPrompt.ts`) |
| `input()` | `@inquirer/prompts.input` with green theme |
| `password()` | `@inquirer/prompts.password` with mask |
| `confirm()` | `@inquirer/prompts.confirm` |

- `ScreenManager` manages terminal output (pause/continue around prompts)
- Respects `CI_ENABLED=true` — returns defaults without prompting

## Error Classes

All extend `UserError` from `@microsoft/teamsfx-api`:

- `MissingRequiredOptionError` — required CLI option not provided
- `MissingRequiredArgumentError` — positional arg missing
- `InvalidChoiceError` — value not in allowed choices
- `UnknownCommandError` — with edit-distance "Did you mean?" suggestions
- `UnknownOptionError`, `ArgumentConflictError`

Messages are localized via `src/resource/` strings.

## Telemetry

```typescript
CliTelemetry.sendTelemetryEvent(eventName, properties, measurements);
CliTelemetry.sendTelemetryErrorEvent(eventName, error, properties, measurements);
```

- Auto-injected: `component`, `commandName`, `projectId`, `correlationId`, `runFrom` (CI platform detection)
- `withRootFolder()` sets shared project-scoped properties
- Secrets in `commandFull` are auto-masked

## Argument Parsing

Custom recursive descent parser — no external library (yargs, commander):

- Boolean: `--flag` (true), `--flag false`
- String: `--key value` or `--key=value`
- Array: `--arr a,b,c` (comma-separated)
- Short names: `-k value`

## Login Flows

Singleton providers in `src/commonlib/`:
- `AzureLogin` — MSAL with interactive code flow
- `M365Login` — MSAL with `CryptoCachePlugin` for secure token caching
- Windows: native broker plugin for WAM integration
- Flow: check cache → interactive browser consent → cache token

## Output & Color

- `CLILogger` implements `LogProvider` with chalk color support
- `ScreenManager.writeLine()` for formatted output
- `colorize.ts` wraps chalk: red=error, green=success, cyan=links
- `Progress` class manages progress bars with `start()`/`end(success)` lifecycle

---

# CLI v4 (`packages/cli-next`) — Next-Generation Architecture

> **Status:** Phase 4b complete + CI. Lifecycle commands (provision/deploy/publish) wired with real auth
> providers and 21 registered drivers. 38/39 templates fully supported (only typeSpec/compile missing).
> Registry-driven command factory, action layer, interactive UI, MSAL auth.
> CI: `ci-next.yml` (lint, format-check, unit-test, integration-test for both core-next & cli-next).
> E2E: `e2e-test-next.yml` (daily schedule + PR + manual; failure summary via `$GITHUB_STEP_SUMMARY`).
> **Feature flag:** `TEAMSFX_V4_CORE` (default: off)
>
> **Scripts:** `lint`, `format`, `format:check`, `test:unit`, `test:integration`, `build`, `bundle`, `package`
> **Bundling:** esbuild (`esbuild.mjs`) — single-file CJS bundle for production. See Bundling section below.
> **ESLint config:** flat config (`eslint.config.mjs`) with `shared` + `header` (no `promise` — too many false positives in CLI stubs).

## Architecture

```
cli.js (entry — sets TEAMSFX_CLI_BIN_NAME)
  → build/index.js (esbuild bundle)
    → index.ts start()
      → buildProgram() via Commander.js
        → command groups (project, account, env, teamsapp, add, list, m365, permission, entraApp, regenerate, misc)
          → wrapHandler(commandName, handler) for telemetry + error handling
          → wrapHandlerWithContext(name, handler) for handlers needing AtkContext
                                                    ↳ registerBuiltinDrivers() — deferred until a real command runs
```

## Bundling (esbuild)

Production builds use esbuild (`esbuild.mjs`) to bundle the CLI into a single CJS file (`build/index.js`).
This replaces the old `packages/cli` webpack approach (which needed 4GB heap and 100+ lines of config).

| Script | Command | Purpose |
|--------|---------|---------|
| `build` | `rimraf build && tsc -p ./` | Dev build — type checking + declarations |
| `bundle` | `node esbuild.mjs` | Dev bundle — fast, no minification |
| `package` | `rimraf build && tsc -p ./ && node esbuild.mjs --production` | Prod bundle — minified, source maps |
| `prepack` | `npm run package` | Auto-runs before `npm pack` / publish |

**Key esbuild settings:**
- Entry: `src/index.ts` → `build/index.js` (single file)
- Platform: `node`, target: `node18`, format: `cjs`
- `keepNames: true` — error class names appear in telemetry, never mangle them
- Source maps enabled, metafile for bundle analysis (`build/meta.json`)
- Externals: `keytar`, `@azure/msal-node-extensions` (native .node addons), `applicationinsights` (dynamic requires)

**Lazy-loading patterns** (to minimise `--help` startup time):
- `applicationinsights` — loaded via `require()` inside `AppInsightsTransport.init()`, not at module level
- `registerBuiltinDrivers()` — deferred from `start()` to `wrapHandlerWithContext()` (only when a real command runs)
- `node-machine-id` — loaded via `require()` inside `CliTelemetryReporter.init()`

### Context Layer (`src/context.ts`)

`createCliContext(projectPath?)` — creates an `AtkContext` wiring:
- `logger` → `CLILogProvider`
- `telemetry` → adapter wrapping `CliTelemetryReporter`
- `ui` → `CLIUserInteraction`
- `auth` → real MSAL-based `TokenProvider` via `createTokenProvider()` from `src/auth/`
- `correlationId` → `crypto.randomUUID()`

### Auth Module (`src/auth/`)

Real MSAL-based authentication ported from v3 CLI (`packages/cli/src/commonlib/`):

| File | Purpose |
|------|---------|
| `constants.ts` | Shared auth constants (client IDs, authority URLs, scopes) |
| `utils.ts` | Token parsing, online check, JWT claim extraction |
| `cacheAccess.ts` | AES-256-GCM encrypted MSAL token cache with optional keytar |
| `codeFlowLogin.ts` | MSAL interactive (auth code + browser) + silent token engine |
| `m365Login.ts` | `M365TokenProvider` implementation (M365 + Graph + App Studio scopes) |
| `azureLogin.ts` | `AzureAccountProvider` with interactive login + subscription listing |
| `azureLoginCI.ts` | Service principal auth for CI/headless environments |
| `index.ts` | `createTokenProvider()` factory — detects CI mode, creates appropriate providers |

- Token cache shared at `~/.fx/account/` for v3→v4 migration compatibility
- `keytar` is an optional dependency — graceful fallback to unencrypted cache
- Dependencies: `@azure/msal-node`, `@azure/identity`, `@azure/arm-subscriptions`, `async-mutex`, `open`

### Command Factory (`src/commands/factory.ts`)

Registry-driven subcommand generation — the extensibility core:

- `buildNewCommands(parent, registry, opts?)` — reads `TemplateRegistry`, groups by category,
  creates Commander subcommands: `atk new da basic`, `atk new bot echo`, etc.
- `mapQuestionToOption(spec)` — converts `QuestionSpec` → Commander `Option` using
  `cliName`, `cliShortName`, `cliDescription`, `isBoolean` from the Question interface
- Category slug mapping (configurable): `declarative-agent→da`, `custom-engine-agent→cea`,
  `ai-agent→ai`, `message-extension→me`, `bot→bot`, `tab→tab`, `connector→connector`, `office-addin→addin`
- Adding a new template to the registry automatically creates a new CLI subcommand

### Actions Layer (`src/actions/`)

Pure async functions that bridge CLI options to core-next operations (testable without Commander):

| Action | File | Core-next operation |
|--------|------|--------------------|
| `createProjectAction` | `actions/createProject.ts` | `runOperation(createProjectOp)` |
| `provisionAction` | `actions/lifecycle.ts` | `runOperation(provisionOp)` |
| `deployAction` | `actions/lifecycle.ts` | `runOperation(deployOp)` |
| `publishAction` | `actions/lifecycle.ts` | `runOperation(publishOp)` |
| `envListAction` | `actions/environment.ts` | `environment.listEnvironments()` |
| `envAddAction` | `actions/environment.ts` | `environment.addEnvironment()` |
| `envResetAction` | `actions/environment.ts` | `environment.resetEnvironment()` |
| `validateAction` | `actions/teamsapp.ts` | `runOperation(validateManifestOp \| validateAppPackageOp)` |
| `packageAction` | `actions/teamsapp.ts` | `runOperation(packageAppOp)` |
| `listTemplatesAction` | `actions/listTemplates.ts` | `registry.list()` → table rows |
| `addActionAction` | `actions/addAction.ts` | `runOperation(addExistingPluginOp)` |
| `addMCPActionAction` | `actions/addAction.ts` | `runOperation(addMCPActionOp)` |
| `addCapabilityAction` | `actions/addCapability.ts` | `runOperation(addKnowledgeOp)` |
| `addAuthConfigAction` | `actions/addAuthConfig.ts` | `runOperation(injectOAuthActionOp \| injectApiKeyActionOp)` |
| `setSensitivityLabelAction` | `actions/setSensitivityLabel.ts` | `runOperation(setSensitivityLabelOp)` |
| `m365SideloadAction` | `actions/m365Sideload.ts` | `runOperation(extendToM365Op)` |

## Key Differences from v3

| Aspect | v3 (`packages/cli`) | v4 (`packages/cli-next`) |
|--------|---------------------|-------------------------|
| Arg parser | Custom recursive-descent | Commander.js |
| Command model | Declarative `CLICommand` objects | Commander.js `.command()` / `.option()` / `.action()` |
| Dependencies | `@microsoft/teamsfx-api` + `@microsoft/teamsfx-core` | `@microsoft/teamsfx-core` v4 (merged API) |
| Error types | Same names, imported from `@microsoft/teamsfx-api` | Same names, imported from `@microsoft/teamsfx-core` |
| Prompt library | `@inquirer/prompts` (custom wrappers) | `@inquirer/prompts` (wired: select, checkbox, input, confirm) |
| Result pattern | `neverthrow` via `@microsoft/teamsfx-api` | `neverthrow` via `@microsoft/teamsfx-core` |

## Package Layout

```
packages/cli-next/
  src/
    index.ts          — Entry point, program parse
    context.ts        — createCliContext() factory (AtkContext for CLI, real auth providers)
    handler.ts        — wrapHandler(), wrapHandlerWithContext(), renderPostActions()
    error.ts          — CLI error classes (UserError/SystemError)
    logger.ts         — CLILogProvider (LogProvider impl)
    auth/              — Real MSAL-based auth providers
      constants.ts       — Client IDs, authority URLs, scopes
      utils.ts           — Token parsing, JWT utilities
      cacheAccess.ts     — AES-256-GCM encrypted token cache (optional keytar)
      codeFlowLogin.ts   — MSAL interactive + silent token engine
      m365Login.ts       — M365TokenProvider implementation
      azureLogin.ts      — AzureAccountProvider (interactive + subscription listing)
      azureLoginCI.ts    — Service principal auth (CI/headless)
      index.ts           — createTokenProvider() factory
    actions/           — Pure action functions (testable without Commander)
      createProject.ts   — createProjectAction(ctx, input)
      lifecycle.ts       — provisionAction, deployAction, publishAction
      environment.ts     — envListAction, envAddAction, envResetAction
      teamsapp.ts        — validateAction, packageAction
      listTemplates.ts   — listTemplatesAction(registry)
      addAction.ts       — addActionAction (wires to addExistingPluginOp), addMCPActionAction (wires to addMCPActionOp)
      addCapability.ts   — addCapabilityAction (wires to addKnowledgeOp)
      addAuthConfig.ts   — addAuthConfigAction (wires to injectOAuth/ApiKeyOps)
      setSensitivityLabel.ts — setSensitivityLabelAction (wires to setSensitivityLabelOp)
      m365Sideload.ts    — m365SideloadAction (wires to extendToM365Op)
    commands/
      index.ts        — buildProgram() assembles all Commander.js commands
      factory.ts      — buildNewCommands() registry-driven subcommand generator
      project.ts      — new (factory-driven), provision, deploy, publish, preview, upgrade
      account.ts      — auth show/login/logout
      env.ts          — env add/list/reset (wired to actions)
      teamsapp.ts     — validate/package (wired to actions), publish/update/doctor
      add.ts          — add action/capability/auth-config (wired to DA ops; action routes via --api-plugin-type: api-spec→addActionAction, mcp→addMCPActionAction)
      list.ts         — list templates (wired to action) /samples
      m365.ts         — m365-sideload (wired to extendToM365Op) /unacquire/launch-info
      permission.ts   — permission grant/status
      entraApp.ts     — entra-app update
      regenerate.ts   — regenerate/regenerate action
      misc.ts         — validate, set (sensitivityLabel wired), share, init
    output/
      colorize.ts     — chalk-based 9-TextType color system
      formatter.ts    — JSON + table (cli-table3) output
    telemetry/
      index.ts               — CliTelemetryReporter singleton (lazy App Insights init)
      appInsightsTransport.ts — Thin wrapper around applicationinsights TelemetryClient
      sanitize.ts            — PII redaction (file paths, tokens, emails)
    ui/
      userInteraction.ts — CLIUserInteraction (interactive + non-interactive)
```

## Telemetry (v4)

Lazy-initialised Application Insights transport — **no-op until `aiKey` is present** in `package.json`.
The `applicationinsights` module itself is also lazy-loaded (via `require()` inside `init()`) to avoid
pulling its transitive dependencies at CLI startup.

```
start() reads package.json → cliTelemetry.init(aiKey, version)
  → AppInsightsTransport.init(key, commonProperties)
    → require("applicationinsights") ← lazy, only when aiKey is present
    → appInsights.setup().setAutoCollect*(false).start()
  → every command: wrapHandler() → sendEvent("cmd-start") / sendEvent("cmd-end") / sendErrorEvent("cmd-error")
  → sanitizeProperties() strips tokens, passwords, emails, user file paths
  → flush() in handler finally block
```

**Key files:**

| File | Purpose |
|------|---------|
| `telemetry/index.ts` | `CliTelemetryReporter` — lazy init, shared properties, debug mode via `TEAMSFX_TELEMETRY_TEST` |
| `telemetry/appInsightsTransport.ts` | `AppInsightsTransport` — lazy-loads `applicationinsights` via `require()` in `init()`, auto-collection disabled, disk retry caching |
| `telemetry/sanitize.ts` | `anonymizeFilePaths()`, `sanitizeProperties()` — PII redaction before sending |
| `context.ts` | `telemetryAdapter` bridges `CliTelemetryReporter` → core's `TelemetryReporter` interface |
| `handler.ts` | `wrapHandler()` — auto-instruments every command with start/end/error events + duration |

**Activation gate:**
- Dev/local: `aiKey` is a dev instrumentation key in `package.json`
- Production: CD workflow (`cd.yml`) replaces `aiKey` with `${{ secrets.CLI_PUBLIC_AIKEY }}`
- No key (empty string): all methods stay no-ops — identical to previous stub behaviour

**Common properties** (set on the App Insights client): `common.os`, `common.platformversion`, `common.cliversion`, `common.machineid`

**Dependencies:** `applicationinsights@^1.8.10`, `node-machine-id@^1.1.12`

## Binary Names

- `atk` — primary entry point
- `teamsapp` — deprecated (shows deprecation warning)
- `m365agentstoolkit-cli` — long-form alias

## Adding a Command (v4)

1. Add a new function in the appropriate `src/commands/<group>.ts` file
2. Use `program.command("name").description("...").option("--flag", "...").action(wrapHandler("name", handler))`
3. Register in `buildProgram()` in `src/commands/index.ts`
4. The handler receives Commander.js parsed options — no custom parser logic needed

## Testing (v4)

Test infrastructure mirrors fx-core/cli pattern (Mocha + Chai + Sinon + NYC):

| Config | Purpose |
|--------|---------|
| `.mocharc.js` | ts-node/register, spec reporter, no-experimental-strip-types |
| `.nycrc` | 40% coverage threshold, excludes cli.js/cliold.js + appInsightsTransport.ts |

**Test file naming:** `*.test.ts` (consistent with all v4 packages).

**Test layout:**

```
tests/
  unit/
    error.test.ts       — MissingRequiredOptionError, InvalidChoiceError, etc.
    handler.test.ts     — wrapHandler() telemetry start/end/error, flush on error, exitCode
    output.test.ts      — colorize() TextType, printResult() JSON/table/empty
    commands.test.ts    — buildProgram() subcommands, global options, version flag
    context.test.ts     — createCliContext(): projectPath, correlationId, wiring
    factory.test.ts     — buildNewCommands(): category subcommands, slug mapping, template
                           options, --language omission, question→option mapping, extensibility proof
    actions.test.ts     — listTemplatesAction, createProjectAction, envListAction, envAddAction,
                           envResetAction (uses real core-next env operations on temp dirs)
    da-actions.test.ts  — addActionAction, addCapabilityAction, addAuthConfigAction,
                           setSensitivityLabelAction, m365SideloadAction (error paths + signatures)
  integration/
    commandExecution.test.ts      — --version parse, command groups, global options, error class integration
    coreNextIntegration.test.ts   — Real core-next ops through Commander: env list/add/reset,
                                     list templates, new command tree inspection (da/bot subcommands)
    addCommands.test.ts           — add/set/m365 command tree inspection, real DA ops via parseAsync
                                     (web-search capability, sensitivityLabel), required option verification
```

**Test commands:**

```bash
cd packages/cli-next
npm run test:unit           # 78 unit tests (with NYC coverage)
npm run test:integration    # 62 integration tests
npm run test                # Alias for test:unit
```

**Testing the `wrapHandler()` pattern:**

```typescript
import * as sinon from "sinon";
import { cliTelemetry } from "../../src/telemetry";
import { wrapHandler } from "../../src/handler";

const sandbox = sinon.createSandbox();
sandbox.stub(cliTelemetry, "sendEvent");
sandbox.stub(cliTelemetry, "sendErrorEvent");
sandbox.stub(cliTelemetry, "flush").resolves();

// Create a mock Commander Command for testing
const cmd = new Command("test-cmd");
cmd.setOptionValue("output", "text");
const wrapped = wrapHandler("mycommand", handler);
await wrapped(cmd);
```
