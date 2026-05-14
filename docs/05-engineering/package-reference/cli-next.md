# `packages/cli-next` — `@microsoft/m365agentstoolkit-cli-next` (v4)

The **v4 CLI** — same `atk` binary, esbuild-bundled, registry-driven, Commander.js-based, real MSAL auth.

## Conventions source

[`.github/instructions/cli.instructions.md`](../../../.github/instructions/cli.instructions.md) §"CLI v4".

## Architecture

```
cli.js (entry — sets TEAMSFX_CLI_BIN_NAME)
  → build/index.js (esbuild bundle)
    → index.ts start()
      → buildProgram() via Commander.js
        → command groups (project, account, env, teamsapp, add, list, m365, permission, entraApp, regenerate, misc)
          → wrapHandler(name, handler)              # telemetry + error
          → wrapHandlerWithContext(name, handler)   # also creates AtkContext
                                                   ↳ registerBuiltinDrivers() — deferred
```

## Layers

### Context (`src/context.ts`)

`createCliContext(projectPath?)` wires:

- `logger` → `CLILogProvider`
- `telemetry` → adapter wrapping `CliTelemetryReporter`
- `ui` → `CLIUserInteraction`
- `auth` → real MSAL `TokenProvider` via `createTokenProvider()`
- `correlationId` → `crypto.randomUUID()`

### Auth (`src/auth/`)

Real MSAL implementation, ported and refactored from v3 `cli/src/commonlib/`:

| File | Purpose |
|------|---------|
| `constants.ts` | Client IDs, authority URLs, scopes |
| `utils.ts` | Token parsing, online check, JWT extraction |
| `cacheAccess.ts` | AES-256-GCM cache with optional keytar |
| `codeFlowLogin.ts` | MSAL interactive + silent token engine |
| `m365Login.ts` | `M365TokenProvider` |
| `azureLogin.ts` | `AzureAccountProvider` with subscription listing |
| `azureLoginCI.ts` | Service principal for CI/headless |
| `index.ts` | `createTokenProvider()` factory |

Cache shared at `~/.fx/account/` with v3.

### Command factory (`src/commands/factory.ts`)

`buildNewCommands(parent, registry)` reads `TemplateRegistry` and generates Commander subcommands per category. Question metadata (`cliName`, `cliShortName`, `cliDescription`, `isBoolean`) maps to options via `mapQuestionToOption()`.

Category slugs (default): `da`, `cea`, `ai`, `me`, `bot`, `tab`, `connector`, `addin`.

### Actions (`src/actions/`)

Pure async functions bridging CLI to core-next:

- `createProjectAction` → `runOperation(createProjectOp)`
- `provisionAction` / `deployAction` / `publishAction` → `runOperation(provisionOp/deployOp/publishOp)`
- `envListAction` / `envAddAction` / `envResetAction` → `environment.*`
- `validateAction` / `packageAction` → `runOperation(validateManifestOp/packageAppOp)`
- `listTemplatesAction` → `registry.list()` formatted as table

### Handler wrappers

- `wrapHandler(name, handler)` — telemetry + error handling (no context)
- `wrapHandlerWithContext(name, handler)` — also creates `AtkContext`
- `renderPostActions(actions)` — displays `PostAction[]` (open URL, show message)

## Bundling

| Script | Command | Purpose |
|--------|---------|---------|
| `build` | `rimraf build && tsc -p ./` | Dev type check + declarations |
| `bundle` | `node esbuild.mjs` | Dev bundle (no minification) |
| `package` | `rimraf build && tsc -p ./ && node esbuild.mjs --production` | Prod bundle |
| `prepack` | `npm run package` | Auto before `npm pack` / publish |

esbuild settings: `platform: "node"`, `target: "node18"`, `format: "cjs"`, `keepNames: true`, externals `keytar` + `@azure/msal-node-extensions` + `applicationinsights`.

Lazy loads:

- `applicationinsights` — `require()` inside `AppInsightsTransport.init()`.
- `node-machine-id` — `require()` inside `CliTelemetryReporter.init()`.
- `registerBuiltinDrivers()` — deferred to first real command.

## Testing

```bash
cd packages/cli-next
npm run test:unit           # 87 tests
npm run test:integration    # 81 tests
```

ESLint flat config: `shared` + `header` (no `promise` — too many false positives in CLI stubs).

## Status

Phase 4b complete + CI. Lifecycle (provision/deploy/publish) wired with real auth and 22 registered drivers. 38/39 templates fully supported (only `typeSpec/compile` missing).

Feature flag: `TEAMSFX_V4_CORE` (default off).

CI: `ci-next.yml` (lint, format-check, unit-test, integration-test). E2E: `e2e-test-next.yml` (daily + PR + manual).
