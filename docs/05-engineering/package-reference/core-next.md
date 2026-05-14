# `packages/core-next` — `@microsoft/teamsfx-core-next` (v4)

The **next-generation engine**. Replaces both `packages/api` and `packages/fx-core`. Publishes as `@microsoft/teamsfx-core` v4.0.0.

## Conventions source

[`.github/instructions/fx-core.instructions.md`](../../../.github/instructions/fx-core.instructions.md) §"core-next (v4 Successor)".

## Layout

```
packages/core-next/src/
├── api/                 — merged contracts (error, types, context, qm, UI, ...)
├── core/                — AtkContext, Operation, AtkError, defineOperation, runOperation
├── declarativeAgent/    — full DA module
├── lifecycle/           — YAML lifecycle engine + provisionOp / deployOp / publishOp
├── project/             — createProjectOp + createProjectInteractive
├── questions/           — question tree infrastructure
├── environment/         — env file CRUD (pure functions)
├── teamsApp/            — validate / package / publish operations
├── templates/           — TemplateRegistry + descriptors + scaffold pipeline + openApi adapter
├── helpers/             — metaOSHelper (DA upgrade)
├── specParser/          — inline OpenAPI parse / validate / filter / optimize
├── drivers/             — DriverRegistry + createDriver + 22 built-in drivers
├── telemetry/           — DI-first helpers, instrumentOperation, correlation
├── secretMasker/        — keyword-based credential masking
├── featureFlags/        — injectable FeatureFlagRegistry
├── localization/        — Localizer class (package.nls.json)
├── http/                — createHttpClient with telemetry interceptors, retry, timeout
├── clients/             — TeamsDevPortal / Graph / Azure ARM / M365 PackageService
├── folder.ts            — getTemplatesFolder() — bundled templates dir resolver
└── index.ts             — public barrel
```

## Key types

| Type | Source | Purpose |
|------|--------|---------|
| `AtkContext` | `core/AtkContext.ts` | Injected context (logger, telemetry, ui, auth, correlationId) |
| `AtkError` | `core/AtkError.ts` | Error type extending `FxError` |
| `Operation<I, O>` | `core/Operation.ts` | Orchestration unit; `defineOperation(name, schema, fn)` |
| `TemplateDescriptor` | `templates/types.ts` | Registry record |
| `DriverDescriptor` | `drivers/types.ts` | Driver record produced by `createDriver()` |

## Key operations

| Op | Source |
|----|--------|
| `createProjectOp`, `createProjectInteractive` | `project/` |
| `provisionOp`, `deployOp`, `publishOp` | `lifecycle/operations.ts` |
| `validateManifestOp`, `packageAppOp`, `publishAppOp` | `teamsApp/` |
| `addKnowledgeOp`, `addActionOp`, `addMCPActionOp`, `setSensitivityLabelOp`, `setConversationStartersOp` | `declarativeAgent/operations.ts` |

## Built-in drivers (22)

See [04-specs/data-model/entities/driver-descriptor.md](../../04-specs/data-model/entities/driver-descriptor.md) and [cross-cutting/driver-system.md](../cross-cutting/driver-system.md). Categories: `file/`, `script`, `cli/`, `teamsApp/`, `aadApp/`, `botAadApp/`, `botFramework/`, `arm/`, `azureAppService/`, `azureFunctions/`, `oauth/`, `apiKey/`.

## Templates

24 built-in `TemplateDescriptor`s registered via `registerBuiltinTemplates()`:

- DA (11), Bot (1), Tab (1), AI Agent (3), Engine Agent (3), Connector (1), Message Extension (1), OpenAPI (3), Foundry (1).

Bundled fallback ZIPs ship in `packages/core-next/templates/fallback/`.

## Service clients

| Client | Purpose |
|--------|---------|
| `TeamsDevPortalClient` | App CRUD, manifest validation, OAuth, API keys |
| `GraphApiClient` | Entra ID app reg + `/beta/appCatalogs/teamsApps` |
| `AzureArmClient` | ARM deploy, Kudu zip deploy |
| `M365PackageService` | V1 / V2 sideloading |

## Build & test

```bash
cd packages/core-next
npm run build               # tsc + postbuild (eslint --fix + prettier --write)
npm run test:unit           # 606 tests with NYC (80% gate)
npm run test:integration    # 48 tests
```

CI: `.github/workflows/ci-next.yml`. Use `pnpm --filter ./packages/core-next` (directory path — required because the npm package name `@microsoft/teamsfx-core` collides with v3).
