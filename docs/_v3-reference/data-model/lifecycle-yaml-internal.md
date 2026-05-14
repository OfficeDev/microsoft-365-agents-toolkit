# v3 lifecycle YAML — internal model

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

How `m365agents.yml` is parsed and represented in memory by v3. Source: [`packages/fx-core/src/component/configManager/`](../../../packages/fx-core/src/component/configManager/).

## On-disk shape

YAML with a fixed set of lifecycle sections. Each section is an array of "driver use" entries.

```yaml
version: 1.0
environmentFolderPath: ./env

provision:
  - uses: aadApp/create
    with:
      name: my-bot-aad
      generateClientSecret: true
    writeToEnvironmentFile:
      clientId: AAD_APP_CLIENT_ID
      clientSecret: SECRET_AAD_APP_CLIENT_SECRET
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID
  - uses: arm/deploy
    with:
      ...

configureApp:
  - ...

deploy:
  - ...

publish:
  - ...

share:
  - ...
```

## Lifecycle names

Tuple literal in source:

```typescript
type LifecycleNames = [
  "registerApp",
  "configureApp",
  "provision",
  "deploy",
  "publish",
  "share"
];
type LifecycleName = LifecycleNames[number];
```

## In-memory types

| Type | Source file | Purpose |
|------|-------------|---------|
| `RawProjectModel` | `configManager/types.ts` | YAML parse result. `version`, `environmentFolderPath`, `additionalMetadata`, plus one array per lifecycle name (raw `DriverDefinition[]`). |
| `ProjectModel` | `configManager/types.ts` | Resolved model. Same shape but each lifecycle holds an `ILifecycle` instance. Adds `aadPermission`. |
| `DriverDefinition` | `configManager/types.ts` | `{ name?, uses, with, env, writeToEnvironmentFile }` |
| `DriverInstance` | `configManager/types.ts` | `DriverDefinition` + `instance: StepDriver` |
| `ExecutionOutput` | `configManager/types.ts` | `Map<string, string>` — env vars after a driver runs |
| `ExecutionResult` | `configManager/types.ts` | `{ result: Result<ExecutionOutput, ExecutionError>, summaries: string[] }` |
| `ExecutionError` | `configManager/types.ts` | Discriminated union: `{ kind: "PartialSuccess"; env; reason }` or `{ kind: "Failure"; error }` |
| `PartialSuccessReason` | `configManager/types.ts` | Driver error or unresolved placeholders |
| `UnresolvedPlaceholders` | `configManager/types.ts` | Set of `${{VAR}}` keys that couldn't be resolved |

## `ILifecycle` interface

Source: [`packages/fx-core/src/component/configManager/interface.ts`](../../../packages/fx-core/src/component/configManager/interface.ts).

```typescript
interface ILifecycle {
  name: LifecycleName;
  driverDefs: DriverDefinition[];
  resolvePlaceholders(): UnresolvedPlaceholders;
  execute(ctx: DriverContext): Promise<ExecutionResult>;
  resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError>;
}
```

## Notes

- Placeholder resolution happens twice: at YAML parse time (string interpolation in `with:`) and at execution time (env vars injected into `process.env` for external file resolution).
- `writeToEnvironmentFile` maps driver output keys to env var names — the executor then routes secret-keyed values to `.user` env files.
- Driver outputs are `Map<string, string>` — no typed shape per driver.

## Why this is forbidden as v4 design input

v4 redesigns this layer:

- Driver inputs are pre-validated with **Zod** schemas (typed per driver, errors point at the YAML field).
- Driver outputs are typed per driver (no `Map<string, string>` boilerplate).
- The lifecycle engine separates parsing, resolution, analysis, and execution into distinct functions (`parseProjectYaml`, `resolveLifecycle`, `analyzeSteps`, `executeLifecycle`) instead of an `ILifecycle` interface with all four behaviours.
- `analyzeSteps` introspects driver IDs to determine prerequisites (M365 auth, Azure auth) — no manual ordering.

The v3 driver IDs (e.g. `aadApp/create`, `arm/deploy`) **are** preserved in v4 for `m365agents.yml` backward compatibility — that's the only allowed carry-over from this layer.
