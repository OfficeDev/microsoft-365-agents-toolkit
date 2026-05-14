# `Coordinator` class — v3 orchestration

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

Source: [`packages/fx-core/src/component/coordinator/index.ts`](../../../packages/fx-core/src/component/coordinator/index.ts).

The Coordinator orchestrates the lifecycle workflows (create, provision, deploy, publish) plus auxiliary VS-specific operations. It is decorated with lifecycle hooks and telemetry.

## Public methods

```typescript
class Coordinator {
  // Project creation
  async create(
    context: Context,
    inputs: Inputs,
    actionContext?: ActionContext
  ): Promise<Result<CreateProjectResult, FxError>>;

  // Project setup helpers
  async ensureTeamsFxInCsproj(projectPath: string): Promise<Result<undefined, FxError>>;
  async ensureTrackingId(...): Promise<...>;

  // Visual Studio pre-checks
  async preProvisionForVS(...): Promise<Result<PreProvisionResForVS, FxError>>;
  async preCheckYmlAndEnvForVS(...): Promise<Result<undefined, FxError>>;

  // Lifecycle stages (the heart of the class)
  async provision(...): Promise<...>;
  async deploy(...): Promise<...>;
  async publish(...): Promise<...>;
  async publishInDeveloperPortal(...): Promise<...>;

  // Result-shape utility
  convertExecuteResult(...): ...;
}
```

A `share()` method exists in source but is commented out (line 832). It is referenced in the v3 share commands but the public surface routes through `FxCore.shareApplication()` instead.

## Notes

- Coordinator holds internal mutable state (selected generator, action context).
- Generator activation in `create()` is **first-wins** — order in `generatorProvider.ts` matters; collisions are silent.
- Telemetry is wired via decorators; consumers cannot easily disable it for tests.
- `FxCore` delegates lifecycle calls (`provisionResources`, `deployArtifacts`, `publishApplication`) to the Coordinator's `provision/deploy/publish` methods after preprocessing. Both layers are part of the public-by-import surface.

## Why this is forbidden as v4 design input

v4 replaces this with the `Operation` pipeline (each lifecycle stage is its own operation), the `TemplateRegistry` (no implicit ordering — explicit lookup), the `DriverRegistry` (drivers are first-class and validated), and the lifecycle engine (`parseProjectYaml` → `resolveLifecycle` → `executeLifecycle`).
