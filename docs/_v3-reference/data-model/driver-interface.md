# v3 driver interface — `StepDriver`

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

Source: [`packages/fx-core/src/component/driver/interface/`](../../../packages/fx-core/src/component/driver/interface/).

## `StepDriver`

```typescript
interface StepDriver {
  description?: string;
  progressTitle?: string;
  execute(
    args: unknown,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>,
    schemaVersion?: string,
    name?: string
  ): Promise<ExecutionResult>;
}
```

Notes:

- `args` is `unknown` — no compile-time validation against the YAML `with:` block.
- Driver outputs are `Map<string, string>` via `outputEnvVarNames` — driver authors must remember to wire this.
- Errors are returned via `ExecutionResult` — not exceptions, but also not typed per driver.
- No formal `validate()` separate from `execute()` — preflight relies on running the driver.

## `DriverContext`

```typescript
interface DriverContext {
  azureAccountProvider: AzureAccountProvider;
  m365TokenProvider: M365TokenProvider;
  ui: UserInteraction | undefined;
  progressBar: IProgressHandler | undefined;
  logProvider: LogProvider;
  telemetryReporter: TelemetryReporter;
  projectPath: string;
  platform: Platform;
}
```

Notes:

- `ui` and `progressBar` are nullable — drivers must check before calling.
- No `correlationId` field — correlation is wired through `Correlator.run()` wrappers higher up.
- The whole shape is a flatter version of v4's `AtkContext` but with TOOLS-derived providers.

## Why this is forbidden as v4 design input

v4's `createDriver()` factory enforces:

- **Zod pre-validation** on `args` — typo in YAML surfaces as a friendly `InvalidDriverInput` with the issue path, never reaches `execute`.
- **Typed inputs and outputs** — `execute(input: InputType, ctx: AtkContext): Promise<Result<OutputType, AtkError>>`.
- **Generated `validate()`** — preflight without running the driver, used by `analyzeSteps`.
- **Telemetry wrapping** — `driver-start` / `driver-end` events with timing, automatic.
- **Error normalisation** — thrown exceptions wrapped into `DriverExecutionError`; plain `AtkError` shapes recognised and unwrapped.
- **`AtkContext` injection** — same shape across all consumers (CLI, VS Code, tests), `correlationId` first-class.

Re-deriving these from v3's `StepDriver` would be a step backwards.
