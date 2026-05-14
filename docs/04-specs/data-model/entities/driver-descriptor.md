# Driver descriptor

A `DriverDescriptor` is what `createDriver()` returns — the v4 unit of lifecycle work. Source: [`packages/core-next/src/drivers/createDriver.ts`](../../../packages/core-next/src/drivers/createDriver.ts).

## Shape

```typescript
type DriverDescriptor = {
  id: string;                                       // YAML "uses:" key
  name: string;                                     // human-readable
  inputSchema: ZodSchema;                           // pre-validation
  execute: (input, ctx) => Promise<Result<DriverOutput, AtkError>>;
  validate: (input) => Result<void, AtkError>;     // generated from inputSchema
  rollback?: (input, ctx) => Promise<Result<void, AtkError>>;
};
```

## Authoring pattern

```typescript
import { z } from "zod";
import { createDriver } from "../createDriver";

export const myDriver = createDriver({
  id: "myNamespace/myAction",
  name: "Do the thing",
  inputSchema: z.object({
    workingDirectory: z.string(),
    threshold: z.number().min(1).max(100),
  }),
  execute: async (input, ctx) => {
    // input is fully typed
    // ctx is AtkContext (logger, telemetry, ui, auth, correlationId)
    // return ok({ ... outputs ... }) or err(new AtkError({...}))
  },
});
```

## What `createDriver()` adds for free

1. **Zod pre-validation** — runs before `execute`. On failure, returns `InvalidDriverInput` with the Zod issue path.
2. **Telemetry** — `driver-start` and `driver-end` events with timing and outcome.
3. **Error normalisation** — wraps thrown exceptions into `DriverExecutionError`. Recognises plain `AtkError` shapes (`code` / `message` / `kind` properties) and returns them unwrapped to prevent `[object Object]` serialisation.
4. **`validate` function** — separate preflight without `execute` side effects, used by `analyzeSteps`.

## Built-in drivers (v4, 22 today)

See [05-engineering/cross-cutting/driver-system.md](../../05-engineering/cross-cutting/driver-system.md) for the full catalogue. Categories:

- `file/*` — `createOrUpdateEnvironmentFile`, `createOrUpdateJsonFile`
- `script` — cross-platform shell with `::set-output` parsing
- `cli/*` — `runNpmCommand`, `runDotnetCommand`
- `teamsApp/*` — TDP and Graph operations
- `aadApp/*`, `botAadApp/create`, `botFramework/create` — identities
- `arm/deploy`, `azureAppService/zipDeploy`, `azureFunctions/zipDeploy` — Azure infra & deploy
- `oauth/register`, `apiKey/register` — auth registration

## v3 equivalents

In v3, drivers live under `packages/fx-core/src/component/driver/` and are not pre-validated. Driver IDs match between v3 and v4 so YAML stays compatible.

## Driver outputs

Drivers return `Result<DriverOutput, AtkError>`. `DriverOutput` is a key-value map; the executor:

1. Maps output keys to env var names via the YAML `writeToEnvironmentFile:` block.
2. Merges into the active `envMap` for use by subsequent steps.
3. `persistEnv` writes the final `envMap` to env files at the end of the operation.
