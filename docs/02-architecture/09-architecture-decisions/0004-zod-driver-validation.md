# 0004 — Pre-validate driver inputs with Zod in `createDriver()`

- **Status:** Accepted
- **Date:** 2025
- **Context tags:** v4 / drivers / validation

## Context

In v3, drivers received untyped input bags from the YAML lifecycle and validated them late, inconsistently, or not at all. Common failure modes:

- Typo in YAML key surfaces as a far-away `undefined` access.
- Required field missing produces a confusing error from inside an Azure SDK.
- No uniform place to attach JSON schema for editor IntelliSense.

Forces:

- v4 lifecycle YAML is user-editable — early, friendly errors matter.
- Driver outputs feed the env-var pipeline — output shape must be predictable.
- Drivers must remain easy to write — adding validation can't add 50 lines per driver.

## Decision

Make Zod validation a first-class part of the driver factory:

```typescript
createDriver({
  id: "azureFunctions/zipDeploy",
  name: "Zip-deploy to Azure Functions",
  inputSchema: z.object({
    workingDirectory: z.string(),
    artifactFolder: z.string(),
    resourceId: z.string(),
  }),
  execute: async (input, ctx) => { /* input is fully typed */ },
});
```

`createDriver()` automatically:

1. Pre-validates input against `inputSchema`. On failure, returns an `InvalidDriverInput` `AtkError` with the Zod issue path — never enters `execute`.
2. Generates a separate `validateFn` for preflight (used by `analyzeSteps`).
3. Wraps `execute` with start/end telemetry events and timing.
4. Catches unexpected throws and normalises them into `DriverExecutionError`.
5. Recognises plain `AtkError` shapes (with `code`/`message`/`kind`) and returns them unwrapped — prevents `[object Object]` serialisation.

## Consequences

- **Positive:** Errors point at the exact YAML field. Telemetry has `error.name` partitioning by validation vs execution.
- **Positive:** Driver authors get full TypeScript inference inside `execute`.
- **Positive:** Zod schemas can be exported to JSON Schema for editor support.
- **Negative:** Adds Zod to runtime deps. Acceptable: small footprint, no other runtime validator on the team.
- **Negative:** Schema and TypeScript type co-exist (`z.object` + `z.infer`). Manageable; eliminates drift.

## Alternatives considered

- **Hand-written validators.** Rejected: inconsistent, easy to forget, no reusable error shape.
- **JSON Schema with `ajv`.** Rejected: no inferred TypeScript types; worse DX in driver code.
- **`io-ts`.** Rejected: less readable; smaller community than Zod.

## References

- [`packages/core-next/src/drivers/createDriver.ts`](../../../packages/core-next/src/drivers/createDriver.ts)
- [`packages/core-next/src/drivers/builtin/`](../../../packages/core-next/src/drivers/builtin/)
- [`fx-core.instructions.md`](../../../.github/instructions/fx-core.instructions.md) §"Driver System"
