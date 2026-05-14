# Operation record

An `Operation` is the v4 unit of orchestration. Created via `defineOperation()`, executed via `runOperation()`.

Source: [`packages/core-next/src/core/Operation.ts`](../../../packages/core-next/src/core/Operation.ts), [`defineOperation.ts`](../../../packages/core-next/src/core/defineOperation.ts).

## Shape

```typescript
type Operation<TInput, TOutput> = {
  name: string;                              // telemetry-stable identifier
  inputSchema: ZodSchema<TInput>;
  execute: (input: TInput, ctx: AtkContext) => Promise<Result<TOutput, AtkError>>;
};
```

## Authoring

```typescript
import { defineOperation } from "../core/defineOperation";
import { z } from "zod";

export const provisionOp = defineOperation(
  "provision",                                 // name
  z.object({                                   // input schema
    projectPath: z.string(),
    envName: z.string(),
    skipConsent: z.boolean().optional(),
  }),
  async (input, ctx) => {                      // execute
    const env = await loadEnv(input.envName);
    if (env.isErr()) return err(env.error);
    // ... etc.
    return ok({ postActions: [...] });
  },
);
```

`defineOperation` takes **3 positional args** (name, schema, executeFn) — not an object.

## What `runOperation()` adds for free

| Behaviour | How |
|-----------|-----|
| Correlation scope | Wraps the call in `correlationScope(uuid, ...)` if not already inside one |
| Start / end telemetry | `instrumentOperation()` emits `<name>-start`, `<name>-end` |
| Error normalisation | Converts thrown exceptions to `AtkError` system errors |
| Input validation | Zod-validates against `inputSchema` before calling `execute` |
| Result return | Always returns `Result<TOutput, AtkError>` |

## Built-in operations

Lifecycle (v4):

- `provisionOp` — full provision pipeline
- `deployOp` — full deploy pipeline
- `publishOp` — full publish pipeline

Project (v4):

- `createProjectOp` — direct (non-interactive) scaffold
- `createProjectInteractive` — wraps the question tree + `createProjectOp`

DA (v4):

- `addKnowledgeOp`, `addActionOp`, `addMCPActionOp`, `setSensitivityLabelOp`, `setConversationStartersOp`

Teams app (v4):

- `validateManifestOp`, `packageAppOp`, `publishAppOp`

Environment (v4):

- Pure functions in `environment/envManager.ts`: `listEnvironments`, `readEnvFile`, `writeEnvFile`, `addEnvironment`, `resetEnvironment`. Not wrapped as `Operation` records — they are pure and don't need orchestration.

## Compose-ability

Operations are composable: an outer operation can call `runOperation(innerOp, innerInput)`. Telemetry events nest naturally because correlation IDs are inherited via `AsyncLocalStorage`.
