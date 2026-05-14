# `packages/api` — `@microsoft/teamsfx-api`

The **public contract surface** for the v3 toolkit. Every other v3 package depends on it.

> **v4 note.** All contracts have been merged into `packages/core-next/src/api/` and republish as `@microsoft/teamsfx-core` v4. v3 consumers stay on `@microsoft/teamsfx-api`. Keep both in sync until v3 is deprecated.

## Conventions source

[`.github/instructions/api.instructions.md`](../../../.github/instructions/api.instructions.md).

## What lives here

- `FxError` interface and concrete `UserError` / `SystemError` classes.
- `Result<T, FxError>` (re-exported from `neverthrow`).
- Question model: `IQTreeNode`, `UserInputQuestion`, validation types.
- CLI command option types: `CLICommandOption`, `CLIBooleanOption`, `CLIStringOption`, `CLIArrayOption`.
- `LogProvider`, `TelemetryReporter`, `TokenProvider`, `UserInteraction` interfaces.
- Manifest type re-exports from `@microsoft/app-manifest`.

## Re-exports

`index.ts` re-exports curated dependencies so consumers have one import surface:

```ts
export * from "neverthrow";              // Result, ok(), err()
export * from "@microsoft/app-manifest"; // Manifest types
```

Adding a new re-export is a public API change — be deliberate.

## Backward compatibility rules

- New properties: optional (`?:`) only.
- Removed members: keep with `@deprecated` for at least one minor version.
- Discriminated unions for new variants — never widen existing required field types.
- Keep `error.name` strings stable.

## Testing

Tests verify interface contracts implicitly via type checking. Error serialisation/deserialisation is explicitly tested. New interfaces require shape-construction tests.
