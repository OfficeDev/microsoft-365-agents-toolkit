---
description: "Use when editing or creating interfaces, error types, question models, or contracts in the api package. Covers backward compatibility, interface evolution, re-exports, and Result pattern."
applyTo: "packages/api/**/*.ts"
---

# API Package Conventions

## Role

This is the **public contract surface** for the current (v3) toolkit. All other v3 packages depend on it. Changes here have the widest blast radius.

> **v4 Migration Note:** All contracts from this package have been merged into
> `packages/core-next/src/api/` (`@microsoft/teamsfx-core` v4.0.0). New v4 packages
> (`cli-next`, future `vscode-extension`) import from `@microsoft/teamsfx-core` instead.
> This package remains the source of truth for v3 packages (`fx-core`, `cli`, `server`,
> `vscode-extension`). Keep both in sync until v3 is fully deprecated.

## Backward Compatibility

- Mark removed or replaced members with `@deprecated` and a migration note before removing
- Add new functionality via **optional properties** (`?:`) — never change existing required properties
- Prefer **discriminated unions** over type assertions when adding new variants:

```typescript
// Good — extend with new variant
export type CLICommandOption = CLIBooleanOption | CLIStringOption | CLIArrayOption;

export interface CLIStringOption extends CLICommandOptionBase {
  type: "string";
  choices?: string[];  // optional extension
}

// Bad — change existing required field type
```

- When removing a re-exported symbol, keep the re-export with `@deprecated` for at least one minor version

## Re-Exports

`index.ts` re-exports key dependencies so consumers get a single import surface:

```typescript
export * from "neverthrow";              // Result, ok(), err()
export * from "@microsoft/app-manifest"; // Manifest types
```

- Consumers import `ok`, `err`, `Result`, manifest types from `@microsoft/teamsfx-api` — not from the underlying packages
- Adding a new re-export is a public API change — be deliberate

## Error Types

`FxError` is the base interface. Two concrete classes:

| Class | When | User sees |
|-------|------|-----------|
| `UserError` | Bad input, missing config, auth issues | Help link, action buttons |
| `SystemError` | Service failure, internal bugs | Issue link |

Constructor via options object:

```typescript
new UserError({
  source: string,       // component name (required)
  name?: string,        // stable error name for telemetry
  message?: string,     // English log message
  displayMessage?: string, // localized user-facing message
  helpLink?: string,    // aka.ms link
  error?: Error,        // wrapped inner error
});
```

- `name` appears in telemetry — keep it stable across releases
- `displayMessage` is what users see — always localized
- `message` is for logs — always English

## Question Model (QM)

Defines the interactive question tree used by VS Code, CLI, and Visual Studio:

- `BaseQuestion` → `UserInputQuestion` with `type` discriminator
- Question types: `singleSelect`, `multiSelect`, `text`, `singleFile`, `multiFile`, `folder`, `confirm`
- Dynamic options via `LocalFunc<T>` — computed from previous answers at runtime
- Validation via `StringValidation` (pattern, enum, minLength) or custom `ValidateFunc<T>`

Adding a new question type:
1. Add the type string to the union in `question.ts`
2. Define the interface extending `UserInputQuestion`
3. Implement rendering in CLI (`userInteraction.ts`) and VS Code (`qm/`)

## Interface Design Guidelines

- Prefer `unknown` over `any` for untyped data; narrow with type guards
- Use `readonly` for properties that should not be mutated after creation
- Document every public interface member with JSDoc — these are consumed by external contributors
- Group related interfaces in dedicated files; export via barrel in `qm/index.ts` or `utils/index.ts`

## Testing

- Tests verify interface contracts implicitly via type checking
- Error serialization/deserialization is explicitly tested
- When adding a new interface, add tests that construct and validate the shape
