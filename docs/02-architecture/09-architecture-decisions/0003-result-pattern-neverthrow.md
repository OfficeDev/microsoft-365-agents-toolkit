# 0003 — Adopt `Result<T, FxError>` via `neverthrow` everywhere

- **Status:** Accepted
- **Date:** 2022 (v3) · re-affirmed in v4
- **Context tags:** cross-cutting / errors

## Context

A toolkit that orchestrates many fallible operations (network, filesystem, cloud APIs, user input) must surface failures to users with help links and emit structured telemetry. Throwing exceptions makes both jobs harder: error sites lose type information, async stack traces are unreliable, and intermediate code must remember to wrap-and-rethrow with context.

Forces:

- Errors carry rich metadata (`source`, `name`, `helpLink`, `displayMessage`).
- Error names are stable identifiers in telemetry.
- TypeScript's `unknown`-typed `catch` clauses are awkward.
- Some failures are user-fixable; others are infrastructure problems — telemetry partitioning depends on this.

## Decision

Use `neverthrow`'s `Result<T, E>` as the universal return type for fallible operations. `E` is constrained to `FxError`. Concrete error classes:

- `UserError` — user-fixable. Renders with help link and action buttons.
- `SystemError` — infrastructure failure. Renders with "report issue" affordance.
- v4: `AtkError` extends the same shape.

Re-export `ok`, `err`, `Result` from `@microsoft/teamsfx-api` (v3) and `@microsoft/teamsfx-core` (v4) so consumers have a single import surface.

Throwing is reserved for **truly exceptional** programming errors (bugs, invariant violations) — never for expected user / network failures.

## Consequences

- **Positive:** Every fallible call site is forced to handle both branches by the type system.
- **Positive:** Telemetry consistently partitions UserError vs SystemError.
- **Positive:** Localisation pipeline (`displayMessage` vs `message`) is uniform.
- **Negative:** Boilerplate: `result.isErr() return err(result.error)`. Mitigated by `andThen`, `map`, `mapErr` chains.
- **Negative:** Async `Result` requires `ResultAsync` for nice chaining.

## Alternatives considered

- **`try`/`catch`.** Rejected: loses type safety, encourages over-broad catches.
- **`Promise<T>` with rejection conventions.** Rejected: rejection types are `unknown`; cannot enforce error metadata.
- **fp-ts `Either`.** Rejected: heavier learning curve; the team is more comfortable with `neverthrow`.

## References

- [`api.instructions.md`](../../../.github/instructions/api.instructions.md) §"Error Types"
- [`codebase.instructions.md`](../../../.github/instructions/codebase.instructions.md) §"Error Handling"
- [`neverthrow` library](https://github.com/supermacro/neverthrow)
