# 0005 — Use `AsyncLocalStorage` for correlation-ID propagation in v4

- **Status:** Accepted
- **Date:** 2025
- **Context tags:** v4 / telemetry

## Context

In v3, correlation IDs propagated via `Correlator.run(id, async () => { ... })`, which works but requires every entry point to remember to wrap its handler. Internal helpers that emit telemetry on a "best-effort" basis miss the ID when invoked outside that wrapper.

Forces:

- Every telemetry event tied to a single user action must share one ID for end-to-end tracing.
- v4 introduces `AtkContext` injection — but threading `correlationId` through every function as a parameter is noisy.
- Async stack traces in Node 16+ are stable enough for `AsyncLocalStorage` to be reliable.

## Decision

Implement correlation as `AsyncLocalStorage<{ correlationId: string }>` in `core-next/src/telemetry/`:

- `correlationScope(id, async () => { ... })` runs an async function inside a scope.
- `getCurrentCorrelationId()` reads the current scope's ID; returns `undefined` outside any scope.
- Operations enter a scope when invoked through `runOperation()`; the ID is also placed on `AtkContext.correlationId` for explicit access.

Consumers (cli-next handlers) call `correlationScope(crypto.randomUUID(), ...)` once per user action. Drivers, clients, and helpers read the ID via `getCurrentCorrelationId()` without parameter threading.

## Consequences

- **Positive:** Telemetry coverage is automatic — even helpers that emit events outside the explicit `Operation` pipeline pick up the right ID.
- **Positive:** No churn in driver / client signatures.
- **Negative:** `AsyncLocalStorage` has a small runtime cost. Negligible for our event volumes.
- **Negative:** Code reading the ID outside a scope must handle `undefined`. We document this; tests assert it.

## Alternatives considered

- **Continue with `Correlator.run` from v3.** Rejected: ergonomic regression; v4 should improve here.
- **Pass `correlationId` as a parameter everywhere.** Rejected: every signature gains a parameter.
- **Use OpenTelemetry context API.** Rejected: overkill; we don't currently emit OTEL traces.

## References

- [`packages/core-next/src/telemetry/`](../../../packages/core-next/src/telemetry/)
- Node docs: <https://nodejs.org/api/async_context.html>
