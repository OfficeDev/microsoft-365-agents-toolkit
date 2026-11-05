# ADR-0028 - Driver-owned args and shared execution context

- **Status:** Superseded by [ADR-0029](ADR-0029-minimal-public-action-boundary.md)
- **Date:** 2026-11-04
- **Source:** Approved follow-up request in chat.
- **Supersedes:** [ADR-0027](ADR-0027-standalone-action-client.md).

## Context

The standalone SDK should aggregate driver contracts, not own parameter types
that drivers must import back from the SDK registry. Action telemetry hooks
consume the supplied context; they do not inject dependencies into child services.

## Decision

- Drivers own pure args interfaces, schemas and validators alongside their
  implementation. The public allowlist aggregates those contracts and derives
  its exported argument types without importing live drivers.
- The existing snapshot validation and safe error formatting are shared below
  the drivers, not duplicated or imported from the SDK registry.
- `PackageService` accepts the existing `DriverContext`. Driver call sites pass
  that context directly instead of constructing a separate tools projection.
  HTTP and event telemetry consume the same context's reporter.
- An explicit context without a reporter never falls back to global Tools.
  Existing logger/token-provider constructor calls retain their legacy behavior.
  Existing context builders forward their experiment provider.
- All other execution boundaries and compatibility decisions in ADR-0027 are
  retained unchanged.

## Consequences

There is no new telemetry-tool abstraction or change to action execution behavior.
The acceptance criteria remain owned by
[execute-action-client](../../03-specs/operations/lifecycle/execute-action-client.md).
