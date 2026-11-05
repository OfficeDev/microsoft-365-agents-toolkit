# ADR-0029 - Minimal public action boundary

- **Status:** Accepted
- **Date:** 2026-11-04
- **Source:** Request to expose direct action execution with minimal existing-code changes.
- **Supersedes:** [ADR-0028](ADR-0028-driver-owned-action-contracts.md).

## Context

The requirement is an exported fx-core action invocation API for CLI consumers.
Reworking existing service construction, experimentation, HTTP telemetry and
shared context builders is outside that requirement.

## Decision

- Add a typed, allowlisted public execution entry and export the driver-owned
  args contracts. Keep runtime validation and metadata independent of live drivers.
- Build the invocation context inside the new entry from the injected Tools and
  explicit project path. Fix the platform to CLI and do not prompt for action
  arguments. The approved 2026-11-05 clarification permits interactive login:
  reuse the injected M365 token provider without overriding authentication requests.
- Reuse existing driver and service behavior. Do not change shared context
  builders, service constructors, telemetry routing or experiment-provider lookup.
  Their existing runtime semantics, including shared engine state, remain intact.
- Keep any required standalone output adaptation at the driver's execution entry,
  not in its shared business handler or in the generic dispatcher.
- The supported action list and compatibility criteria belong to the
  [operation spec](../../03-specs/operations/lifecycle/execute-action-client.md).
  Existing uninstall paths are unchanged.

## Consequences

This delivery does not introduce a new multi-instance isolation model for shared
services. Adding a direct invocation API must not become an engine-runtime refactor.
