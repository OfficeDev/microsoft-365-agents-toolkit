# ADR-0027 - Allowlisted standalone action execution

- **Status:** Superseded by [ADR-0028](ADR-0028-driver-owned-action-contracts.md)
- **Date:** 2026-11-03
- **Source:** Approved SDK handoff and implementation request in chat.

## Context

In-process callers need to execute one supported action without running a YAML
lifecycle. Lifecycle execution owns project discovery, placeholder expansion and
environment persistence; those behaviors do not belong in this SDK boundary.
The first consumers need ZIP installation and title-ID uninstallation only.
No CLI, WIQD, PRD or product-scenario change is needed in this delivery.

## Decision

- A closed, typed action catalogue owns argument contracts, static JSON Schema
  draft-07 definitions and cancellation capability. Drivers consume the same
  argument types. Metadata and validation have a driver-free public entry point.
- `FxCore` and `FxCoreClient` delegate to one executor. It validates an independent
  argument snapshot before lazily resolving an allowlisted driver and returns the
  driver's existing `ExecutionResult`, without business-specific result mapping.
- Tools belong to the instance; absolute `projectPath` belongs to the invocation.
  The context fixes `Platform.CLI`, disables UI and requests silent M365 tokens.
  Injected token providers must honor `showDialog: false`; their errors propagate.
- There is no lifecycle, YAML, placeholder, environment-loading or output-
  persistence step. Existing driver implementations and `PackageService` own
  business behavior. No arbitrary container or module resolution is exposed.
- Installation reuses `teamsApp/extendToM365`. Without an output mapping its
  outputs use business keys; supplied legacy mappings retain their semantics.
  `teamsApp/uninstallFromM365` is a thin title-ID-only driver.
- Cancellation is optional per action. Unsupported cancellation is rejected.
  Supported cancellation is cooperative, awaits started work, preserves execution
  errors and does not imply rollback.

## Consequences

Adding another standalone action requires an explicit contract, lazy driver
registration and tests, not executor branches. Other lifecycle drivers and CLI
commands are not automatically opened. Public subpath shims preserve existing
package deep imports without introducing a restrictive package exports map.

The behavioral contract and acceptance criteria belong to
[execute-action-client](../../03-specs/operations/lifecycle/execute-action-client.md).
