# ADR-0002 — M365 login plumbing: share vs duplicate across surfaces

- **Status:** Proposed
- **Date:** 2026-05-27
- **Source:** [`../external-dependencies/identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs)

## Context

The CLI and VS Code surfaces each carry their own near-identical M365 login
implementation: same first-party AAD client ID, same MSAL
`PublicClientApplication` setup, same broker handling on Windows. The
duplication exists today; the question is whether it should be removed.

Removing it would reduce drift (a bug fixed in one copy lands in the other
for free) and shrink the constraint count from §2 of the fact page.
Preserving it keeps VS Code APIs out of the CLI module graph and lets each
surface evolve its UX (prompts, cache location, sign-out semantics)
independently.

## Options considered

- **A —** Extract MSAL plumbing into a new shared package; surfaces inject UX / storage hooks.
- **B —** Extract only the broker / cache helpers; keep `PublicClientApplication` wiring per surface.
- **C —** Status quo (deliberate duplication); add a parity test to prevent silent drift.

## Decision

(Pending.)

## Consequences

(Pending.)
