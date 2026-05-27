# ADR-0003 — Native broker (WAM) gating policy

- **Status:** Proposed
- **Date:** 2026-05-27
- **Source:** [`../external-dependencies/identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs)

## Context

The native Windows broker (WAM) is loaded on a different policy per surface
(see fact-page §1.3):

- **VS Code:** off by default; opt-in via the `TEAMSFX_BROKER_AUTH` feature flag.
- **CLI:** on by default whenever the platform is Windows and the plugin loads; silent fallback otherwise.

This asymmetry was not designed; it grew. Question: should VS Code also
default to on (with an off-switch for the rare regression), or should the CLI
also be opt-in (consistency at the cost of UX), or is the asymmetry actually
correct given each surface's risk tolerance?

## Options considered

- **A —** VS Code defaults to on; both surfaces share the same flag for opt-out.
- **B —** Both surfaces opt-in via the same flag (CLI behavior changes).
- **C —** Status quo, but document the asymmetry as deliberate with a justification per surface.

## Decision

(Pending.)

## Consequences

(Pending.)
