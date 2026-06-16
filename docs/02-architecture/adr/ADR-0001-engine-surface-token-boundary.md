# ADR-0001 — Engine ↔ surface token-provider contract

- **Status:** Proposed
- **Date:** 2026-05-27
- **Source:** [`../external-dependencies/identity-and-login.md`](../external-dependencies/identity-and-login.md)

## Context

This decision is internal-architecture, not an external fact — the boundary
shape between surfaces and the engine is a design choice, not something the
outside world forces on us. It is listed here because the identity fact page
is what surfaced the problem.

Surfaces (VS Code, CLI) acquire tokens; the engine consumes them. Today the
engine accepts two separate provider shapes — one for M365 user tokens, one
for Azure ARM credentials — and the two have incompatible error contracts:
one returns `Result<…, FxError>` from every method, the other throws on user
cancel / failure. Both grew in parallel rather than being designed as a pair.

The boundary also carries auxiliary concerns beyond token acquisition: tenant
switching, sign-out, status-change callbacks, subscription enumeration. Each
of these was added when a feature needed it, with no top-down decision about
whether the engine should depend on those shapes.

Open question: what is the engine-facing shape of "the thing that gives you a
token"? Is it one interface or two? Result-shaped or throw-shaped? How much
of the auxiliary surface does the engine actually need?

## Options considered

- **A —** Keep two providers; unify on `Result<…, FxError>`.
- **B —** Keep two providers; unify on throw-on-failure.
- **C —** Collapse to a single provider with two token methods.
- **D —** Reduce the engine-facing surface to a minimal `getToken(resource, scopes)` and move tenant / sign-out / status-change concerns back to the surface.

## Decision

(Pending.)

## Consequences

(Pending.)
