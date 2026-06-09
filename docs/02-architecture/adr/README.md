# Architecture Decision Records (ADRs)

Numbered, dated, immutable records of architectural decisions for the
Microsoft 365 Agents Toolkit engine.

See [`../README.md`](../README.md) for what counts as an architectural
decision. The format for a new ADR is defined inline below under
[Adding a new ADR](#adding-a-new-adr).

## Status legend

- **Proposed** — open backlog item. Problem stated, options listed, decision
  pending. Safe to edit until status changes.
- **Accepted** — decided. **Immutable**. To revisit, write a new ADR and mark
  the old one `Superseded by ADR-NNNN`.
- **Superseded** — replaced by a newer ADR. Kept for history.

## Index

| ID | Title | Status | Triggered by |
|---|---|---|---|
| ADR-0001 | [Engine ↔ surface token-provider contract](ADR-0001-engine-surface-token-boundary.md) | Proposed | [`identity-and-login.md`](../external-dependencies/identity-and-login.md) (removed §1.8) |
| ADR-0002 | [M365 login plumbing: share vs duplicate across surfaces](ADR-0002-m365-login-plumbing-sharing.md) | Proposed | [`identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs) |
| ADR-0003 | [Native broker (WAM) gating policy](ADR-0003-broker-gating.md) | Proposed | [`identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs) |
| ADR-0004 | [TDP region state: singleton vs request-scoped](ADR-0004-tdp-region-state.md) | Proposed | [`identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs) |
| ADR-0005 | [Env-override configuration model](ADR-0005-env-override-config-model.md) | Proposed | [`identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs) |

## Adding a new ADR

1. Create `ADR-NNNN-<slug>.md` (next free number) using the template below.
2. Fill in **Status** (`Proposed`), **Date**, **Source**, **Context**, and
   **Options considered**. Leave **Decision** and **Consequences** as
   `(Pending.)` until the decision is made.
3. Add a row to the Index above.
4. Link the ADR from the fact page / open question that triggered it.

## Template

Copy everything inside the fenced block into a new `ADR-NNNN-<slug>.md`:

````markdown
# ADR-NNNN — <title>

- **Status:** Proposed
- **Date:** YYYY-MM-DD
- **Source:** <link to the fact page / open question that triggered this ADR>

## Context

What is the problem? Cite the relevant fact page under
[`../external-dependencies/`](../external-dependencies/README.md) — or the
architecture page — that triggered this decision, rather than restating its
content. Some ADRs are forced by external facts; others arise from purely
internal concerns (composition pattern, error model, module boundaries). Both
are valid; name which kind this is. Keep this to the *why* of the decision —
do not describe the current code.

## Options considered

- **A —** …
- **B —** …
- **C —** …

## Decision

(Pending. Filled in when status moves to `Accepted`.)

## Consequences

(Pending. Filled in when status moves to `Accepted`. List any new constraints
this decision introduces; add them to the relevant fact page or architecture
page in the same PR.)
````
