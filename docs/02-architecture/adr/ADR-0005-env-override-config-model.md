# ADR-0005 — Env-override configuration model

- **Status:** Proposed
- **Date:** 2026-05-27
- **Source:** [`../external-dependencies/identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs)

## Context

`APP_STUDIO_ENV=int` (and similar one-off `process.env.*` reads scattered
across the codebase) switch internal endpoints / behaviors for dogfood. Each
one is a direct `process.env` access without a central registry, a
documented default, or schema validation. As more fact pages land it is
likely more such knobs will surface (`TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT`,
`TEAMSFX_BROKER_AUTH`, future ones).

Question: should env-driven knobs go through a single configuration surface
(typed, listed in one place, validated at startup), or should they stay as
ad-hoc reads?

## Options considered

- **A —** Central typed config object; all env reads migrate to it.
- **B —** Lightweight registry: name + default + read site, no behavior change.
- **C —** Status quo; add a doc page listing every supported env var.

## Decision

(Pending.)

## Consequences

(Pending.)
