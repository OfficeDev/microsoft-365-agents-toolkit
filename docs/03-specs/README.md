# Specs

Authoritative behavioral contract layer for the Microsoft 365 Agents Toolkit.
Tests are derived 1:1 from spec acceptance-criteria rows; implementation is
written to make those tests green.

## Layer hierarchy

```
PRD (docs/01-product/prd/)
  └─ Scenario (docs/01-product/scenarios/)
      └─ Domain Spec (docs/03-specs/domains/<nn>-<domain>.md)
          └─ Operation Spec (docs/03-specs/operations/<domain>/<operation>.md)
              └─ Acceptance Criteria table   ← tests are derived from this table
                  └─ Tests (1:1 with AC rows, name carries AC-ID)
                      └─ Code (implementation makes failing tests green)
```

Architectural decisions that span multiple specs live as ADRs under
[`docs/02-architecture/`](../02-architecture/README.md). Data contracts and
entities live under `data-model/`.

## Spec kinds

| Kind | Path pattern | Purpose |
|------|--------------|---------|
| Domain Spec | `domains/<nn>-<domain>.md` | Boundary, vocabulary, and rules for one of the capability domains. |
| Operation Spec | `operations/<domain>/<operation>.md` | One atomic engine action: inputs, outputs, AC table, flow, boundary, invariants. |
| Data-model Entity | `data-model/entities/<entity>.md` | Stable contract for a domain entity (shape, identifiers, lifecycle). |

## Required sections in an Operation Spec

An operation spec is **complete** (eligible for tests/code) only when all of these
sections are filled:

- `## Acceptance Criteria` — ID-based table; one row per testable behavior.
- `## Flow` — Mermaid diagram for stateful or cross-step behavior.
- `## Boundary` — what the operation does NOT do.
- `## Invariants` — constraints that must never be violated.

If a section cannot be completed because of upstream ambiguity, stop and surface
the gap as a question to PM rather than guessing. See the
[`vibe-coding`](../../.github/skills/vibe-coding/SKILL.md) skill for the full
spec → tests → code gate.

## Glossary (authoritative for this repo)

- **Capability** — PM word for a feature area users perceive.
- **Domain** — engineering word for the same area; one domain spec per domain.
- **Operation** — one atomic engine action belonging to a domain.
- **Template** — composes Operations into a shippable starting point.
- **Driver** — implementation primitive used by Operations to interact with
  external systems (clouds, services, files). Its shape is governed by
  [`docs/02-architecture/`](../02-architecture/README.md).

## Status

This folder is being populated. Until per-domain and per-operation specs land, the
spec format above is the contract; treat the
[`vibe-coding`](../../.github/skills/vibe-coding/SKILL.md) skill as the operating
workflow.
