# Architecture

Engine-neutral architecture documents for the Microsoft 365 Agents Toolkit. This
folder answers a single question:

> *"How is the engine itself built — what shape must any new piece of code fit
> into, and what cross-cutting contracts does it have to honor?"*

Architecture is engine-internal: it describes the toolkit's own codebase. The
runtime substrate the toolkit deploys *to* lives in
[`docs/03-infrastructure/`](../03-infrastructure/README.md). Business behavior —
what a feature does from the user's point of view — lives in
[`docs/01-product/`](../01-product/README.md) and is contracted in
[`docs/04-specs/`](../04-specs/README.md).

## What lives here

Three kinds of content, in this order:

1. **Chosen structure** — package boundaries, layering, dependency direction,
   composition rules, lifecycle composition. The deliberate shape of the engine.
2. **Cross-cutting contracts** — error model, logging / telemetry boundary,
   configuration model, context propagation, public vs internal API surface, i18n.
   Things every package must agree on for the engine to work as one system.
3. **Resulting constraints** — rules new code must follow because of (1) and (2)
   (e.g. "no cross-package internal imports", "no new module-scoped runtime
   state"). Constraints are the *output* of architectural decisions, not the
   definition of architecture.

Binding decisions are recorded as **ADRs** (numbered, dated, immutable). Every
constraint in this folder should be traceable to an ADR that explains *why*.

## What does NOT live here

- Per-feature behavior, AC tables, operation specs, data-model entities → [`docs/04-specs/`](../04-specs/README.md).
- Provisioning / deployment / runtime topology → [`docs/03-infrastructure/`](../03-infrastructure/README.md).
- Product intent, scenarios, surface UX → [`docs/01-product/`](../01-product/README.md).
- Per-package coding conventions (file headers, lint rules, naming) → [`.github/instructions/`](../../.github/instructions/).

## Conventions

- ADRs are Markdown, English, numbered sequentially, and never edited after they
  are accepted. To change an accepted decision, add a new ADR whose status is
  `Accepted` and update the old one's status to `Superseded by ADR-NNNN`.
- Architecture pages reference ADRs by number rather than restating them.
- Specific architectural choices (e.g. composition pattern, error type, registry
  vs class hierarchy) belong in ADRs, not in this README — this page is a
  contents page, not a decision log.

## Status

This folder is being populated. Until subpages exist, individual ADRs are the
authoritative source as they land; where no ADR exists yet, the corresponding
decision is treated as open.
