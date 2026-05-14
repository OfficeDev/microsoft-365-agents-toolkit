# Microsoft 365 Agents Toolkit — Documentation

Welcome. This is the **engineering and product documentation** for the Microsoft 365 Agents Toolkit monorepo. End-user product documentation lives at <https://aka.ms/teamsfx-docs>; this site covers what the product *is*, how it is *built*, and how it *evolves*.

## How this site is organised

The structure follows a **lifecycle spine** — product intent flows downward into concrete code:

```
Requirements (inform specs, can drift)
  01-product/  — WHY: PRD, personas, scenarios, success metrics
  01-product/ux/  — HOW users interact: surfaces, flows, question model

Constraints (specs must respect these)
  02-architecture/  — system structure, ADRs
  03-infrastructure/  — platform boundaries, Azure limits

Contracts (source of truth, enforced by tests)
  04-specs/          — domain specs + operation specs + data model entities
       ↓ derives
  Tests (packages/*/tests/)
       ↓ constrains
  Code (packages/*/src/)
```

It is loosely aligned to two well-known frameworks:

- **[arc42](https://docs.arc42.org/)** — for the architecture spine (sections 02 / 09)
- **[Diátaxis](https://diataxis.fr/)** — for separating *explanation*, *reference*, *how-to*, and *tutorial* within each section

The full rationale lives in [00-overview/docs-design-principles.md](00-overview/docs-design-principles.md).

## Map

| # | Section | What lives here |
|---|---------|-----------------|
| 00 | [Overview](00-overview/README.md) | Vision, glossary, roadmap pointer |
| 01 | [Product (PRD + UX)](01-product/README.md) | Personas, scenarios, capability matrix, success metrics; UX flows and surfaces under `01-product/ux/` |
| 02 | [Architecture](02-architecture/README.md) | arc42 sections 1–11, including ADRs |
| 03 | [Infrastructure](03-infrastructure/README.md) | Azure topology the toolkit provisions, Bicep, identity, secrets, CI/CD emitted |
| 04 | [Specs](04-specs/README.md) | Domain specs (boundaries + inter-domain contracts), operation specs (AC tables → tests), data model entities |
| 05 | [Engineering](05-engineering/README.md) | Monorepo layout, package reference, build/bundling, testing, cross-cutting modules |
| 06 | [Operations](06-operations/README.md) | Branching, CI pipelines, E2E strategy, publishing, versioning |
| 07 | [Contributing](07-contributing/README.md) | Docs style guide and contributor playbooks |
| 08 | [Troubleshooting](08-troubleshooting/README.md) | Known errors and mitigations |
| — | [`_v3-reference/`](_v3-reference/README.md) | **Quarantined** — deep extractions of v3 internal shapes for archival reference. **Forbidden as input to v4 design.** |

## Two lenses on the same code

Two generations of the engine coexist behind a feature flag:

| Lens | Packages | Status |
|------|----------|--------|
| **v3 (current)** | `api`, `manifest`, `fx-core`, `cli`, `vscode-extension`, `server` | Shipping |
| **v4 (next)** | `core-next`, `cli-next` | Preview, gated by `TEAMSFX_V4_CORE` |

Whenever a topic differs between v3 and v4, pages call this out explicitly using a *v3 / v4* split.

**Designing v4? Start here:** [05-engineering/v4-design-strategy.md](05-engineering/v4-design-strategy.md). The strategy document maps PRD intent to v4 architectural decisions and links the relevant ADRs and skills.

## Authoritative sources for contributors

These two locations remain the **machine-readable source of truth** that this site explains in prose. When the prose disagrees with them, they win:

- `.github/instructions/*.instructions.md` — coding conventions and per-package rules consumed by Copilot
- `.github/skills/*` and `.agents/skills/*` — workflow skills (dev/test/lint/e2e/etc.)
- `.dev/features.json` + `packages/fx-core/tests/integration/featureRegistry.ts` — the canonical feature registry

## How to contribute to these docs

See [07-contributing/docs-style-guide.md](07-contributing/docs-style-guide.md). In short: prefer linking to source over duplicating it; keep diagrams in [`assets/diagrams/`](assets/diagrams/README.md); use Mermaid; write Diátaxis-style (one purpose per page).
