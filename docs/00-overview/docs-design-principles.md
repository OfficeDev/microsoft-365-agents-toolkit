# Docs design principles and references

This page records **why this site is shaped the way it is**. It is meta-documentation: read it before redesigning sections, splitting pages, or adding a new top-level area.

## Design principles

### P1 — Lifecycle spine

The top-level folder order mirrors how a product is built: **overview → product (PRD + UX) → architecture → infrastructure → specs → engineering → operations → contributing → troubleshooting**.

UX lives under `01-product/ux/` — not as a peer of architecture — because user flows are requirements input, not constraints on implementation. Specs (`04-specs/`) come after architecture and infrastructure because ADRs and platform limits *constrain* what specs can say; a spec cannot require behavior that contradicts a platform limit.

This matches the user's mental flow ("I want to understand *what* before *how*") and is stable — features come and go, but a product always has those layers.

Numeric prefixes (`00-`, `01-`, ..., `09-`) give a deterministic sort on every filesystem and a clear reading order, without relying on a doc-site renderer.

### P2 — One purpose per page (Diátaxis)

Every page is **one of**: tutorial, how-to, reference, or explanation — never a mix. If a page is doing two, split it.

| Quadrant | Where it lives here |
|----------|---------------------|
| Tutorial | rare; mostly belongs on learn.microsoft.com |
| How-to | [07-contributing/adding-*.md](../07-contributing/README.md), troubleshooting pages |
| Reference | [04-specs/data-model/](../04-specs/data-model/README.md), [05-engineering/package-reference/](../05-engineering/package-reference/README.md), [glossary](glossary.md) |
| Explanation | [02-architecture/](02-architecture/README.md), ADRs, [PRD overview](01-product/prd-overview.md) |

### P3 — Source-of-truth precedence

This site **explains** the codebase; it does not duplicate authoritative sources. When two locations disagree, the higher-precedence one wins:

1. The actual source code.
2. `.github/instructions/*.instructions.md` (Copilot conventions).
3. `.github/skills/*` and `.agents/skills/*` (workflow skills).
4. `.dev/features.json` + `featureRegistry.ts` (feature inventory).
5. This site (prose explanation).

A doc fix should chase the higher-precedence source first; only update prose when the underlying contract is correct.

### P4 — Two engines, one product

Whenever a topic differs between the **v3** engine (`fx-core` + `cli`) and the **v4** engine (`core-next` + `cli-next`), the page calls it out explicitly with a v3 / v4 split — usually a side-by-side table. The product behaviour the user sees is the same; the internals differ.

### P5 — Repository-relative everything

Links use workspace-relative paths so they resolve in GitHub, VS Code, and any markdown renderer. No `file://`, no `vscode://`, no absolute drive letters, no `aka.ms` for in-repo content (only for external help links the toolkit emits).

### P6 — Diagrams as text

Diagrams are **Mermaid** (and a few `.mmd` source files under [`assets/diagrams/`](assets/diagrams/README.md)). They diff cleanly in PRs, render on GitHub, and degrade gracefully to a code block. Every diagram pairs with a one-line text description for accessibility — Mermaid SVG has no semantic structure for screen readers.

### P7 — Stable error names

Error `name` strings are the partition key for telemetry and the join key for `helpLink` URLs. Documentation **never silently renames** them; if a name changes in code, it gets a deprecated alias and a doc note in the same PR.

### P8 — Update docs in the same PR as code

The [`docs-guard`](../.github/skills/docs-guard/SKILL.md) skill runs after implementation turns and adds doc-update tasks automatically. Doc drift is the failure mode this site is designed to prevent.

### P9 — English-only here

End-user docs at <https://aka.ms/teamsfx-docs> are localised through LocStudio across 13 languages. **This site is engineering documentation and stays English-only** — keeping it that way avoids a translation backlog that would otherwise outpace the engineering changes it tries to describe.

### P10 — No duplication of `.github/instructions/*`

Copilot consumes `.github/instructions/*.instructions.md` literally. This site **links to them** and explains them in prose — it never copies their content. When a convention changes, edit the instructions file; the linking page stays valid.

## References

### Primary frameworks

- **[Diátaxis](https://diataxis.fr/)** — the four-quadrant model (tutorials · how-to guides · reference · explanation). Used to keep one purpose per page.
- **[arc42](https://docs.arc42.org/)** — the 12-section architecture template. Drove the [02-architecture/](02-architecture/README.md) spine: introduction & goals (1), constraints (2), context & scope (3), solution strategy (4), building blocks (5), runtime views (6), deployment view (7), crosscutting concepts (8), ADRs (9), quality (10), risks (11).

### Supporting practices

- **[C4 model](https://c4model.com/)** (Simon Brown) — hierarchical diagrams: Level 1 (system context), Level 2 (containers / packages). Used in [02-architecture/05-building-blocks.md](02-architecture/05-building-blocks.md) and [02-architecture/03-context-and-scope.md](02-architecture/03-context-and-scope.md).
- **[ADRs — Michael Nygard format](https://github.com/joelparkerhenderson/architecture-decision-record/blob/main/locations/nygard/index.md)** — Status / Context / Decision / Consequences / Alternatives. Used in [02-architecture/09-architecture-decisions/](02-architecture/09-architecture-decisions/README.md).
- **PRD shape** — Atlassian / Marty Cagan-style structure (problem · solution · in/out of scope · success metrics · constraints · risks). Used in [01-product/prd-overview.md](01-product/prd-overview.md).
- **Personas + scenarios** — standard product-discovery pattern: a small set of personas, each scenario cross-references them.
- **Docs-as-code** — repository-as-source-of-truth, PR-reviewed prose, markdown over a heavier toolchain. Common in modern engineering orgs (e.g. GitLab handbook, Stripe docs).
- **Numeric folder prefixes** — common in arc42-influenced repositories for deterministic sort + obvious reading order.

### Deliberately not used

- **DITA / DocBook** — too heavyweight for an engineering-docs folder rendered on GitHub.
- **Sphinx-style `toctree`** — works for a published doc site, not for a folder of markdown viewed in-place.
- **TypeDoc generated reference** — belongs per-package alongside source; this site explains *why and how*, not generated `.d.ts` listings.
- **Wiki-style flat hierarchy** — loses the lifecycle spine that makes the table of contents predictable.

## When to revisit this page

Update this page when:

- A new top-level section is added or removed (the lifecycle spine changes).
- The source-of-truth precedence list changes (e.g. a new instruction-file system replaces `.github/instructions/`).
- A new diagram convention is adopted (e.g. moving from Mermaid to PlantUML).
- The v3/v4 split resolves (one engine remains).

For mechanical authoring rules — link format, Mermaid usage, voice — see the [docs style guide](../07-contributing/docs-style-guide.md).
