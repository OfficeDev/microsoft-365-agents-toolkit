# Contributing to docs

This file is the **docs-only** contributor entry point. For the project-wide contributor guide (CLA, Code of Conduct, PR conventions), see [`CONTRIBUTING.md`](../CONTRIBUTING.md) at the repo root.

## TL;DR

- Each page has **one purpose** ([Diátaxis](https://diataxis.fr/)). Split if it grows two.
- Don't duplicate `.github/instructions/*` — link to them.
- Mermaid for diagrams. Workspace-relative links for files. Backticks only for symbols.
- Update docs **in the same PR** as the code change they describe.

## Style

Full guide: [docs-style-guide.md](docs-style-guide.md).
Design principles and the frameworks they're drawn from: [../00-overview/docs-design-principles.md](../00-overview/docs-design-principles.md).

## Site structure

The folder structure follows a lifecycle spine: overview → product (PRD + UX) → architecture → infrastructure → specs → engineering → operations → contributing → troubleshooting. Each numbered folder has its own `README.md` index.

Top-level entry: [README.md](README.md).

## Source-of-truth precedence

The precedence depends on **direction of travel** — describing what exists vs. proposing what comes next — and on **which engine** is in scope.

### Backward — describing existing behaviour (default)

Applies to all v3 packages (`api`, `manifest`, `fx-core`, `cli`, `vscode-extension`, `server`) and to v4 packages when reverse-engineering shipped behaviour.

Highest first:

1. The actual source code.
2. `.github/instructions/*.instructions.md` (Copilot conventions).
3. `.github/skills/*/SKILL.md` and `.agents/skills/*/SKILL.md` (workflow skills).
4. `.dev/features.json` + `featureRegistry.ts` (feature inventory).
5. This site (prose explanation).

Doc fixes chase the higher-precedence source first; only update prose when the underlying contract is correct.

### Forward — proposing new v4 behaviour

Applies to **all non-trivial AI-driven changes to `core-next` and `cli-next`**.

Highest first:

1. The design page in [`02-architecture/`](../02-architecture/README.md) or [`04-specs/data-model/`](../04-specs/data-model/README.md), and any related ADR in [`02-architecture/09-architecture-decisions/`](../02-architecture/09-architecture-decisions/README.md).
2. Updates to [`01-product/capabilities-matrix.md`](01-product/capabilities-matrix.md) and [`.dev/features.json`](../.dev/features.json) for new capabilities.
3. The code that implements the design.
4. `.github/instructions/*.instructions.md` updates (if conventions change).
5. `.github/skills/*/SKILL.md` updates (if workflow changes).

The design page is authoritative until the code lands. Code that drifts from the design either updates the design (in the same PR) or doesn't merge. See [codebase.instructions.md §Source-of-Truth Workflow](../../.github/instructions/codebase.instructions.md) and [05-engineering/v3-to-v4-migration.md §Workflow asymmetry](../05-engineering/v3-to-v4-migration.md).

v3 packages are **not** design-first — pretending the new docs are authoritative for shipping v3 behaviour would silently lie to contributors.

### Forbidden inputs to v4 design

[`_v3-reference/`](_v3-reference/README.md) holds extracted v3 internal shapes (FxCore class, lifecycle YAML model, driver interface, error catalogue, generator categories, Bicep archetypes, naming conventions). It is for archival reference only and **must not** be used as input to v4 design. See [`_v3-reference/README.md`](_v3-reference/README.md) §"The hard rule" for the full list and rationale.

## Adding a new page

1. Pick the right section folder.
2. Create the page.
3. Link it from the section's `README.md`.
4. Run a quick read-through — does it have one purpose? Are links workspace-relative?

## Renaming or moving a page

Avoid if possible — internal links break silently in markdown. If you must:

1. Move the file.
2. `grep` the rest of the site for the old path; fix every reference.
3. Mention the rename in the PR description.

## Diagrams

- Inline Mermaid for short ones.
- `.mmd` source under [`assets/diagrams/`](assets/diagrams/README.md) for anything reused or non-trivial.
- Pair every diagram with a one-line text description for accessibility.

## Localisation

This site is **English only** (engineering docs). End-user docs at <https://aka.ms/teamsfx-docs> are localised through the LocStudio pipeline; docs here are not.
