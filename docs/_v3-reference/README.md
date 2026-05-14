# `_v3-reference/` — Quarantined v3 internal extractions

> **STOP. READ THIS BEFORE YOU READ ANYTHING ELSE IN THIS FOLDER.**

## What this folder is

This folder contains **deep extractions of v3 internal shapes** — the FxCore class signature, the YAML lifecycle in-memory model, the driver interface, the error catalogue, the QuestionNames enum, the templates' Bicep archetypes, naming conventions, and so on.

It exists for one purpose: **archival reference**. Auditors, historians, contributors trying to understand a v3 bug report, or anyone debugging a v3 install can look here to see what the v3 engine actually does internally.

## What this folder is NOT

It is **not** an input to v4 design.

The v4 engine (`packages/core-next` and `packages/cli-next`) is being designed clean-room from **product-level requirements** — the PRD, the features list, the UX. **Not** from how v3 is built internally. Importing v3 internals into v4 would re-import v3's accidental complexity (e.g. the `TOOLS` global singleton, implicit generator activation order, ad-hoc YAML driver dispatch, `FxCore` god-class with 50+ methods) and forfeit the entire point of the rewrite.

## The hard rule

When proposing a v4 change (`packages/core-next`, `packages/cli-next`):

| You may use as v4 design input | You may NOT use as v4 design input |
|--------------------------------|------------------------------------|
| [`docs/01-product/`](../01-product/README.md) — PRD, personas, capabilities, scenarios | Anything in `_v3-reference/` |
| [`docs/01-product/ux/`](../01-product/ux/README.md) — surfaces, flows, question model, errors | The `FxCore` class signature |
| Microsoft 365 platform contracts (manifest schemas, Graph, TDP) | The v3 lifecycle YAML internal model |
| Existing `docs/02-architecture/` ADRs | The v3 driver interface (`StepDriver`) |
| Existing `docs/04-specs/data-model/` and `docs/05-engineering/` v4 sections | The v3 generator activation pattern |
| Microsoft Foundry / Agents SDK / Teams AI library upstream contracts | The v3 error catalogue (names that are stable; *names only* may be reused for back-compat — not the catalogue's organising structure) |
| | Bicep templates from v3 (re-design from the topology requirements, not from existing `.bicep` files) |

If a v4 design copies a shape from this folder, that's a code review red flag. Reviewers should ask: *"What product requirement justifies this shape, independent of how v3 happens to do it?"*

## Why the asymmetry

| Direction | Source of truth |
|-----------|-----------------|
| Describing **v3 behaviour** | v3 source code → this folder → `docs/` prose |
| Describing **shipping product behaviour** (v3 ∪ v4) | source code → instructions → `docs/` prose |
| Designing **new v4 behaviour** | `docs/01-product/`, `docs/01-product/ux/`, ADRs in `docs/02-architecture/09-architecture-decisions/` |

This folder lives at the bottom of all three precedence chains for design — and is **explicitly excluded** from the third.

See:

- [codebase.instructions.md §"Source-of-Truth Workflow"](../../.github/instructions/codebase.instructions.md)
- [docs/07-contributing/docs-contributing.md §"Source-of-truth precedence"](../07-contributing/docs-contributing.md)
- [docs/05-engineering/v3-to-v4-migration.md §"Workflow asymmetry"](../05-engineering/v3-to-v4-migration.md)

## What's in here

```
_v3-reference/
├── README.md                          ← you are here
├── data-model/                        — v3 internal shapes
│   ├── README.md
│   ├── api-package.md                 — packages/api exports
│   ├── fxcore-class.md                — FxCore class signature (57 public methods)
│   ├── lifecycle-yaml-internal.md     — RawProjectModel / ProjectModel / DriverDefinition / ILifecycle
│   ├── driver-interface.md            — StepDriver + DriverContext
│   ├── error-catalog.md               — every error name in fx-core/src/error/
│   ├── question-names.md              — full QuestionNames enum
│   └── generator-categories.md        — fx-core/src/component/generator/ subdirs
└── infra/                             — v3 IaC archetypes
    ├── README.md
    ├── template-inventory.md          — every shipped template + its archetype
    ├── archetypes.md                  — Bot/AppService, Bot/SQL, Functions/API, Functions/Connector, Static Web App, Tab/AppService
    └── naming-conventions.md          — resourceBaseName, location handling, RG conventions
