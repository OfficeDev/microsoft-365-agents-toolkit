# Specs

Authoritative behavioral contract layer for the Microsoft 365 Agents Toolkit.
Tests are derived 1:1 from spec acceptance-criteria rows; implementation is
written to make those tests green.

## Layer hierarchy

```
PRD (docs/01-product/prd/)
  └─ Scenario (docs/01-product/scenarios/<group>/)         ← user journey, engine-neutral
      ├─ Domain Spec (docs/03-specs/domains/<nn>-<domain>.md)
      │   └─ Operation Spec (docs/03-specs/operations/<domain>/<operation>.md)
      │       └─ Acceptance Criteria table   ← horizontal: one atomic action, template-agnostic
      │           └─ Tests (1:1 with AC rows, name carries AC-ID)
      │               └─ Code (implementation makes failing tests green)
      └─ Scenario Spec (docs/03-specs/scenarios/<group>/<slug>.md)
          └─ Acceptance Criteria table   ← vertical: one template end-to-end, composes Operations
              └─ Tests (scenario tier; assert the scaffolded artifacts)
```

Architectural decisions that span multiple specs live as ADRs under
[`docs/02-architecture/`](../02-architecture/README.md). Data contracts and
entities are defined inline in the owning operation spec's inputs/outputs tables;
a consumer links to the owner rather than restating the shape (single source of
truth without a separate entity layer).

## Spec kinds

| Kind | Path pattern | Purpose |
|------|--------------|---------|
| Domain Spec | `domains/<nn>-<domain>.md` | Boundary, vocabulary, and rules for one of the capability domains. |
| Operation Spec | `operations/<domain>/<operation>.md` | One atomic engine action: inputs, outputs, AC table, flow, boundary, invariants. **Horizontal** — template-agnostic. |
| Scenario Spec | `scenarios/<group>/<slug>.md` | One template end-to-end: the concrete artifacts a single template produces, as AC rows that **compose** Operation Specs (referenced, never restated). **Vertical** — per-template. Drives scenario-tier (ADR-0018 T3) tests. |

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

## Required sections in a Scenario Spec

A scenario spec is the *vertical* counterpart to an operation spec: it pins what
**one template** produces end-to-end. It is **complete** only when all of these
are filled:

- header metadata — `Status`, `Domain`, the product `Scenario ID` (`SCN-…`) it
  mirrors, and the **template id** it validates.
- `## Acceptance Criteria` — one row per concrete, template-specific output fact
  (the produced files, the manifest values, the env-var names), tagged with a
  runtime tier (L1/L2/L3).
- `## Composed operations` — the Operation Specs this scenario flows through,
  **linked, not restated**. This is the anti-duplication seam: mechanism lives in
  the operation spec; only the template-specific facts live here.
- `## Flow` — Mermaid for the end-to-end scaffold (may reference the product
  scenario's flow rather than redraw it).
- `## Boundary` — what this scenario does NOT assert (every cross-template
  mechanism, which belongs to the composed operation specs).

## Operation Spec vs Scenario Spec — orthogonal cuts, not duplication

Two spec kinds carry AC tables, on **perpendicular axes** — keeping them separate
is what stops one from restating the other:

- An **Operation Spec** is *horizontal*: one atomic engine action
  (`resolve-template-source`, `run-scaffold-pipeline`), **template-agnostic**,
  exercised by every template that flows through it. Its AC protect the action's
  contract.
- A **Scenario Spec** is *vertical*: one template **end-to-end** (the
  `da/mcp-server` create scenario), composing those operations and pinning the
  **concrete** artifacts *that* template produces (the `ai-plugin.json`
  namespace, the `m365agents.yml` `oauth/register` block, the `MCP_DA_AUTH_ID_*`
  env var). Its AC protect the template's output.

They never restate each other: a scenario spec **references** the operation specs
it composes and adds only the template-specific facts no operation spec knows.
The two axes feed two test tiers — operation AC → operation-integration tests
(per action); scenario AC → scenario-tier tests (per template scaffold,
ADR-0018 T3).

## Test tiers — what each protects

Two independent axes describe a test. Keep them separate
([ADR-0013](../02-architecture/adr/ADR-0013-test-tiering-and-coverage-gate.md)).

**Runtime axis (L1/L2/L3)** — where it runs and how fast. Each AC row is tagged
with one:

- **L1** — in-memory through an injected port (`ScaffoldRuntime`, or an
  operation's port); no real fs / network / process. **Both**
  `operation-integration` **and** `scenario` (ADR-0018 T3) run here.
- **L2** — CLI E2E. Documented now; not yet a hard PR gate.
- **L3** — VS Code UI. Documented now; not yet a hard PR gate.

**Purpose axis — what the test protects.** This is what decides whether an
uncovered line needs a test at all. Two tiers are **AC-derived** but on
**perpendicular cuts** — `operation-integration` is *horizontal* (one atomic
action, from an Operation Spec) and `scenario` is *vertical* (one template
end-to-end, from a Scenario Spec) — so they are named separately and never
collapsed into a single "integration" bucket. The surface tier is **not**
AC-derived from this folder; it sources from a **product** scenario:

| Purpose tier | Cut / source | Protects | Runtime | When to write it |
|---|---|---|---|---|
| **operation-integration** (AC-derived) | *horizontal* — an Operation Spec AC row | one atomic-action behavior, run through the operation's port with in-memory fakes | L1 | **Always** — one per AC row; test name carries the AC ID. The primary protected tier. |
| **scenario** (AC-derived, ADR-0018 T3) | *vertical* — a Scenario Spec AC row | one whole template scaffolded end-to-end through `InMemoryRuntime`; asserts the produced artifacts. **Composes** the operations, never restates them | **L1** (engine-through, *not* a real surface) | **Always** — one per scenario AC row. |
| **file-unit** (pure) | not AC-derived | an intricate pure module's internal logic (parser, semver/range, digest) | L1 | **Optional** — only when the logic is genuinely complex. Never to chase a line number. |
| **CLI-E2E / UI** (surface) | a **product** scenario ([`01-product/scenarios`](../01-product/scenarios/)) | cross-surface behavior — the CLI flag tree, the VS Code Quick Pick / input / CodeLens | L2 / L3 | Documented now, progressively gated later. Highest project value. Traces to the product scenario, **not** a `03-specs` AC row. |

**Decision rule for an uncovered line:**

- Uncovered **behavior** → add its AC-derived test: an **operation** AC row
  (atomic action) or a **scenario** AC row (template output), whichever owns it.
- Uncovered **complex pure logic** → optionally add a file-unit test.
- Uncovered **thin adapter / glue / barrel** → cover it with *one* real
  integration test across the real boundary (temp dir; stub only at the network
  edge), **or** exclude it with `/* istanbul ignore next -- <reason> */`
  carrying a one-line reason. Never back-fill with mock-heavy micro-units whose
  only purpose is to lift a file's line %.

The coverage floor is a backstop, not the definition of done. "Done" is: every
AC row has its AC-derived test (operation-integration or scenario), and every
excluded line carries a justified ignore.

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
