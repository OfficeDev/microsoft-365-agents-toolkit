# v4 design strategy

> **Status:** Strategy. **Audience:** Anyone designing or implementing v4 (`packages/core-next`, `packages/cli-next`).
>
> **Source-of-truth precedence for v4 work:** the [01-product/](../01-product/README.md) PRD + features + UX surfaces in [01-product/ux/](../01-product/ux/README.md). v3 internals under [`_v3-reference/`](../_v3-reference/README.md) are **explicitly excluded** as design input. See [docs/07-contributing/docs-contributing.md §"Source-of-truth precedence"](../07-contributing/docs-contributing.md).

This document maps the product requirements (the PRD and feature surface) to the v4 architectural decisions that satisfy them, and to the engineering culture v4 is being built under. It is the connective tissue between product intent and implementation.

The five things a reader should take away:

1. **v4 is design-first; v3 stays code-first.** New features in `core-next`/`cli-next` start with a docs change. New work in `fx-core`/`cli` continues to follow the v3 code-first convention.
2. **`cli-next` is the leading edge.** The CLI is where new v4 capabilities land first; the VS Code extension consumes `core-next` later through the same operations.
3. **Verification weight sits on E2E + integration tests, not unit tests.** This is an orchestration tool — most of the truth about whether it works lives at the integration boundary.
4. **Vibe coding is enabled by architecture, not slogans.** Registries, declarative descriptors, Zod-validated drivers, single-file-readable modules — these are the choices that make AI agents productive contributors with hard safety gates (tests, types, design-first).
5. **Docs are part of the build.** A v4 PR that changes behaviour without changing the relevant `docs/` page does not merge.

## Why v4 exists

v3 ships and works. The reasons to build v4 are not aesthetic — they are concrete blockers to the PRD:

| PRD intent | v3 limitation | v4 design choice |
|------------|---------------|-------------------|
| Reproducible runs in CI | `TOOLS` global singleton makes parallel runs and isolated tests impossible | `AtkContext` injected per operation; pure functions where possible |
| Day-1 platform support | Generators activate by implicit ordering — collisions are silent and shipping a new template requires touching the engine | `TemplateRegistry` + `TemplateDescriptor` records — adding a template adds a CLI subcommand and a question-tree branch automatically |
| Ship with confidence | YAML driver inputs are untyped — typos surface as `undefined` errors far from the YAML site | `createDriver()` factory with **Zod pre-validation**; `InvalidDriverInput` error names the field |
| Get started fast | CLI cold start is multi-second (webpack 4 GB heap, eager `applicationinsights`) | `cli-next` is esbuild single-file CJS bundle; `applicationinsights`, drivers, and `node-machine-id` are lazy-loaded. Target: `atk --help` < 200 ms |
| AI-assisted authoring | v3 surface is hard for LLMs to reason about (god class, implicit ordering, scattered manifest manipulation) | Registries + descriptors + `defineOperation()` + `Result<T, AtkError>` everywhere — the surface is enumerable, typed, and self-describing |
| Localisation in 13 languages | `getLocalizedString()` singleton is hard to mock or override | `Localizer` class is injectable through `AtkContext` |
| Telemetry with strict secret masking | v3 mixes the masker with `TOOLS`; hard to reason about per-event | `secretMasker/` is pure; HTTP clients sanitise URLs in interceptors |

If a v4 design proposal does not solve a real PRD-derived limitation, it is over-engineering. Reject.

## The v4 architectural pillars

The pillars below are the "why this shape" behind every operation, driver, and template descriptor in `core-next`. They are not negotiable per-feature — when a new design contradicts a pillar, the **design** changes, not the pillar.

### Pillar 1 — Registries everywhere, ordering nowhere

v3's deepest anti-pattern is **first-activated-wins** generator ordering. v4 replaces this with explicit registries:

| Registry | What it holds | Effect |
|----------|---------------|--------|
| [`TemplateRegistry`](../05-engineering/cross-cutting/template-system.md) | `TemplateDescriptor` records | Adding a template auto-generates a CLI subcommand + question-tree branch |
| [`DriverRegistry`](../05-engineering/cross-cutting/driver-system.md) | `DriverDescriptor` records (Zod-validated) | Lifecycle YAML resolves drivers by ID; collisions are detected at registration |
| `FeatureFlagRegistry` | `FeatureFlagDescriptor` records | Injectable for tests; defaults from `process.env` |

If a new feature wants to "register itself by being imported first" — that is the v3 anti-pattern. Use a registry.

### Pillar 2 — Descriptors over class hierarchies

A `TemplateDescriptor` is a record, not a class. A `DriverDescriptor` is a record, not a class. An `Operation` is a record, not a class. This makes them:

- **Enumerable** — listing, filtering, and codegen become trivial.
- **Diffable** — code review on a registry-add is a one-screen change.
- **AI-readable** — an LLM agent can reason about the entire registry by reading one file.

Inheritance is reserved for error types (`UserError`/`SystemError` extend `FxError`). Everywhere else: factory functions returning records.

### Pillar 3 — Errors as values; names are stable

Every fallible operation returns `Result<T, AtkError>`. `AtkError` carries `source`, `name`, `displayMessage`, `helpLink`, `innerError`. The `name` field is the telemetry partition key — it is stable across releases. Renaming is a deprecation event, not a refactor.

This is the same pattern v3 uses; v4 keeps it. See [02-architecture/09-architecture-decisions/0003-result-pattern-neverthrow.md](../02-architecture/09-architecture-decisions/0003-result-pattern-neverthrow.md).

### Pillar 4 — Driver inputs are Zod-validated before execution

Every `createDriver()` call declares an `inputSchema: ZodSchema`. The factory validates `with:` from YAML before calling `execute`. Failures surface as `InvalidDriverInput` with the exact issue path.

This eliminates a class of v3 bugs where a YAML typo surfaces as `Cannot read property 'x' of undefined` 200 lines deep into a driver. See [ADR 0004](../02-architecture/09-architecture-decisions/0004-zod-driver-validation.md).

### Pillar 5 — `AtkContext` injection; no module-scoped state

Every operation, driver, client, and helper takes `ctx: AtkContext` as the first or last argument (convention: last for drivers, first for operations). `ctx` carries `logger`, `telemetry`, `ui`, `auth`, `correlationId`, optional `projectPath`. There are **no** module-scoped singletons that hold runtime state.

For testing, [`createMockContext()`](../05-engineering/testing-strategy.md#v4-test-helper) returns a fully-stubbed context. There is one canonical way to construct a real context per surface (`createCliContext()` for cli-next; future `createVsCodeContext()` for the extension).

### Pillar 6 — Single-file-readable modules

A new contributor (or AI agent) should be able to understand a module by reading **one file** end to end. Every module's `index.ts` exports the public surface; per-feature files are < 300 lines where possible. When a file passes ~500 lines, that's a refactor signal — not a "we'll get to it" signal.

This is partly why v3's `FxCore.ts` (57 public methods on one class) is forbidden as v4 design input.

### Pillar 7 — Operations compose; lifecycle is data

The lifecycle (provision/deploy/publish/share) is not a chain of method calls — it is a YAML document parsed into a `RawProjectModel`, resolved against the `DriverRegistry`, and executed by [`executeLifecycle`](../05-engineering/cross-cutting/lifecycle-engine.md). Operations (`provisionOp`, `deployOp`, `publishOp`) compose this with prerequisites + env-persistence.

Adding a new lifecycle stage means: (a) define a driver, (b) compose an operation, (c) add a YAML section. No engine changes.

## cli-next is the leading edge

See [ADR 0006 — cli-next as the v4 reference adapter](../02-architecture/09-architecture-decisions/0006-cli-next-as-reference-adapter.md).

Three reasons:

1. **CLI is the smallest viable surface.** A new operation needs a flag-parser entry, an action function, and `wrapHandlerWithContext`. That's ~30 lines. The same operation in VS Code needs tree-view registration, command palette wiring, telemetry attribution, and a webview if interactive — easily 200+ lines per surface.
2. **CLI is what CI runs.** Every E2E test in `e2e-test-next.yml` is a `atk` invocation. If an operation works in cli-next, it has been exercised by the same code path CI uses. VS Code parity comes after.
3. **CLI is what AI agents drive.** GitHub Copilot Coding Agent, Claude Code, Cursor — they all shell out to CLIs. A clean cli-next surface is what makes the toolkit AI-driveable.

**Operational consequence:** new v4 features land in `core-next` operation + `cli-next` action + `cli-next` E2E test, in one PR. VS Code extension wiring follows in a separate PR after `TEAMSFX_V4_CORE` flips for the extension.

## E2E and integration are the primary verifications

See [ADR 0007 — Inverted test pyramid for v4 lifecycle code](../02-architecture/09-architecture-decisions/0007-inverted-test-pyramid-for-lifecycle.md).

Standard test-pyramid wisdom (lots of unit tests, few integration tests, very few E2E) does not fit a CLI orchestration tool well. Most of `core-next`'s code is glue: parsing YAML, calling Azure / Graph / TDP APIs, shelling out to npm. The **unit boundary is artificial** — a unit test of `provisionOp` with every driver mocked tests very little of what could break in production.

For v4 lifecycle code, the priority is:

| Layer | Weight | Where |
|-------|--------|-------|
| Pure-function unit tests | High | `secretMasker`, `featureFlags`, `localization`, `questions/treeBuilder`, `templates/scaffold/replaceMap`, `lifecycle/analyze` — anything with no I/O |
| Integration tests | **Highest** | `lifecycle/executor` ↔ real drivers ↔ stubbed services; `templates/scaffold` ↔ real ZIPs ↔ real fs |
| E2E tests | **Highest** | cli-next invoking real `atk` against real or sandbox tenants. Required gating before declaring a feature "shipped". |

This **does not** mean writing fewer tests — it means moving them down only when a unit test is **better** than the integration test it would replace, not just because the pyramid says so. Fowler's [practical test pyramid §"Avoid test duplication"](https://martinfowler.com/articles/practical-test-pyramid.html) §rule 1 still holds: if a high-level test catches an error and no lower-level test does, that gap needs filling.

**Operational consequence:**

- Every new template in `features.json` requires an E2E test before it can be marked "supported".
- Every new operation requires an integration test that runs the full pipeline against the in-memory test helper.
- A unit test that just re-mocks the same surface as the integration test is a code review delete signal.

## Vibe coding for the toolkit

See [ADR 0008 — Architecture choices that enable AI-agent contributions](../02-architecture/09-architecture-decisions/0008-vibe-coding-architecture.md) and the [`vibe-coding`](../../.github/skills/vibe-coding/SKILL.md) skill for the workflow.

"Vibe coding" in this codebase means **agent-driven implementation with hard safety gates** — never throwaway, never sandbox-only. The PRD ships in production, so every line of agent-written code goes through the same review, testing, and design-first process as human-written code.

What makes the v4 architecture vibe-coding-friendly:

| Architecture choice | Why it helps an agent |
|--------------------|----------------------|
| Registries everywhere | Agent can see the full feature surface in one file |
| `TemplateDescriptor` / `DriverDescriptor` records | Adding a feature = adding a record. No control-flow archaeology. |
| Zod input schemas on drivers | Agent gets typed `input` parameter; no guessing the shape |
| `Result<T, AtkError>` everywhere | Agent doesn't have to reason about thrown vs returned errors |
| `AtkContext` injection | Agent can write tests immediately with `createMockContext()` |
| Single-file-readable modules | Agent can fit a module in one read; no chasing imports |
| Source-of-truth precedence | Agent knows what to trust when sources disagree |
| Design-first workflow | Agent writes the design page first; the design is the spec it then implements |

What makes vibe coding **safe** in this codebase:

| Safety gate | Where enforced |
|-------------|----------------|
| Design page in same PR | [codebase.instructions.md §"Source-of-Truth Workflow"](../../.github/instructions/codebase.instructions.md) |
| Zod input validation on every driver | `createDriver()` factory |
| Typed errors with stable names | ESLint + code review |
| E2E test required for new features | `ci-next.yml` + `e2e-test-next.yml` |
| Integration test for every operation | Same |
| Lint + format + 80% coverage gate | `ci-next.yml` |
| Conventional Commits | commitlint hook |
| Source-of-truth precedence enforced by docs structure | [docs/07-contributing/docs-contributing.md](../07-contributing/docs-contributing.md) |
| `_v3-reference/` quarantine | [_v3-reference/README.md](../_v3-reference/README.md) |

When an agent (or a human) submits a PR that bypasses any of these — design page missing, no E2E, driver without Zod schema, error name not stable — review rejects it. The architecture *enables* fast agent contribution; the gates *ensure* the contribution is safe to merge.

Read the [`vibe-coding`](../../.github/skills/vibe-coding/SKILL.md) skill for the step-by-step workflow.

## Docs are part of the build

This is the most important meta-decision. v4 docs are **not** a parallel artefact maintained when someone has time. They are part of the contract:

- A new v4 feature requires a design-page update or new ADR **in the same PR** as the code.
- The [`docs-guard`](../../.github/skills/docs-guard/SKILL.md) skill nudges agents to do this automatically.
- Reviewers reject PRs that change behaviour without changing docs.
- The [features registry](../01-product/v3-feature-inventory.md#coverage-in-featuresjson) is consumed by integration tests — drift is caught at CI time.

This is what makes the v4 design-first claim real, not aspirational.

## What this strategy does NOT mean

- It does **not** mean v3 is deprecated. v3 ships, supports millions of installs, and gets bug fixes on the v3 line.
- It does **not** mean every v4 PR needs an ADR. Only architectural decisions need ADRs. Bug fixes, dependency bumps, doc-only changes do not.
- It does **not** mean unit tests are bad. Pure-function logic in `secretMasker`, `featureFlags`, `localization` is best-tested as units. The inversion only applies to lifecycle / orchestration code.
- It does **not** mean cli-next gets every feature first forever. Once `TEAMSFX_V4_CORE` flips for the VS Code extension, the surfaces co-evolve. Until then, cli-next leads.
- It does **not** sanction "vibe coding" in the throwaway-sandbox sense. Every PR is production-bound; every PR has the same gates.

## Open questions

These are intentionally unresolved at strategy time; ADRs to follow when they're decided:

- When does `TEAMSFX_V4_CORE` flip from default-off to default-on? Tracked in [02-architecture/09-architecture-decisions/0001-feature-flag-v4-core.md](../02-architecture/09-architecture-decisions/0001-feature-flag-v4-core.md).
- Does v4 ship a unified VS Code + CLI command surface? See open question in [01-product/ux/surfaces/copilot-chat-participant.md](../01-product/ux/surfaces/copilot-chat-participant.md).
- Do samples and templates unify under the same `TemplateRegistry`? See [01-product/ux/flows/sample-gallery-and-upgrade.md §"v4 design implication"](../01-product/ux/flows/sample-gallery-and-upgrade.md).
- Are the two bot output schemas (`AZURE_APP_SERVICE_RESOURCE_ID` vs `BOT_AZURE_APP_SERVICE_RESOURCE_ID`) converged in v4 templates? Tracked at [_v3-reference/infra/archetypes.md §"Bot on App Service with Identity"](../_v3-reference/infra/archetypes.md#archetype-bot-on-app-service-with-identity).

## Authoritative pointers

| Concern | Doc |
|---------|-----|
| PRD | [01-product/prd-overview.md](../01-product/prd-overview.md) |
| What v4 must build | [01-product/capabilities-matrix.md](../01-product/capabilities-matrix.md) + [01-product/v3-feature-inventory.md](../01-product/v3-feature-inventory.md) |
| What surfaces v4 must expose | [01-product/ux/surfaces/](../01-product/ux/surfaces/README.md) + [01-product/ux/flows/](../01-product/ux/flows/README.md) |
| Architecture decisions | [02-architecture/09-architecture-decisions/](../02-architecture/09-architecture-decisions/README.md) |
| Per-package conventions | [.github/instructions/](../../.github/instructions/) |
| Workflows | [.github/skills/](../../.github/skills/) |
| Code-level patterns | [05-engineering/](README.md) |

If a doc disagrees with this strategy, this strategy wins for v4 *forward* design; the conflicting page should be brought into line. For v3 *backward* description, see the precedence list in [docs/07-contributing/docs-contributing.md](../07-contributing/docs-contributing.md).
