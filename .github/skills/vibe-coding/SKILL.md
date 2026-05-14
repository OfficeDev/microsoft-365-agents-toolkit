---
name: vibe-coding
description: "End-to-end workflow for AI-agent-driven contributions to v4 packages (core-next, cli-next). Use when: implementing a new v4 operation, driver, or template; fixing a bug in core-next/cli-next; adding a CLI command. Codifies the design-first → Zod-validated → integration-tested → docs-in-same-PR cycle that keeps agent contributions safe to merge."
argument-hint: "Describe the v4 feature, fix, or capability"
---

# Vibe Coding (v4)

## When to use this skill

Use this skill for **any non-trivial change to `core-next` or `cli-next`**, whether a human or an AI agent (Copilot, Claude Code, Cursor, etc.) is doing the work. The workflow is the same; "vibe coding" in this codebase means **agent-driven implementation with hard safety gates**, never throwaway sandbox work.

Don't use this skill for:

- v3 (`fx-core`, `cli`, `vscode-extension`, `server`) — see [`dev-workflow`](../dev-workflow/SKILL.md). v3 is code-first.
- Trivial v4 changes (typo fixes, dependency bumps, doc-only edits) — go straight to PR.
- Reviewing or understanding existing v3 internals — read [`docs/_v3-reference/`](../../docs/_v3-reference/README.md). That folder is **forbidden as v4 design input**.

## What this skill enforces

The architecture is laid out in [ADR 0008](../../docs/02-architecture/09-architecture-decisions/0008-vibe-coding-architecture.md). The workflow below realises it.

The non-negotiable gates per PR:

1. Design page written or updated **in the same PR** as code.
2. Every new driver has a Zod `inputSchema`.
3. Every new `Operation` has an integration test.
4. Every new feature added to `.dev/features.json` has a passing cli-next E2E test.
5. Lint clean, format clean, 80% coverage gate green.
6. `Result<T, AtkError>` returned from every fallible function — no thrown exceptions for expected failures.
7. Stable `error.name` strings (telemetry partition keys; renaming is a deprecation event).
8. Conventional Commits (`feat(core-next):`, `fix(cli-next):`).

A PR that bypasses any of these is rejected — by CI or by review.

## Workflow

### Phase -1 — Capture (chat trigger only)

Skip this phase if the work originates from a GitHub Issue or ADO Work Item.

When triggered from a chat description (e.g. "I want X to support Y"):

1. **Classify**: `feature` / `bug` / `chore`
2. **Identify domain**: which of the 7 capability domains does this touch? (`docs/04-specs/domains/`)
3. **Create issue**: GitHub Issue or ADO Work Item; use the chat description as body
4. **Continue to Phase 0** with the newly created issue as input

### Phase 0 — Design first

For any non-trivial change, locate or write the spec **before** writing code.
Specs are the authoritative behavioral contracts that tests are derived from.

**Start here:** [`docs/04-specs/README.md`](../../docs/04-specs/README.md) — defines the full layer hierarchy (PRD → User Story → Domain Spec → Operation Spec → Tests → Code) and spec formats.

#### Spec location

| Change kind | Where the spec lives |
|-------------|---------------------|
| New / changed operation, driver, or lifecycle stage | [`docs/04-specs/operations/<domain>/<operation>.md`](../../docs/04-specs/operations/) |
| Domain boundary change or new cross-domain interface | [`docs/04-specs/domains/<nn>-<domain>.md`](../../docs/04-specs/domains/) |
| Architectural decision | New ADR under [`docs/02-architecture/09-architecture-decisions/`](../../docs/02-architecture/09-architecture-decisions/README.md) using the [template](../../docs/02-architecture/09-architecture-decisions/adr-template.md) |
| New / changed data contract or entity | [`docs/04-specs/data-model/entities/`](../../docs/04-specs/data-model/entities/README.md) |
| New CLI surface, command group, flow | [`docs/01-product/ux/surfaces/cli.md`](../../docs/01-product/ux/surfaces/cli.md), [`docs/01-product/ux/flows/`](../../docs/01-product/ux/flows/README.md) |
| New capability / template | [`docs/01-product/capabilities-matrix.md`](../../docs/01-product/capabilities-matrix.md) + [`.dev/features.json`](../../.dev/features.json) |

#### What a spec must contain before code is written

For an operation spec, all of these sections must be complete:
- `## Acceptance Criteria` — ID-based table, each row maps to one test case (no rows = blocked)
- `## Boundary` — explicit list of what this operation does NOT do
- `## Invariants` — constraints that must never be violated

**AI gap discovery:** Attempt to complete the spec. Where sections cannot be filled,
output specific questions rather than guessing:
> "Cannot complete AC table because:
>  1. Should an unreachable MCP server fail immediately or retry? How many times?
>  2. Is adding the same action URL twice an error or idempotent?"

Blocked spec = upstream ambiguity. Do not proceed to implementation. Surface to PM → update PRD first.

**Two human gates before any code is written:**
1. **Gate 1 — Answer AI questions**: resolve all ambiguities surfaced during spec drafting
2. **Gate 2 — Approve AC table**: review each row; scenario clear? expected result correct?

Only after Gate 2 approval does AI generate tests and implement.

For an agent: treat `## Acceptance Criteria` rows as the test plan,
`## Boundary` + `## Invariants` as hard constraints on implementation scope.

### Phase 1 — Confirm allowed inputs

Before designing, verify the design has the right inputs:

| Allowed input to v4 design | Forbidden as v4 design input |
|----------------------------|------------------------------|
| [`docs/01-product/`](../../docs/01-product/README.md) (PRD, features, scenarios) | The v3 `FxCore` class signature |
| [`docs/01-product/ux/`](../../docs/01-product/ux/README.md) (surfaces, flows, question model) | The v3 lifecycle YAML internal model (`RawProjectModel`, `ILifecycle`) |
| Microsoft 365 platform contracts (manifest schemas, Graph, TDP) | The v3 `StepDriver` interface |
| Existing v4 ADRs | The v3 generator activation pattern |
| Existing v4 design pages in `docs/04-specs/data-model/`, `docs/05-engineering/` | The full v3 error catalogue's organising structure |
| Microsoft Foundry / Agents SDK / Teams AI library upstream contracts | Existing v3 Bicep templates (re-design from topology requirements) |

If the design copies a shape from `_v3-reference/`, that's a code-review red flag. Ask: *"What product requirement justifies this shape, independent of how v3 happens to do it?"*

See [`docs/_v3-reference/README.md`](../../docs/_v3-reference/README.md) §"The hard rule" for the full quarantine.

### Phase 2 — Implement against the design

Follow the v4 architectural pillars. The pillars are **not negotiable per-feature** — when a new design contradicts a pillar, the design changes, not the pillar.

| Pillar | What it means in code |
|--------|----------------------|
| Registries everywhere | New template = `templateRegistry.register({ ... })` in `descriptors/<category>.ts`. New driver = `createDriver({ ... })` + add to `builtin/index.ts`. |
| Descriptors over class hierarchies | No new class hierarchies. New work = factory function returning a record. |
| Errors as values | Return `Result<T, AtkError>`. `import { ok, err } from "neverthrow"`. |
| Driver inputs Zod-validated | `inputSchema: z.object({ ... })` is required in `createDriver()`. |
| `AtkContext` injection | Operations take `(input, ctx)`. Drivers take `(input, ctx)`. Clients are constructed with `(ctx)`. No globals. |
| Single-file-readable | Keep new files under ~300 lines. Refactor at ~500. |
| Operations compose | New lifecycle stage = new driver + new operation composing prerequisites + executor. |

For per-area conventions:

- Operations / drivers / templates / question model: [`fx-core.instructions.md`](../../.github/instructions/fx-core.instructions.md) §"core-next (v4 Successor)"
- CLI commands / actions / auth: [`cli.instructions.md`](../../.github/instructions/cli.instructions.md) §"CLI v4"
- Cross-cutting (formatting, lint, async, security): [`codebase.instructions.md`](../../.github/instructions/codebase.instructions.md)

### Phase 3 — Test (inverted pyramid)

For lifecycle / orchestration code, integration and E2E tests are the **primary** verification. See [ADR 0007](../../docs/02-architecture/09-architecture-decisions/0007-inverted-test-pyramid-for-lifecycle.md).

Tests are derived from the approved AC table in the operation spec.
**Test names must include the AC ID** so reviewers can trace test ↔ spec:

```typescript
// AC-01: valid inputs → success
it("AC-01: returns clientId and objectId on success", ...)

// AC-03: name > 120 chars → UserError
it("AC-03: returns UserError(AadAppNameTooLong) when name exceeds limit", ...)
```

| What you wrote | What test it needs |
|----------------|-------------------|
| Pure function (no I/O, no `AtkContext`) | Unit test (`tests/unit/<area>/`) |
| New `Operation` | Integration test exercising full pipeline (`tests/integration/`) — mock only outermost HTTP |
| New driver with Zod schema | Unit test of schema validation + driver's participation in integration test |
| New template added to `.dev/features.json` | cli-next E2E test in `packages/cli-next/tests/e2e/` |
| New CLI action | cli-next integration test + E2E if user-facing |

A unit test that only re-mocks what an integration test already covers is a delete signal.

For local test setup: [`dev-test-next` skill](../dev-test-next/SKILL.md).

### Phase 4 — Verify the gates

Before opening the PR, run:

```bash
# In packages/core-next or packages/cli-next:
npm run build              # tsc + postbuild (eslint --fix + prettier --write)
npm run test:unit          # unit tests with NYC coverage
npm run test:integration   # integration tests
npm run lint               # 0 errors required
npm run format:check       # CI gate
```

For a v4 feature that adds to `features.json`, also run the relevant cli-next E2E:

```bash
cd packages/cli-next
npm run test:e2e -- --grep "<your-template-id>"
```

See the [`dev-test-next` skill](../dev-test-next/SKILL.md) for full local E2E setup (M365 + Azure credentials, etc.).

### Phase 5 — Update docs in the same PR

Did the design page from Phase 0 match what you actually built? If not, update it now. The PR must contain both the code change and the doc change. PRs that ship code without doc updates are rejected at review.

Specifically:

- If you added a new template: update [`docs/01-product/capabilities-matrix.md`](../../docs/01-product/capabilities-matrix.md) and [`docs/01-product/v3-feature-inventory.md`](../../docs/01-product/v3-feature-inventory.md), and add a `.dev/features.json` entry.
- If you added a new operation: update [`docs/05-engineering/cross-cutting/lifecycle-engine.md`](../../docs/05-engineering/cross-cutting/lifecycle-engine.md) (or wherever the operation lives).
- If you added a new driver: update [`docs/05-engineering/cross-cutting/driver-system.md`](../../docs/05-engineering/cross-cutting/driver-system.md) §"Built-in driver catalogue".
- If you added a new CLI command: update [`docs/01-product/ux/surfaces/cli-v3-command-reference.md`](../../docs/01-product/ux/surfaces/cli-v3-command-reference.md) — or split off a v4-specific reference if it's a v4-only command.
- If the change has architectural impact: file an ADR under [`docs/02-architecture/09-architecture-decisions/`](../../docs/02-architecture/09-architecture-decisions/README.md).

### Phase 6 — PR

Conventional Commits format (commitlint enforces):

```
feat(core-next): add <thing> operation
fix(cli-next): handle <case> in provision action
docs(architecture): add ADR-NNNN for <decision>
```

PR description must reference the design page or ADR.

CODEOWNERS will auto-assign reviewers. Reviewers will check:

- ✅ Design page or ADR present in this PR
- ✅ Driver inputs Zod-validated (if drivers were touched)
- ✅ Integration test for any new `Operation`
- ✅ E2E test for any new `features.json` entry
- ✅ `Result<T, AtkError>` returned, not thrown
- ✅ Stable `error.name` strings
- ✅ Lint / format / coverage gates green

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| "I copied this shape from `fx-core/...`" | Used v3 internals as design input | Stop. Re-read the design page from Phase 0. Re-derive the shape from the PRD requirement. |
| Driver throws an exception instead of returning `Result` | Forgot the `Result` pattern | Wrap in try/catch; return `err(new AtkError({...}))`. |
| Driver `with:` typo crashes deep in `execute()` | Skipped Zod schema | Add `inputSchema: z.object({...})` to `createDriver()`. |
| Test passes locally, fails in CI | Probably a `process.env` race or async cleanup gap | Run `tests/integration/` locally; check for missing `await`. |
| Reviewer says "no docs change" | Skipped Phase 5 | Add the doc update to this PR; don't open a follow-up. |
| Reviewer says "this looks like v3" | Used v3 shapes as input | Same as the first row. |
| File is 800 lines | Single-file-readable pillar violated | Refactor — split per-feature concerns. Reviewer will request this. |

## Anti-patterns to flag in review

- A new class hierarchy in `core-next` — should be factory functions returning records.
- A new module-scoped variable holding runtime state — should be on `AtkContext`.
- A new driver without `inputSchema` — required.
- A new `Operation` without an integration test — required.
- A new template in `features.json` without an E2E test — required.
- A "TODO: add tests later" comment — not allowed; this is the test-PR.
- A "TODO: add docs later" comment — not allowed; this is the doc-PR.
- Anything that mirrors the v3 `FxCore` god-class shape — re-derive from PRD.
- Use of `console.log` — use `ctx.logger`.
- Throwing for an expected failure — return `err(...)`.
- Catching `unknown` and silently swallowing — preserve via `innerError`.
- A vague `error.name` like `Error` or `Failed` — name it for the failure mode.

## See also

- [v4 design strategy](../../docs/05-engineering/v4-design-strategy.md) — the why behind this workflow
- [ADR 0006 — cli-next as reference adapter](../../docs/02-architecture/09-architecture-decisions/0006-cli-next-as-reference-adapter.md)
- [ADR 0007 — Inverted test pyramid](../../docs/02-architecture/09-architecture-decisions/0007-inverted-test-pyramid-for-lifecycle.md)
- [ADR 0008 — Architecture for AI-agent contributions](../../docs/02-architecture/09-architecture-decisions/0008-vibe-coding-architecture.md)
- [`dev-test-next` skill](../dev-test-next/SKILL.md) — local test setup
- [`dev-workflow` skill](../dev-workflow/SKILL.md) — for v3 work (code-first, not design-first)
- [`docs-guard` skill](../docs-guard/SKILL.md) — keeps docs in sync after each turn
- [`plan-tracker` skill](../plan-tracker/SKILL.md) — for multi-step v4 work
