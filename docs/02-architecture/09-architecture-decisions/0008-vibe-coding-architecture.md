# 0008 — Architecture choices that enable AI-agent contributions

- **Status:** Accepted
- **Date:** 2026
- **Context tags:** v4 / vibe-coding / architecture / contribution-velocity

## Context

The toolkit is increasingly contributed to by AI coding agents (GitHub Copilot Coding Agent, Claude Code, Cursor, internal experiments). Two pressures are reshaping how we design v4:

- **Agents need a different surface than humans.** A human can hold a god-class like `FxCore` (57 public methods) in working memory across multiple files. An agent reads one file at a time and reasons over what it can see; the more state lives in unreferenced files, the more the agent has to guess.
- **The PRD ships in production.** "Vibe coding" in the throwaway sense (per [Karpathy's original definition](https://karpathy.bearblog.dev/vibe-coding-menugen/) — accept-all, never read the diff) is incompatible with shipping a product to millions of installs. Agent contributions need the same review, testing, and design-first gates as human contributions.

We need to decide what architecture choices specifically *enable* agent contribution while keeping the safety bar.

## Decision

Adopt the following architecture pillars in v4 explicitly to make the codebase agent-readable and agent-extensible, while pairing each with a matching safety gate.

### Enabler 1 — Registries everywhere (no implicit ordering)

`TemplateRegistry`, `DriverRegistry`, `FeatureFlagRegistry` hold `*Descriptor` records. Adding a feature is adding a record.

- **Why it helps an agent:** the agent can see the entire feature surface by reading one file (e.g. [`packages/core-next/src/templates/descriptors/index.ts`](../../../packages/core-next/src/templates/descriptors/index.ts)). No "first-imported wins" magic to discover.
- **Safety gate:** `TemplateDescriptor` and `DriverDescriptor` include enough metadata that registration collisions are rejected at startup.

### Enabler 2 — Descriptors over class hierarchies

A driver is a record produced by `createDriver({ id, name, inputSchema, execute })`. A template is a record produced by `templateRegistry.register({ id, ..., questions, scaffoldFn? })`. An operation is a record produced by `defineOperation(name, schema, fn)`.

- **Why it helps an agent:** the agent doesn't have to chase inheritance chains, decorators, or DI containers. A new driver is a new file + a registry call.
- **Safety gate:** every descriptor has a typed schema (Zod for drivers; TypeScript types for templates). Missing fields are compile-time errors.

### Enabler 3 — Zod input schemas on drivers

Every `createDriver()` call declares an `inputSchema: ZodSchema`. The factory validates `with:` from YAML before calling `execute`.

- **Why it helps an agent:** the agent gets a fully-typed `input` parameter inside `execute`. No guessing the YAML shape from comments or example calls.
- **Safety gate:** validation happens *before* `execute` runs. Typos in YAML surface as `InvalidDriverInput` with the issue path, not as `Cannot read property of undefined` 200 lines deep.

### Enabler 4 — `Result<T, AtkError>` everywhere

No throwing for expected failures. Every fallible function returns `Result<T, AtkError>` from `neverthrow`.

- **Why it helps an agent:** the agent can reason about the type signature. `Promise<Result<T, E>>` tells it both branches must be handled. There's no hidden "may throw" channel.
- **Safety gate:** the type system rejects unhandled `Result`s. ESLint rejects floating promises.

### Enabler 5 — `AtkContext` injection; no module-scoped state

Every operation, driver, client, and helper takes `ctx: AtkContext` (with `logger`, `telemetry`, `ui`, `auth`, `correlationId`, optional `projectPath`). No `TOOLS` global.

- **Why it helps an agent:** the agent can write a test for any function immediately by calling [`createMockContext()`](../../05-engineering/testing-strategy.md#v4-test-helper). No setup ritual; no "where do I get a logger from" archaeology.
- **Safety gate:** `createMockContext()` returns a fully-stubbed context with sinon spies. Tests catch any code path that tries to bypass the injection.

### Enabler 6 — Single-file-readable modules

Per-feature files stay under ~300 lines where possible. Module `index.ts` exports the public surface. When a file passes ~500 lines, that's a refactor signal.

- **Why it helps an agent:** the agent can fit a module in one file read. No context-window spillage from 1500-line god classes.
- **Safety gate:** no automated gate for this — it's a code-review heuristic. Reviewers flag files that grow past the threshold without justification.

### Enabler 7 — Source-of-truth precedence is documented

[docs/07-contributing/docs-contributing.md §"Source-of-truth precedence"](../../07-contributing/docs-contributing.md) explicitly states: code → instructions → skills → `features.json` → docs prose, for backward description; and design pages → code, for forward design.

- **Why it helps an agent:** when the agent finds two sources that disagree, it knows which to trust. No hallucinated "this is the answer" that contradicts source.
- **Safety gate:** the precedence list is itself in the docs site. Validation passes (we just ran one) catch drift between sources.

### Enabler 8 — Design-first workflow; docs in same PR

Non-trivial v4 changes land docs + code in the same PR. The design page is the spec the implementation realises.

- **Why it helps an agent:** the agent's first task is to write or update the design page. That doubles as the prompt for the implementation step. The agent gets to read its own spec back when it implements.
- **Safety gate:** [`codebase.instructions.md` §"Source-of-Truth Workflow"](../../../.github/instructions/codebase.instructions.md) requires this. Reviewers reject PRs that change v4 behaviour without changing v4 docs.

### Enabler 9 — `_v3-reference/` quarantine

v3 internal shapes (FxCore class, lifecycle YAML internal model, error catalogue, generator categories) are quarantined under [`docs/_v3-reference/`](../../_v3-reference/README.md) and **forbidden as v4 design input**.

- **Why it helps an agent:** the agent designing v4 can't accidentally import v3's accidental complexity (god class, implicit ordering, untyped YAML dispatch).
- **Safety gate:** the quarantine README explicitly lists what cannot be carried over. PRs that mirror v3's shape get review-rejected with a pointer to that file.

### Vibe-coding workflow gates

The [`vibe-coding`](../../../.github/skills/vibe-coding/SKILL.md) skill codifies the agent workflow. The non-negotiable gates are:

| Gate | Where enforced |
|------|----------------|
| Design page written or updated in same PR | `codebase.instructions.md` + reviewer |
| Driver inputs Zod-validated | `createDriver()` factory |
| Operation has integration test | `ci-next.yml` integration job |
| New feature in `features.json` has E2E test | `e2e-test-next.yml` |
| Lint + format + 80% coverage | `ci-next.yml` lint, format-check, unit-test jobs |
| `Result<T, AtkError>` returned, not thrown | TypeScript + reviewer |
| Conventional Commits | commitlint hook |
| Stable error names | reviewer (telemetry partition key) |

A PR that bypasses any of these is rejected — by the system or by review. The architecture *enables* fast agent contribution; the gates *ensure* the contribution is safe to merge.

## Consequences

- **Positive:** Agent contributions land faster because the surface is enumerable, typed, and self-describing.
- **Positive:** Human contributors get the same benefits — the architecture choices that help an agent also help a new human.
- **Positive:** Drift between v3 and v4 internal patterns is bounded by the explicit quarantine, not by hope.
- **Negative:** Up-front cost — Zod schemas, registry registration, `AtkContext` plumbing — every new feature has more boilerplate than v3's "just throw it in a file" pattern. Acceptable: the cost is paid once per feature; the readability gain compounds.
- **Negative:** Risk of "design page theatre" — writing a thin design page just to satisfy the gate. Mitigation: reviewers check design pages match the implementation, not just exist.
- **Neutral:** This ADR doesn't mandate any specific AI-coding tool. The architecture works for any agent (Copilot, Claude, Cursor, future tools) and any human.

## Alternatives considered

- **No vibe-coding-specific architecture.** Rejected: the architecture choices above (registries, descriptors, Zod, DI, Result) are good engineering regardless. Naming them as enablers makes the cultural choice explicit; the choices stand on their own merits.
- **Sandbox / experimental directory for vibe coding.** Rejected per the [user direction](../../../.github/skills/vibe-coding/SKILL.md): all PRs are production-bound. Sandbox would create a parallel codebase that drifts.
- **Adopt a copilot-friendly DSL.** Rejected: domain-specific languages add learning cost for humans and don't actually help agents (they reason fine over standard TypeScript). Stick to typed records.

## References

- Karpathy, [Vibe coding MenuGen](https://karpathy.bearblog.dev/vibe-coding-menugen/) — concrete pain points (services not designed for LLMs, env-var leakage, gluing a "monster")
- Simon Willison, [Not all AI-assisted programming is vibe coding](https://simonwillison.net/2025/Mar/19/vibe-coding/) — the production-vs-throwaway distinction
- [Command Line Interface Guidelines](https://clig.dev/) — CLI design principles
- [v4 design strategy §"Vibe coding for the toolkit"](../../05-engineering/v4-design-strategy.md)
- [`vibe-coding` skill](../../../.github/skills/vibe-coding/SKILL.md)
- [ADR 0006 — cli-next as reference adapter](0006-cli-next-as-reference-adapter.md)
- [ADR 0007 — Inverted test pyramid](0007-inverted-test-pyramid-for-lifecycle.md)
