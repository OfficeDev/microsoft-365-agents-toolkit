# 0007 — Inverted test pyramid for v4 lifecycle code

- **Status:** Accepted
- **Date:** 2026
- **Context tags:** v4 / testing / cli-next / core-next

## Context

The classic test pyramid (Mike Cohn, popularised by [Martin Fowler](https://martinfowler.com/articles/practical-test-pyramid.html)) prescribes lots of fast unit tests, fewer integration tests, and very few slow E2E tests. It is sound advice for most application code, where a "unit" maps to a meaningful behaviour boundary.

The Microsoft 365 Agents Toolkit is not most application code. Its lifecycle subsystem is **orchestration glue**: parsing YAML, dispatching to drivers, calling Azure / Graph / TDP APIs over HTTP, shelling out to `npm` / `dotnet` / `func`, propagating env-vars between steps, and persisting state to env files. The "unit" boundary in this code is largely artificial — a unit test of `provisionOp` with every driver mocked tests almost nothing that could realistically break in production.

Forces:

- The thing that goes wrong in production is almost always at the integration boundary: a Bicep deployment failure, a TDP 4xx, a Kudu zip-deploy timeout, an env-var resolution gap between drivers.
- v3's experience: the integration tests in `packages/fx-core/tests/integration/` (driven by [`features.json`](../../../.dev/features.json)) catch the bugs that ship; the thousands of fx-core unit tests catch refactoring breakage but rarely product issues.
- The cli-next E2E suite (`packages/cli-next/tests/e2e/`) and the `e2e-test-next.yml` daily workflow are the only signals that prove a v4 template scaffolds, provisions, deploys, and publishes against a real tenant.
- Fowler's own caveat — **"avoid test duplication"** and **"if a higher-level test catches an error and no lower-level test does, write the lower-level test"** — is still valid; we are not abandoning lower-level tests, only re-weighting them.

## Decision

For v4 lifecycle code (`core-next/src/lifecycle/`, `core-next/src/drivers/builtin/`, `core-next/src/clients/`, `cli-next/src/actions/`, `cli-next/src/commands/`), the verification weight is:

| Layer | Weight | What goes here |
|-------|--------|---------------|
| **Pure-function unit tests** | High | `secretMasker`, `featureFlags`, `localization`, `templates/scaffold/replaceMap`, `lifecycle/analyze`, `questions/treeBuilder`, `http/` retry/timeout helpers — anything with **no I/O and no `AtkContext` dependency** |
| **Integration tests** | **Highest** | `lifecycle/executor` end-to-end with real drivers + stubbed external HTTP; `templates/scaffold` with real ZIPs + real filesystem; operation pipelines (`provisionOp`, `deployOp`, `publishOp`) wired to `createMockContext()` + driver registry |
| **E2E tests** | **Highest, gating** | `cli-next` invoking real `atk` against real or sandbox M365 + Azure tenants. Required for any feature shipped in `features.json` |

For non-lifecycle code (`secretMasker`, `featureFlags`, `localization`, `localizer`, etc.) the conventional pyramid still applies — those are pure logic and unit tests are exactly the right shape.

**Operational rules:**

1. Every new template in `features.json` requires a passing cli-next E2E test before it can be marked supported in [capabilities-matrix.md](../../01-product/capabilities-matrix.md).
2. Every new `Operation` requires an integration test exercising the full pipeline (parse → resolve → execute → persist). Mocking is allowed only at the **outermost HTTP boundary** (Azure ARM, Graph, TDP, Kudu).
3. A unit test that **only re-mocks the same shape as an integration test already covers** is a code-review delete signal.
4. Pure functions get unit tests, no integration test required — the function *is* the boundary.
5. Drivers get **both**: a unit test of the Zod input validation (fast, isolated) plus participation in the integration test of the operation that uses them (slow, real).

## Consequences

- **Positive:** Test failures map to real product breakage. When CI is green, the user-facing capability works.
- **Positive:** Encourages writing integration tests against the actual driver pipeline rather than mocking it away. New driver = new integration test, automatically.
- **Positive:** E2E tests against real tenants double as smoke tests for upstream platform changes (Graph, TDP, Azure ARM API drift).
- **Negative:** Integration tests are slower than unit tests. Mitigated by: (a) integration tests still mock outbound HTTP via stubs, (b) the slow E2E suite runs daily on a schedule, not per-PR-blocking.
- **Negative:** Coverage metrics look lower than they would with many small unit tests. Acceptable: we measure by behaviour-coverage (does the user-facing capability work?), not line-coverage. The 80% NYC gate stays in effect for code where it's meaningful.
- **Neutral:** Onboarding a new contributor takes longer to reach "first green test" — they have to set up the test helper context. Mitigated by the [`createMockContext()`](../../05-engineering/testing-strategy.md#v4-test-helper) pattern.

## Alternatives considered

- **Stick with the conventional pyramid.** Rejected: the v3 experience shows the unit tests in fx-core catch refactoring bugs but rarely product bugs. Repeating the pattern in v4 wastes effort.
- **E2E only, no integration tests.** Rejected: E2E tests are too slow to run per-PR, and require real tenants. Integration tests are the per-PR signal.
- **Contract tests (Pact-style) between operations and drivers.** Rejected: drivers and operations are in the same package; the contract is enforced by Zod and TypeScript already. Pact would be ceremony without benefit.

## How this affects code review

Reviewers should ask:

- Did this PR add an integration test for any new `Operation`? If not, why?
- Did this PR add an E2E test for any new template entry in `features.json`? If not, why?
- Does this PR add a unit test that's just re-mocking what an integration test already covers? If so, delete the unit test.
- Does the PR's new pure function (e.g. a new question factory, a new env-var formatter) have a unit test? If not, add one.

## References

- Martin Fowler, [The Practical Test Pyramid](https://martinfowler.com/articles/practical-test-pyramid.html) — see §"Avoid Test Duplication" for the duplication rule we still honour
- [v4 design strategy](../../05-engineering/v4-design-strategy.md) §"E2E and integration are the primary verifications"
- [`testing-strategy.md`](../../05-engineering/testing-strategy.md)
- [`dev-test-next` skill](../../../.github/skills/dev-test-next/SKILL.md)
- [`packages/cli-next/tests/e2e/`](../../../packages/cli-next/tests/e2e/) — the v4 E2E surface
