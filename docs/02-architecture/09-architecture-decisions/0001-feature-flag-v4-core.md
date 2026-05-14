# 0001 — Gate the v4 engine behind `TEAMSFX_V4_CORE`

- **Status:** Accepted
- **Date:** 2025
- **Context tags:** v4 / cross-cutting / rollout

## Context

`fx-core` (v3) has accumulated significant architectural debt: a `TOOLS` global singleton, implicit generator activation order, ad-hoc YAML driver dispatch, mixed responsibilities between `api`, `manifest`, and `fx-core`. A clean rewrite (`core-next`) is desirable, but we cannot afford to break the tens of thousands of v3 projects in the wild.

Forces:

- Need to iterate on the v4 design without disrupting shipping product.
- Need real telemetry from preview users on v4 paths.
- Need v3 tests, CI, and shipping pipeline to continue unaffected.
- Need a clean cutover path when v4 reaches parity.

## Decision

Ship `core-next` and `cli-next` alongside v3 packages. Introduce a single feature flag `TEAMSFX_V4_CORE` (read by both engines) that:

- Defaults to **off**.
- When **on**, makes consumer adapters route through `core-next` operations instead of `fx-core` methods.
- Is checked once per consumer entry point — not threaded through internal calls.

Preserve the npm package name `@microsoft/teamsfx-core` for both — `core-next` publishes as v4.0.0 while `fx-core` continues at the v3 line. CI uses `pnpm --filter ./packages/core-next` (directory path) to disambiguate.

## Consequences

- **Positive:** Both engines can ship in the same release. Rollback is one env var. Internal teams and OSS contributors can opt in.
- **Positive:** New v4 features land without coordination tax on v3.
- **Negative:** Two engine paths to maintain in CI and docs. Mitigated by `ci-next.yml` running independently.
- **Negative:** Some duplication (DI shapes, error types). Acceptable through migration window.

## Alternatives considered

- **Branch-based development.** Rejected: long-lived branches drift; integration becomes painful.
- **Direct cutover at a major version bump.** Rejected: risk of regressions affecting all users simultaneously.
- **Per-feature flags inside `fx-core`.** Rejected: `core-next` rewrites cross-cutting concepts (DI, registries, telemetry) — gating each individually is impractical.

## References

- [`packages/core-next`](../../../packages/core-next/)
- [`packages/cli-next`](../../../packages/cli-next/)
- [`fx-core.instructions.md`](../../../.github/instructions/fx-core.instructions.md) §"core-next (v4 Successor)"
