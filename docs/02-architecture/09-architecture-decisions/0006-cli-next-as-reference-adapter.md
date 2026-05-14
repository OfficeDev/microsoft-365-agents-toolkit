# 0006 — `cli-next` is the v4 reference adapter

- **Status:** Accepted
- **Date:** 2026
- **Context tags:** v4 / cli-next / surfaces / rollout

## Context

v4 (`core-next` engine) needs a consumer adapter to be exercised end-to-end. The two candidate adapters are the VS Code extension and the CLI. We have to pick which one leads.

Forces:

- **Surface cost.** A new operation in CLI is a flag declaration + an action function + telemetry wrapper (~30 lines). The same operation in VS Code is tree-view registration, command-palette entry, package.json contribution, telemetry attribution, error-display wiring, often a webview — easily 200+ lines and several files.
- **CI is CLI.** [`e2e-test-next.yml`](../../../.github/workflows/e2e-test-next.yml) and the cli-next E2E suite both invoke `atk` directly. The CLI is the path CI exercises; making it primary aligns the verification target with the implementation target.
- **AI-agent adopters.** GitHub Copilot Coding Agent, Claude Code, Cursor, and other coding agents all drive CLIs over shells. A clean cli-next surface is what makes v4 features reachable from those agents — the VS Code extension is, by design, not addressable from a sibling agent.
- **Registry-driven CLI.** v4's `TemplateRegistry` already auto-generates `atk new <category>` subcommands via [`buildNewCommands()`](../../../packages/cli-next/src/commands/factory.ts). Adding a template adds a CLI subcommand for free. The same auto-generation is harder in VS Code because each command needs `package.json` registration.
- **VS Code is sticky.** Forcing the VS Code extension to migrate to v4 immediately would either (a) ship two engines side by side in one extension binary or (b) require a coordinated cutover. Both are riskier than the CLI's clean opt-in via `TEAMSFX_V4_CORE`.

## Decision

`cli-next` is the **leading edge** for v4. Every new v4 capability lands in this order:

1. New operation in `core-next` (with `defineOperation`, Zod input schema, integration test).
2. New action in `cli-next` (with `wrapHandlerWithContext`, telemetry name, E2E test).
3. New design-page entry under `docs/01-product/ux/surfaces/cli-v3-command-reference.md` (or a new v4 reference if it's a v4-only command).
4. (Later, in a separate PR) VS Code extension consumes the same operation through the eventual `createVsCodeContext()` shim, gated by `TEAMSFX_V4_CORE`.

Steps 1–3 land in **one PR**. Step 4 happens after the v4 engine is proven against the cli-next E2E suite for the feature in question.

The corollary: **the VS Code extension is not blocked on every cli-next change.** It continues to ship on the v3 engine for as long as `TEAMSFX_V4_CORE=false` is the default. When the flag flips, the extension catches up to the v4 surface that cli-next has already proven.

## Consequences

- **Positive:** v4 ships features faster. The cost of a new operation is one CLI surface, not three.
- **Positive:** CI exercises every new feature end-to-end immediately, because cli-next is what CI runs.
- **Positive:** AI coding agents can drive new v4 features the moment they ship — no extension marketplace lag.
- **Positive:** The VS Code extension stays on stable v3 ground. No coordinated cutover risk.
- **Negative:** VS Code-only users see new v4 capabilities later than CLI users. Mitigation: most CLI features map cleanly to existing VS Code commands; the extension migration adds them to the tree view at flip time.
- **Negative:** Two adapter implementations of the same operation (CLI today, VS Code later) means temporary duplication. Acceptable: the duplication is at the adapter layer, not the engine layer; both consume the same `Operation`.

## Alternatives considered

- **VS Code extension first.** Rejected: 6–10× the surface cost per feature; CI doesn't drive the VS Code extension end-to-end (only its tests); coordinated cutover risk.
- **Both adapters in lock-step.** Rejected: forces every PR to touch both packages; breaks the design-first-in-one-PR principle; doubles per-feature cost.
- **A new cross-surface adapter layer.** Rejected: premature abstraction. We don't yet know what the right shared shape is. Repeat ourselves a few times across CLI and VS Code first; abstract when the duplication actually costs.

## When this ADR ends

This ADR is in effect until the VS Code extension migrates onto `core-next` (i.e. `TEAMSFX_V4_CORE` is the default for the extension). After that, this ADR is superseded by a new one describing co-equal adapter status.

## References

- [v4 design strategy](../../05-engineering/v4-design-strategy.md) §"cli-next is the leading edge"
- [`packages/cli-next/src/commands/factory.ts`](../../../packages/cli-next/src/commands/factory.ts) — `buildNewCommands()` auto-generation
- [ADR 0001 — TEAMSFX_V4_CORE feature flag](0001-feature-flag-v4-core.md)
- [`cli.instructions.md`](../../../.github/instructions/cli.instructions.md) §"CLI v4"
