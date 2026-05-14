# 9 — Architecture decision records (ADRs)

We use light-weight ADRs (Michael Nygard format) for significant architectural decisions. Each ADR has a stable number; we never delete or renumber.

## Status taxonomy

- **Proposed** — under discussion
- **Accepted** — in effect
- **Superseded** — replaced; the superseding ADR is linked
- **Deprecated** — no longer relevant

## Index

| # | Title | Status |
|---|-------|--------|
| [0001](0001-feature-flag-v4-core.md) | Gate the v4 engine behind `TEAMSFX_V4_CORE` | Accepted |
| [0002](0002-esbuild-over-webpack.md) | Use esbuild for vscode-extension and cli-next | Accepted |
| [0003](0003-result-pattern-neverthrow.md) | Adopt `Result<T, FxError>` via `neverthrow` everywhere | Accepted |
| [0004](0004-zod-driver-validation.md) | Pre-validate driver inputs with Zod in `createDriver()` | Accepted |
| [0005](0005-async-local-storage-correlation.md) | Use `AsyncLocalStorage` for correlation-id propagation in v4 | Accepted |
| [0006](0006-cli-next-as-reference-adapter.md) | `cli-next` is the v4 reference adapter (leading edge) | Accepted |
| [0007](0007-inverted-test-pyramid-for-lifecycle.md) | Inverted test pyramid for v4 lifecycle code | Accepted |
| [0008](0008-vibe-coding-architecture.md) | Architecture choices that enable AI-agent contributions | Accepted |

## Adding a new ADR

1. Take the next number; never reuse.
2. Filename: `NNNN-kebab-case-title.md`.
3. Use the template in [adr-template.md](adr-template.md).
4. Link from this index.
5. PR description must reference the ADR number.
