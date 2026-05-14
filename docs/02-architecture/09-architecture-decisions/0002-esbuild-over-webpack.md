# 0002 — Use esbuild for `vscode-extension` and `cli-next`

- **Status:** Accepted
- **Date:** 2025
- **Context tags:** build / v4

## Context

The legacy `cli` (v3) bundles with webpack. The configuration is ~100 lines, requires a 4 GB Node heap (`NODE_OPTIONS=--max-old-space-size=4096`), and produces multi-second cold builds. The same problems exist in the original VS Code extension build before its esbuild migration. As `cli-next` was designed, we needed a fast inner loop and a small, single-file bundle.

Forces:

- Single-file CJS bundle for fast Node startup.
- Native node addons (`keytar`, `@azure/msal-node-extensions`) cannot be bundled — must remain externals.
- Dynamic `require` in `applicationinsights` breaks naive bundlers.
- Error-class names appear in telemetry — must not be mangled.
- Source maps required for production debugging.

## Decision

Use **esbuild** with a small `esbuild.mjs` script for both `vscode-extension` and `cli-next`:

- Entry → single `build/index.js`.
- `platform: "node"`, `target: "node18"`, `format: "cjs"`.
- `keepNames: true`.
- Externals: `keytar`, `@azure/msal-node-extensions`, `applicationinsights`.
- Sourcemaps + metafile (`build/meta.json`) for analysis.
- Lazy-load heavy deps (`applicationinsights`, `node-machine-id`, driver registration) via runtime `require()`.

Keep `tsc` for type checking and declaration emit; esbuild handles transpilation and bundling.

`core-next` does *not* bundle (pure `tsc`); it is bundled inline by `cli-next`'s esbuild.

## Consequences

- **Positive:** Sub-second incremental builds. Tens of seconds for a clean `package` (production) build.
- **Positive:** No special heap settings needed.
- **Positive:** Smaller bundle than webpack equivalent.
- **Negative:** `eslint-plugin-import` cannot resolve workspace TypeScript paths — disabled `import-x/no-unresolved` in v4 packages; rely on `tsc` for real import errors.
- **Neutral:** Configuration lives in plain JS, not JSON — discoverable.

## Alternatives considered

- **Stay on webpack.** Rejected: too slow, too much config, worse DX.
- **Rollup.** Rejected: better suited to libraries (`spec-parser`, `sdk`), worse for application bundling with externals.
- **Keep tsc-only and accept multi-file output.** Rejected: pnpm symlink flattening and dynamic `require` calls would break production deploys.

## References

- [`packages/cli-next/esbuild.mjs`](../../../packages/cli-next/esbuild.mjs)
- [`packages/vscode-extension/esbuild.mjs`](../../../packages/vscode-extension/esbuild.mjs)
- [`cli.instructions.md`](../../../.github/instructions/cli.instructions.md) §"Bundling (esbuild)"
