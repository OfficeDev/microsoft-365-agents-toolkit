# Build and bundling

## Bundlers in use

| Package | Bundler | Config | Notes |
|---------|---------|--------|-------|
| `vscode-extension` | **esbuild** | `esbuild.mjs` | Single-file CJS, native MSAL handling |
| `cli-next` | **esbuild** | `esbuild.mjs` | Single-file CJS, externals for native addons |
| `cli` (v3) | webpack | `webpack.config.js` | Legacy — needs `NODE_OPTIONS=--max-old-space-size=4096` |
| `fx-core` | webpack | `webpack.config.js` | Legacy |
| `vscode-ui` | webpack | — | Legacy |
| `mcp-server` | webpack | — | Legacy |
| `spec-parser` | rollup | `rollup.config.js` | Library shape |
| `core-next` | none | — | Pure `tsc`; bundled inline by cli-next |
| `sdk`, `sdk-react`, `adaptivecards-tools-sdk` | rollup | — | Library shape |

## esbuild specifics (`vscode-extension`, `cli-next`)

- Entry → single `build/index.js`.
- `platform: "node"`, `target: "node18"`, `format: "cjs"`.
- `keepNames: true` — error class names appear in telemetry; never mangle.
- Source maps + metafile (`build/meta.json`).
- Externals: `keytar`, `@azure/msal-node-extensions` (native `.node` addons), `applicationinsights` (dynamic `require`).

### Lazy-loading patterns

To minimise CLI cold start (`atk --help` < 200 ms):

- `applicationinsights` — `require()` inside `AppInsightsTransport.init()`, not at module top.
- `node-machine-id` — `require()` inside `CliTelemetryReporter.init()`.
- `registerBuiltinDrivers()` — deferred from `start()` to `wrapHandlerWithContext()` (only when a real command runs, not for `--help`).

## Build commands

### Monorepo

```bash
npm run setup              # full install + build
npm run setup:cli          # CLI v3 + dependencies
npm run setup:vsc          # VS Code extension + dependencies
npm run watch              # watch mode (all packages)
pnpm --filter @microsoft/teamsfx-core run build   # single v3 package
pnpm --filter ./packages/core-next run build      # single v4 package (DIR path required because npm name collides with v3)
```

### v4 packages

```bash
cd packages/core-next
npm run build            # tsc + postbuild (eslint --fix + prettier --write)
npm run test:unit        # 606 unit tests with NYC
npm run test:integration # 48 integration tests

cd packages/cli-next
npm run build            # dev: tsc only — type checking + declarations
npm run bundle           # dev: esbuild bundle, no minification
npm run package          # prod: tsc + esbuild --production (minified, source maps)
npm run prepack          # alias for npm run package — runs before publish
npm run test:unit        # 87 unit tests
npm run test:integration # 81 integration tests
```

### `postbuild` hook

Both `core-next` and `cli-next` run `eslint --fix` + `prettier --write` automatically after `tsc` build. No separate format step needed.

## Bundle size

`cli-next` ships as a single CJS file. Inspect with:

```bash
node packages/cli-next/build/index.js --help          # cold-start sanity
node -e "require('fs').statSync('packages/cli-next/build/index.js')" # size
```

Bundle analysis: open `packages/cli-next/build/meta.json` in <https://esbuild.github.io/analyze/>.
