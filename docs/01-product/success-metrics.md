# Success metrics

Indicators we use to judge whether the product is delivering on its [vision](../00-overview/product-vision.md). Each metric pairs a **what we measure** with a **where it surfaces** so engineers can wire up new instrumentation correctly.

## Activation

| Metric | Definition | Surface |
|--------|-----------|---------|
| Time-to-first-scaffold | Minutes from extension install to a non-empty project | VS Code telemetry: `extension.activate` → first `create-project-end` |
| Time-to-first-run | Minutes from scaffold to first successful F5 / `atk preview` | `local-debug-end` success per project |
| Scaffold completion rate | `create-project-end` / `create-project-start` | Both engines |

## Lifecycle reliability

| Metric | Definition | Surface |
|--------|-----------|---------|
| Provision success rate | `provision-end` (success) / `provision-start` per template | v3: `fx-core` telemetry · v4: `instrumentOperation(provisionOp)` |
| Deploy success rate | Same shape, `deploy-*` events | v3 + v4 |
| Publish success rate | Same shape, `publish-*` events | v3 + v4 |
| Mean retries to green | Average `provision-end (failure)` count per project before a `provision-end (success)` | Aggregated |

## Coverage

| Metric | Definition | Source |
|--------|-----------|--------|
| Templates shipped | Count of entries in `.dev/features.json` (testable + tracked) | [`features.json`](../../.dev/features.json) |
| v4 E2E coverage | Templates with verified scaffold/provision/deploy/publish on `core-next` | [`features.instructions.md`](../../.github/instructions/features.instructions.md) §"Current features" |
| Driver coverage | Built-in driver count vs. distinct YAML actions referenced by templates | `packages/core-next/src/drivers/builtin/index.ts` |

## Engineering health

| Metric | Definition | Source |
|--------|-----------|--------|
| Test count (v3) | Mocha suite count in `packages/fx-core` | `npm run test:unit` |
| Test count (v4) | core-next: 606 unit + 48 integration · cli-next: 87 unit + 81 integration | `ci-next.yml` artifacts |
| Coverage gate | NYC threshold 80% in core-next | `.nycrc` |
| Lint gate | 0 errors (warnings allowed: `no-explicit-any`) | `eslint.config.mjs` per package |
| Bundle size (cli-next) | `build/index.js` size after esbuild `--production` | `build/meta.json` |

## Adoption (external)

External adoption signals (npm downloads, marketplace installs, GitHub stars) are tracked outside this repo. Pointers:

- VS Code marketplace: <https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension>
- npm: `@microsoft/m365agentstoolkit-cli`, `@microsoft/teamsfx-core`, `@microsoft/teamsfx-api`, `@microsoft/teamsfx-core-next` (preview)

## Telemetry contracts

All counters above must satisfy the [security baseline](../05-engineering/security.md): no PII, secret masking via [secretMasker](../05-engineering/cross-cutting/service-clients.md#secret-masker), correlation IDs propagated from entry to exit. See [05-engineering/telemetry.md](../05-engineering/telemetry.md).
