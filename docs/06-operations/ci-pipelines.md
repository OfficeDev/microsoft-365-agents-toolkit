# CI pipelines

GitHub Actions workflows under [`.github/workflows/`](../../.github/workflows/).

## Per-area workflows

| Workflow | Triggers | Scope |
|----------|---------|-------|
| `unit-test.yml` | PR + push to `dev`/`main` | v3 packages — fx-core, cli, vscode-extension, sdk, sdk-react, ... |
| `ci-next.yml` | PR + push touching `core-next` / `cli-next` | v4 packages — build → lint, format-check, unit-test (80% gate), integration-test |
| `e2e-test-next.yml` | Daily schedule + PR + manual `workflow_dispatch` | cli-next E2E |
| `codeql-analysis.yml` | Weekly + PR | CodeQL scans (TS/JS) |
| `dotnetsdk-ci.yml` | PR touching `packages/dotnet-sdk` | .NET SDK build/test |
| `FunctionExtensionCI.yml` | PR touching `packages/function-extension` | Function extension build/test |

## v4 CI (`ci-next.yml`)

```
build (matrix: ubuntu / windows · node 18 / 20)
  ↓
lint              (ESLint flat config, 0 errors required)
format-check      (Prettier dry-run)
unit-test         (NYC 80% coverage gate)
integration-test  (core-next + cli-next together)
```

Uses `pnpm --filter ./packages/core-next` (directory path) — required because `core-next` and v3 `fx-core` share the npm package name `@microsoft/teamsfx-core`.

## E2E (`e2e-test-next.yml`)

Runs against real (or sandbox) Microsoft 365 / Azure tenants daily plus on demand.

Failure surfaces a GitHub Step Summary with:

- Test stats (passed / failed / skipped).
- Failed test table (name + duration + first failure line).
- Tail of the failure log.

See [`e2e-troubleshooting.md`](../../.github/skills/e2e-troubleshooting/SKILL.md) for triage flow.

## Caching

- `node_modules` and pnpm store cached per workflow run.
- TypeScript `*.tsbuildinfo` not cached — rebuilds are cheap with esbuild.

## Secrets

| Secret | Workflow |
|--------|----------|
| `AZURE_CLIENT_ID` / `AZURE_TENANT_ID` / OIDC | E2E |
| `M365_TEST_*` | E2E |
| `NPM_TOKEN` | publish workflow (separate from CI) |

All secrets are pulled from the repo's GitHub Environments with required reviewers for production publish.

## Dependabot / Renovate

Renovate auto-PRs upstream updates. PRs run through normal CI; major bumps are batched and reviewed manually.
