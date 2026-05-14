# E2E strategy

End-to-end tests verify the full lifecycle (scaffold → provision → deploy → publish) against real (or sandbox) Microsoft 365 + Azure tenants.

## v4 Architecture

### Test Flow

```
lifecycle.test.ts  ──►  templateRegistry.list()
                              │
                    for each (template × language):
                              │
          ┌───────────────────┼────────────────────┐
          ▼                   ▼                    ▼
      scaffold           provision             deploy
   (createProjectOp)    (provisionOp)        (deployOp)
          │                   │                    │
          ▼                   ▼                    ▼
    template zip         m365agents.yml        m365agents.yml
    extraction           provision: steps      deploy: steps
    + Mustache render    ├─ teamsApp/create     ├─ cli/runNpmCommand
                         ├─ aadApp/create       ├─ cli/runDotnetCommand
                         ├─ arm/deploy          ├─ azureFunctions/zipDeploy
                         └─ oauth/register      └─ azureAppService/zipDeploy
```

### Key Files

| File | Purpose |
|------|---------|
| `packages/cli-next/tests/e2e/lifecycle.test.ts` | Data-driven E2E test generator |
| `packages/cli-next/tests/e2e/infra/testContext.ts` | Creates mock context for operations |
| `packages/cli-next/tests/e2e/infra/config.ts` | CI env var configuration |
| `packages/cli-next/tests/e2e/infra/validators.ts` | Tag-driven post-provision validators |
| `packages/cli-next/tests/e2e/infra/checkpoint.ts` | Resume-from-failure retry (skip completed phases) |
| `packages/cli-next/tests/e2e/infra/tracer.ts` | TestTracer (telemetry capture) + verifyTelemetry() |
| `packages/cli-next/tests/e2e/infra/azure.ts` | Tagged resource group create/delete/sweep |
| `packages/core-next/src/drivers/builtin/index.ts` | Driver registry (all builtin drivers) |
| `packages/core-next/src/templates/descriptors/` | Template descriptors (id, languages, tags, testable) |
| `packages/core-next/templates/fallback/*.zip` | Template zip archives (ts, js, csharp, python, common) |
| `.github/workflows/e2e-test-next.yml` | CI workflow (matrix of test files) |

### Key Design Decisions

- **Programmatic API primary** — tests call `runOperation(provisionOp, ctx, input)` directly, not subprocesses
- **Data-driven** — `templateRegistry.list()` generates tests; new templates get E2E coverage automatically
- **Lifecycle-aware** — tests read `m365agents.yml` once and detect which lifecycle sections exist; templates without the relevant section skip those phases gracefully
- **Env bootstrapping** — provision phase ensures `env/` dir + `.env.dev` with required vars exist before lifecycle execution
- **`testable` filtering** — templates with `testable: false` in their descriptor are excluded from E2E generation
- **Tag-based cleanup** — Azure RGs tagged with `atk-test`, `test-run-id`, `created-at` for reliable sweep
- **Checkpoint retry** — on Mocha retry, completed phases (scaffold, create-rg) are skipped
- **Telemetry verification** — `verifyTelemetry()` runs 6 contract rules after each test

### CI Matrix Structure

The CI runs each matrix entry as `lifecycle.test.ts::<template-id>`. Each matrix job runs **all language variants** for that template (the Mocha `--grep` matches all languages).

### Template Language Matrix

Templates live in zip files under `packages/core-next/templates/fallback/`:

| Zip | Languages |
|-----|-----------|
| `ts.zip` | TypeScript templates |
| `js.zip` | JavaScript templates |
| `csharp.zip` | C# templates (use `cli/runDotnetCommand` driver) |
| `python.zip` | Python templates |
| `common.zip` | Language-agnostic templates (DA manifests, configs) |

Each template descriptor declares its supported languages. The `convertToLangKey()` function maps: `TypeScript` → `ts`, `JavaScript` → `js`, `CSharp` → `csharp`, `Python` → `python`.

---

## v4 Coverage Status

cli-next E2E tests live in `packages/cli-next/tests/e2e/`. Driven by `ci-next.yml` (integration job) and `e2e-test-next.yml` (daily + PR + manual).

| Template | E2E status |
|----------|-----------|
| `default-bot` (TS / JS / Python) | ✓ |
| `basic-tab` (TS) | ✓ |
| `custom-copilot-basic` (TS / JS / Python) | ✓ |
| `cea/basic` (TS / JS / Python) | ✓ |
| `cea/weather` (TS / JS / Python) | ✓ |
| `da/basic` (common) | ✓ |
| `connector/graph-connector` (TS) | ✓ |
| `me/v2` (TS / Python) | tracked |
| RAG variants | tracked |

## Scaffold-only E2E

A subset of templates is verified for **scaffold only** (no provision / deploy) where the lifecycle requires interactive input or real-tenant resources that aren't suitable for daily CI:

- `da/api-plugin-from-spec` — needs `apiSpecPath` input.
- `da/graph-connector` (registry artifact gap).
- `me/from-spec` — needs spec input.

## OpenAPI spec-parser E2E

9 spec-parser E2E tests verify parsing, validation, filtering, and optimisation against real OpenAPI documents.

## v3

Comprehensive E2E suites in `packages/cli/tests/e2e/` (and equivalent for the VS Code extension test harness). Driven by the v3 feature registry (`.dev/features.json` × `featureRegistry.ts`).

## Failure summary

Failed E2E runs publish a GitHub Step Summary including:

- Stats (passed / failed / skipped / duration).
- Failed test table (name, duration, first failure line).
- Tail of the failure log.

For failure diagnosis and common fix patterns, see [`.github/skills/e2e-troubleshooting/SKILL.md`](../../.github/skills/e2e-troubleshooting/SKILL.md).

## Test data hygiene

- Each test creates a uniquely-suffixed resource group; cleanup runs in the test's `after()` hook.
- A daily sweeper cleans up orphaned resource groups whose suffix matches the e2e pattern.
- Test M365 tenants are dedicated to the toolkit and not used for any production data.

## Local E2E

See [`dev-test-next` skill](../../.github/skills/dev-test-next/SKILL.md) for local setup instructions (credentials, subscription, running specific templates, cleanup).
