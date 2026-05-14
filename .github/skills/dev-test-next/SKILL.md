---
name: dev-test-next
description: "Local development, testing, and verification workflow for core-next and cli-next packages. Use when: building core-next or cli-next, running unit/integration/e2e tests, verifying a feature end-to-end, running lifecycle tests locally, setting up developer credentials, debugging test failures, checking test results."
argument-hint: "Describe what you want to build, test, or verify"
---

# Dev Test Next — Local Development & Verification

> **For the full v4 contribution workflow** (design-first, implement, test, docs, PR) see the [`vibe-coding`](../vibe-coding/SKILL.md) skill. This skill covers the **build / test / verify** mechanics only. For the per-layer testing rationale (why integration tests outweigh unit tests for lifecycle code) see [ADR 0007 — Inverted test pyramid](../../docs/02-architecture/09-architecture-decisions/0007-inverted-test-pyramid-for-lifecycle.md) and [`docs/05-engineering/testing-strategy.md`](../../docs/05-engineering/testing-strategy.md). For the E2E test architecture (test flow, key files, CI matrix) see [`docs/06-operations/e2e-strategy.md`](../../docs/06-operations/e2e-strategy.md).

## When to Use

- Building or testing `packages/core-next` or `packages/cli-next`
- Running unit, integration, or E2E tests locally
- Verifying a feature works end-to-end before pushing
- Debugging test failures from CI

## Phase 0 — Design First (v4 Only)

**v4 (`core-next`, `cli-next`) is design-first.** Before writing or modifying code for any non-trivial change, locate or write the design page in `docs/`:

| Change kind | Where the design lives |
|-------------|-----------------------|
| New / changed contract or entity | [`docs/04-specs/data-model/entities/`](../../docs/04-specs/data-model/entities/README.md) |
| New / changed operation, driver, lifecycle stage, runtime view | [`docs/02-architecture/06-runtime-views.md`](../../docs/02-architecture/06-runtime-views.md), [`docs/05-engineering/cross-cutting/`](../../docs/05-engineering/cross-cutting/README.md) |
| Architectural decision | New ADR under [`docs/02-architecture/09-architecture-decisions/`](../../docs/02-architecture/09-architecture-decisions/README.md) using the [template](../../docs/02-architecture/09-architecture-decisions/adr-template.md) |
| New CLI surface, command group, flow | [`docs/01-product/ux/surfaces/cli.md`](../../docs/01-product/ux/surfaces/cli.md), [`docs/01-product/ux/flows/`](../../docs/01-product/ux/flows/README.md) |
| New capability / template | [`docs/01-product/capabilities-matrix.md`](../../docs/01-product/capabilities-matrix.md) + [`.dev/features.json`](../../.dev/features.json) |

**Land docs in the same PR as code.** If implementation diverges from the design, update the design in the same PR — do not merge code that contradicts a design page. Skip Phase 0 only for trivial changes (bug fix, dependency bump, doc-only edit, test-only edit).

See [codebase.instructions.md §Source-of-Truth Workflow](../../instructions/codebase.instructions.md).

## Package Overview

| Package | Path | Role |
|---------|------|------|
| **core-next** | `packages/core-next` | Engine: templates, drivers, lifecycle operations, project scaffolding |
| **cli-next** | `packages/cli-next` | CLI: commands, auth, output formatting, E2E tests |

Dependency: `core-next` → `cli-next` (cli-next depends on core-next)

**TypeScript:** Both packages use TypeScript 6.0 (`~6.0.0`) with `moduleResolution: "bundler"` and `module: "commonjs"`. The deprecated `node10` resolution is no longer used.

## Quick Reference

```bash
# Build both packages (always build core-next first)
cd <repo-root>
npm run setup:next                            # install + build both

# Or build individually — postbuild auto-runs eslint --fix + prettier
pnpm --filter ./packages/core-next build
pnpm --filter ./packages/cli-next build

# Link CLI globally for local testing
cd packages/cli-next && pnpm link --global    # makes `atk` command available
```

> **Note:** `build` includes a `postbuild` hook that automatically runs `eslint --fix`
> on `src/` and `tests/`. Formatting is handled by `eslint-plugin-prettier` (single source
> of truth). No separate format step needed after building.

## Test Pyramid

### 1. Unit Tests (no credentials needed)

```bash
# core-next — ~500 unit tests
cd packages/core-next && npm run test:unit

# cli-next — ~80 unit + ~55 integration tests
cd packages/cli-next && npm run test:unit
```

- Framework: Mocha + Chai + Sinon, ts-node/register via `.mocharc.js`
- Coverage: NYC (reports inline after test run)
- Pattern: `tests/unit/**/*.test.ts`
- E2E uses separate `.mocharc.e2e.js` (chains `ts-node/register` + `tests/e2e/setup.ts`)

### 2. Integration Tests (no credentials needed)

```bash
cd packages/cli-next && npm run test:integration
```

- Pattern: `tests/integration/**/*.test.ts`
- Tests cross-module wiring without Azure or M365 services

### 3. E2E Tests

Two test suites with different credential requirements:

| Suite | Command | Credentials | What it tests |
|-------|---------|-------------|---------------|
| **CLI syntax** | `npm run test:e2e:cli` | None (just needs `atk` on PATH) | Binary arg parsing, `--version`, `--help`, scaffold smoke tests (bot, tab, DA, ME) |
| **MCP scaffold** | via `npm run test:e2e` | None | `atk new da mcp-remote` scaffold, MCP server URL, file validation, error cases |
| **Auth commands** | via `npm run test:e2e` | None | `atk auth` subcommand help parsing (show, login azure/m365, logout) |
| **Add capability** | via `npm run test:e2e` | None | `atk add capability` on DA project (web-search), error cases |
| **Lifecycle** | `npm run test:e2e:lifecycle` | Azure + M365 login | Full scaffold → provision → deploy → validate per template |
| **All E2E** | `npm run test:e2e` | Azure + M365 login | All 5 test files |
| **Cleanup** | `npm run test:e2e:clean` | Azure login | Deletes stale Azure resource groups |

Test files: `packages/cli-next/tests/e2e/`

## Local E2E Setup

### CLI Syntax Tests (no credentials)

```bash
cd packages/cli-next
pnpm link --global          # register `atk` command
npm run test:e2e:cli        # runs cli-syntax.test.ts
```

### Lifecycle Tests (credentials required)

**Step 1 — Login once using the CLI:**

```bash
atk auth login azure        # opens browser, caches token to ~/.fx/account/
atk auth login m365          # opens browser, caches M365 token
```

Tokens persist in `~/.fx/account/` (AES-encrypted MSAL cache). No env vars needed locally.

**Step 2 — Run:**

```bash
npm run test:e2e:lifecycle
```

**How auth works locally:**
- Lifecycle operations (provision, deploy) use `createTokenProvider()` which picks
  up cached tokens from `atk auth login` via `AzureAccountManager` (MSAL silent flow).
- Resource group management uses `DefaultAzureCredential` which discovers the same
  cached credentials automatically.
- Subscription defaults to `af46c703-f714-4f4c-af42-835a673c2b13`
  (Teams Cloud – E2E Testing, TTL = 1 day). Override with `AZURE_SUBSCRIPTION_ID` env var.

### Run a Specific Template

Use Mocha `--grep` to filter by template ID (tests are named `E2E lifecycle: <id> [<lang>]`):

```bash
# Single template
npx mocha --require tests/e2e/setup.ts --timeout 1200000 \
  --grep "bot/echo.*TypeScript" "tests/e2e/lifecycle.test.ts"

# All bot templates
npx mocha --require tests/e2e/setup.ts --timeout 1200000 \
  --grep "bot/" "tests/e2e/lifecycle.test.ts"

# All TypeScript templates
npx mocha --require tests/e2e/setup.ts --timeout 1200000 \
  --grep "TypeScript" "tests/e2e/lifecycle.test.ts"
```

### Cleanup Stale Resources

```bash
npm run test:e2e:clean                    # sweep resources older than 2h
npx ts-node tests/e2e/clean.ts --run-id <id>  # cleanup a specific run
```

Resources are tagged with `atk-test=true` and `created-at` timestamp for reliable tag-based cleanup.

## CI Pipelines

| Workflow | File | Trigger | What it runs |
|----------|------|---------|--------------|
| **CI Next** | `.github/workflows/ci-next.yml` | PR/push to `dev`/`release/**` | Build → Lint → Unit tests → Integration tests |
| **E2E Test Next** | `.github/workflows/e2e-test-next.yml` | PR/push, nightly, manual | Matrix of E2E test files with Azure creds |

### CI Next Verification

CI Next runs automatically on every PR touching `packages/core-next/**` or `packages/cli-next/**`.
To replicate locally:

```bash
# Full CI check
pnpm --filter ./packages/core-next build
pnpm --filter ./packages/cli-next build
pnpm --filter ./packages/core-next lint
pnpm --filter ./packages/cli-next lint
pnpm --filter ./packages/core-next test:unit
pnpm --filter ./packages/cli-next test:unit
pnpm --filter ./packages/cli-next test:integration
```

### E2E in CI vs Local

| Aspect | CI (`CI_ENABLED=true`) | Local (default) |
|--------|------------------------|-----------------|
| Auth | Env vars (service principal / username+password) | `atk auth login` cached tokens |
| Subscription | `AZURE_SUBSCRIPTION_ID` env var | `af46c703-...` (E2E Testing, TTL 1d) |
| Credential for RG mgmt | `UsernamePasswordCredential` / `ClientSecretCredential` | `DefaultAzureCredential` |
| Env vars | All required, throws on missing | All optional |
| Retry | Mocha `.retries(1)` + workflow rerun up to 5× | Mocha `.retries(1)` + checkpoint resume |
| Cleanup | Per-test + tear-down job + stale sweep | Per-test + manual `npm run test:e2e:clean` |

## Verification Checklist

Before pushing a change to core-next or cli-next, verify:

1. **Build** — `pnpm --filter ./packages/core-next build && pnpm --filter ./packages/cli-next build`
   (auto-runs lint fix + format via `postbuild` hook)
2. **Lint check** — `pnpm --filter ./packages/core-next lint && pnpm --filter ./packages/cli-next lint`
   (should show 0 errors; warnings are acceptable)
3. **Unit tests** — `pnpm --filter ./packages/core-next test:unit && pnpm --filter ./packages/cli-next test:unit`
4. **Integration tests** — `pnpm --filter ./packages/cli-next test:integration`
5. **(Optional) E2E CLI** — `cd packages/cli-next && npm run test:e2e:cli`
6. **(Optional) E2E Lifecycle** — `cd packages/cli-next && npm run test:e2e:lifecycle` (after `atk auth login`)

Steps 1–4 match what CI Next runs. Steps 5–6 match what E2E Test Next runs.

> **Shortcut:** Since `build` now auto-formats via eslint-plugin-prettier, steps 1–2 can
> be collapsed into just `build` + `lint`. No separate format-check needed.

## E2E Architecture

For the full E2E architecture (test flow diagram, key design decisions, file layout, CI matrix, template language matrix), see [`docs/06-operations/e2e-strategy.md`](../../docs/06-operations/e2e-strategy.md).

Quick layout reference:

```
tests/e2e/
├── infra/
│   ├── config.ts        — Auth config (CI vs local mode)
│   ├── tracer.ts        — TestTracer + verifyTelemetry()
│   ├── checkpoint.ts    — Resume-from-failure retry
│   ├── azure.ts         — Tagged resource group create/delete/sweep
│   ├── validators.ts    — Tag-driven post-provision validators
│   └── testContext.ts   — Creates AtkContext with real auth
├── setup.ts             — Mocha root hook
├── clean.ts             — Standalone cleanup script
├── lifecycle.tests.ts   — Data-driven: templateRegistry.list() × languages
└── cli-syntax.tests.ts  — Subprocess CLI verification
```

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| `atk: command not found` | `cd packages/cli-next && pnpm link --global` |
| `Cannot find module '@microsoft/teamsfx-core'` | Build core-next first: `pnpm --filter ./packages/core-next build` |
| `E2E config: missing required environment variable` | You're in CI mode. Unset `CI_ENABLED` or set all required vars |
| `DefaultAzureCredential: no credential` | Run `atk auth login azure` first |
| Tests hang at provision | Check `atk auth login azure` is still valid; tokens expire after ~1h |
| `SyntaxError: Unexpected token ':'` | Missing ts-node/register — `.mocharc.js` should load it automatically |
| Resource group not cleaned up | Run `npm run test:e2e:clean` to sweep stale resources |
| Coverage threshold failure (exit code 1) | Unit tests passed but NYC coverage < 50% — this is pre-existing, not blocking |
