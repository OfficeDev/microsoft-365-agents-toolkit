# Refactor Assessment: `fx-core` → `core-next`, `cli` → `cli-next`

_Date: 2026-04-17_
_Scope: packages/core-next, packages/cli-next vs. packages/fx-core, packages/cli_

## 1. Inventory (objective)

| Metric                            | Old              | New                               | Ratio |
| --------------------------------- | ---------------- | --------------------------------- | ----- |
| `fx-core/src` LOC / files         | 62,875 / 493     | `core-next/src` 15,658 / 160      | ~25%  |
| `cli/src` LOC / files             | 9,880 / 97       | `cli-next/src` 3,476 / 44         | ~35%  |
| `fx-core/tests` LOC               | 97,127           | `core-next/tests` 10,637          | ~11%  |
| `cli/tests` LOC                   | 7,475            | `cli-next/tests` 5,415            | ~72%  |
| Unit tests (2026-04-17 local run) | —                | core-next: 618 passing; cli-next: 87 passing | green |

## 2. What the refactor did well

1. **Architectural simplification.** The `@hooks` middleware chain + global `TOOLS` singleton are replaced by an explicit, flat pipeline — `validate (Zod) → telemetry start → execute → telemetry end` — in `packages/core-next/src/core/operation.ts`. Every operation receives an injected `AtkContext` (`packages/core-next/src/core/context.ts`) instead of reading globals. Net improvement for testability and readability.
2. **Registry pattern.** Drivers and templates self-register at module load (`drivers/registry.ts`, `templates/registry.ts`), replacing opaque `GeneratorProvider` / `DriverContext` lookups with flat, enumerable registries.
3. **Error model.** `neverthrow` `Result<T, AtkError>` with tagged `user` / `system` kinds is more ergonomic than the prior `FxError` hierarchy plumbed through manual `err()` / `ok()` calls.
4. **Schema-driven inputs.** Zod replaces ad-hoc input validation; e.g. `drivers/builtin/aadApp/create.ts` declares its input shape, constraints, and defaults in one place.
5. **Consolidation.** Former separate packages (`teamsfx-api`, `spec-parser`, `app-manifest`) are folded into `core-next` subtrees (`api/`, `specParser/`, `manifest/`), reducing cross-package churn.
6. **Minimal CLI entry.** `cli-next` uses `commander` with clean command-factory modules (`commands/index.ts`). The old CLI's bespoke `factory/` tree and heavier argparse code are gone.

## 3. Honest concerns

### 3.1 Functional coverage gap is large

The built-in driver set is a small subset of `fx-core/component/driver`:

| Area       | `fx-core/component/driver`                               | `core-next/drivers/builtin` |
| ---------- | -------------------------------------------------------- | --------------------------- |
| arm        | full deploy                                              | `deploy.ts` only            |
| script     | 6 drivers (dotnet / npm / npx / base / scriptDriver)     | `run.ts` only               |
| deploy     | azure + spfx                                             | —                           |
| middleware | telemetry, SWA                                           | —                           |
| devTool / devChannel / typeSpec / m365 / add / share / util | present                    | —                           |
| feature (sso, collaboration, createAuthFiles)            | present                    | —                           |

At the `fx-core/component` level the following are also not yet ported: `coordinator`, `configManager`, `deps-checker`, `migrate` (project upgrade), `developerPortalScaffoldUtils`, `local` debugging, most `generator/*`. The ~25% LOC ratio is consistent with that gap.

### 3.2 Integration with consumers is not started

- `packages/vscode-extension/**` has **zero** imports of `@microsoft/teamsfx-core-next`; it still consumes `@microsoft/teamsfx-core`.
- `packages/server` still consumes `@microsoft/teamsfx-core`.
- Only `cli-next` is wired to `core-next`, so the new stack is not yet proven against the full surface area.

### 3.3 Test coverage has blind spots

From `nyc` on `core-next`:

- `src/project/create.ts`: **0%** (the main "create project" entry point)
- `src/localization/*`: **0%**
- `src/templates/scaffold/*`: ~21% (download 21%, render 15%, scaffolder 18%)
- Several template descriptors (`bot`, `tab`, `connector`, `declarativeAgent`, `engineAgent`, `messageExtension`): 28–60%
- `src/teamsApp/packageBuilder.ts`: ~29%

Pipeline / parser / resolver code is well covered (90%+), but the **scaffolding path users actually hit** is largely untested. The test-LOC drop (97k → 10.6k) suggests many old regression tests were not carried over.

### 3.4 Questionable package.json bits

- `core-next/package.json` pins `"typescript": "~6.0.2"`. TypeScript 6 has not shipped; this looks like a typo for `~5.6.2`. May only be resolving via an installed 5.x via lockfile.
- `core-next/package.json` declares `"repository": "https://github.com/"` (empty). Metadata is inconsistent with other packages in the monorepo.

### 3.5 Coexistence risk

Both stacks live in the tree simultaneously with no `MIGRATION.md`, no deprecation markers on `fx-core`, and no in-repo roadmap for `core-next`. Bug fixes landing in `fx-core` will not auto-propagate; expect drift.

### 3.6 Design details worth scrutinizing

- `createAtkContext` generates correlation IDs with `Math.random()` rather than `crypto.randomUUID()` — acceptable but inconsistent with the rest of the code's rigor (Zod, neverthrow).
- `executeLifecycle` mutates a caller-provided `envMap: Map<string,string>` and auto-injects `ctx.projectPath` into driver configs. Convenient, but it reintroduces implicit behavior that the "explicit flat pipeline" claim was meant to eliminate. Worth documenting.
- `TemplateRegistry` / `DriverRegistry` are process-global singletons. Tests reset them via `testHelper.ts`, which re-introduces the global-state problem the refactor removed from `TOOLS`. Consider making the registries an explicit field on `AtkContext` (or a per-entrypoint instance) so tests don't need shared reset hooks.

## 4. Overall verdict

- **Direction: good.** Operation + AtkContext + registries + Zod + Result is genuinely cleaner than `@hooks` + global `TOOLS` + class hierarchies. The CLI is noticeably simpler.
- **Progress: ~25–35% of the real work.** The skeleton is solid for the scaffold / lifecycle / driver model, but the feature-parity gap with `fx-core` is substantial, VS Code extension and server have not been touched, and several high-traffic code paths have no tests.
- **Biggest risks:**
  1. The refactor stalls at the "demo works" stage before the long tail of drivers / features is ported.
  2. Silent regressions in scaffolding paths that core-next tests do not exercise.
  3. Drift between old and new while both are maintained.

## 5. Recommended next moves

1. Add a `MIGRATION.md` / tracking doc listing every `fx-core` driver and feature with its port status (not started / in progress / ported / won't port).
2. Raise coverage on `src/project/create.ts` and `src/templates/scaffold/*` before migrating more drivers — these are user-visible and currently near 0%.
3. Put a thin `core-next` adapter behind one VS Code extension command (e.g. a single `createProject` path) to prove the integration end-to-end and surface API gaps early.
4. Fix the `typescript: "~6.0.2"` pin and repository metadata in `core-next/package.json`.
5. Decide on registry scoping: either document the singleton contract formally, or move `TemplateRegistry` / `DriverRegistry` into `AtkContext` to preserve the "no global state" promise.
6. Port the missing driver families in priority order implied by real usage: `script/*` build drivers, `deploy/azure`, `arm` (full), then `middleware`, `feature` (sso / collaboration).

## 6. E2E verification (the quality signal that matters)

E2E tests for `cli-next` are the primary real-world quality signal because the unit/integration tests mostly verify the new pipeline shape, not end-to-end parity with `fx-core`. Inspection and local artifacts give a mixed picture.

### 6.1 Infrastructure design — solid

Files: `packages/cli-next/tests/e2e/`.

- **7 suites**: `lifecycle.test.ts`, `cli-syntax.test.ts`, `auth-commands.test.ts`, `add-capability.test.ts`, `mcp-scaffold.test.ts`, `openapi-spec.test.ts`, `gap-features-scaffold.test.ts`.
- **Data-driven lifecycle.** `lifecycle.test.ts` enumerates `templateRegistry.list().filter(t => t.testable !== false)` × each supported language, generating one `describe` per template/lang. Each test runs scaffold → (create RG if Azure) → provision → deploy → publish → final validators → telemetry contract checks.
- **Checkpoint resume.** `TestCheckpoint` + Mocha `this.retries(1)` lets a retry skip already-completed phases — a real feature, not just happy-path.
- **Guaranteed cleanup.** `afterEach` deletes RG, removes project folder, and unpublishes the Teams app via `Promise.allSettled`, so one failure does not leak Azure / catalog resources.
- **Tag-driven validators.** `validators.ts` maps template tags (`bot`, `tab`, `aad`, `function`, `teamsApp`, `publishedApp`) to assertion functions. New template types can add coverage by adding a tag.
- **Real auth in CI.** `ciTokenProvider.ts` acquires tokens via `msal-node` `acquireTokenByUsernamePassword` using env vars the workflow injects from secrets. Non-interactive, no browser.
- **Structured traces.** `TestTracer` implements `TelemetryReporter`, captures every telemetry span in-process, and `verifyTelemetry` runs contract checks. Step logs are written as JSONL under `tests/e2e/logs/` for later analysis.
- **CI plumbing.** `.github/workflows/e2e-test-next.yml` runs on `schedule` (22:00 daily), on `workflow_dispatch`, and on PRs touching `packages/core-next/**` or `packages/cli-next/**`. The `setup` job fans out `lifecycle.test.ts` into a per-template matrix (`lifecycle.test.ts::<templateId>`), so templates run in parallel.
- **Stale resource cleanup.** `tests/e2e/clean.ts` deletes orphaned Azure resources before a run (`continue-on-error`) — protects against flaky prior runs.

### 6.2 Current results — CI is green

Run [#24500383004](https://github.com/OfficeDev/microsoft-365-agents-toolkit/actions/runs/24500383004) on branch `zhiyou/v4` (PR #15663), triggered 2026-04-16:

- **Status: success, 43m 29s.**
- `setup` (4m 11s) ✅ — discovers test cases and expands `lifecycle.test.ts` into per-template matrix entries.
- `execute-case` matrix: **all 22 jobs ✅** — 16 `lifecycle.test.ts::<templateId>` cases (da/basic, da/api-plugin-{no-auth, oauth, bearer, entra-sso}, da/existing-action, da/metaos, bot/echo, tab/basic, ai-agent/{chat, rag-ai-search, rag-custom}, engine-agent/{basic, function-calling, teams-collaborator}, me/search-based) plus the 6 non-lifecycle suites (add-capability, auth-commands, cli-syntax, gap-features-scaffold, mcp-scaffold, openapi-spec).
- `tear-down` (4m 22s) ✅ — stale-resource cleanup completed.
- 22 `test-result-*` artifacts produced, 13–17 KB each for lifecycle suites (consistent with a full scaffold→provision→deploy→publish→validate pipeline writing step logs).
- Warnings: 264, all lint `no-explicit-any` on `packages/core-next/src/api/{error,qm/*}.ts`. Zero test-level failures, zero annotations flagged as errors.

This invalidates the earlier concern about the `Driver "X" is not registered` symptom — the local JSONL logs under `packages/cli-next/tests/e2e/logs/` (4/9–4/13) predate the fix that made CI green. **The dual-instance hypothesis in my prior draft of this section has been either fixed or was never the real problem on the integrated branch.**

### 6.3 Coverage gaps still worth noting

- The matrix covers **16 lifecycle templates**. The template registry contains more (connector, mcp-server, and some Foundry variants flagged `testable: false`). Confirm that the remaining gaps are intentional rather than "broken, silenced."
- Each template runs **one language** per matrix entry. If parity with `fx-core` requires TS and C#/Python coverage where supported, either the parameterization needs to expand to `<templateId>::<language>` or the inner `describe` must iterate languages within the 20-minute budget.
- Non-lifecycle suites (`add-capability`, `auth-commands`, `cli-syntax`, `gap-features-scaffold`, `mcp-scaffold`, `openapi-spec`) are scaffold/CLI-shape checks — no provision/deploy. Their green state is a weaker signal than lifecycle jobs.
- No e2e for `regenerate`, `collaborator grant/list`, `env list/add`, or CLI interactive prompts (only `--help` syntax).
- The 264 lint `no-explicit-any` warnings on core-next `api/error.ts` and `api/qm/*.ts` should be resolved before the API surface is treated as stable.

### 6.4 Quality verdict on e2e

- **Design: 8/10.** Best part of the refactor — checkpoint resume, guaranteed cleanup, tag-driven validators, in-process telemetry capture, matrix fan-out.
- **Execution: 7/10.** The integrated branch passes the full matrix cleanly in ~43 min. Points docked for (a) single-language per template, (b) `testable: false` templates needing an explicit plan, (c) missing commands (regenerate / collaborator / env).
- **Bottom line:** e2e is now a credible gate. The workflow already triggers on `packages/core-next/**` and `packages/cli-next/**` — the next step is to **make it required** in branch protection rather than advisory.

### 6.5 Recommended e2e-specific next steps

1. Make `e2e-test-next.yml` a required status check on PRs touching `packages/core-next/**` or `packages/cli-next/**`.
2. Expand the matrix key to `<templateId>::<language>` so TS/C#/Python variants each run in parallel rather than sharing a 20-min budget.
3. Add scaffold-only e2e for templates currently marked `testable: false`, so the gap is tracked rather than invisible.
4. Add e2e for `regenerate`, `collaborator`, and `env` commands to close the remaining surface-area gaps.
5. Resolve the 264 `no-explicit-any` warnings before tagging the core-next API as stable (see concern #1 in Section 3).

