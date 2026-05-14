# Lifecycle engine

Source: [`packages/core-next/src/lifecycle/`](../../../packages/core-next/src/lifecycle/).

## Pipeline

```
parseProjectYaml(yamlPath)              → RawProjectModel
  │
resolveLifecycle(model, registry)       → DriverStep[] (resolved against DriverRegistry)
  │
analyzeSteps(steps, envMap?)            → LifecycleAnalysis { needsM365, needsAzure, driverIds, unresolvedVars }
  │
prerequisites (composable)              → tokens, subscription, RG, suffix, consent
  │
executeLifecycle(steps, envMap, ctx, progress?) → LifecycleResult
  │
persistEnv(envName, envMap)
```

## Files

| File | Purpose |
|------|---------|
| `parser.ts` | `parseProjectYaml()` — reads `m365agents.yml` into `RawProjectModel` |
| `resolver.ts` | `resolveLifecycle()` — matches actions to drivers via `DriverRegistry` |
| `executor.ts` | `executeLifecycle()` — runs resolved actions in sequence; optional `LifecycleProgress` callbacks |
| `types.ts` | `DriverStep`, `RawProjectModel`, `LifecycleResult`, `LifecycleProgress`, `LifecycleAnalysis`, `LifecycleOperationResult`, `PostAction`, `M365TenantInfo`, `AzureAccountInfo`, `ResourceGroupInfo` |
| `analyze.ts` | `analyzeSteps()` — driver introspection |
| `progress.ts` | `createProgressAdapter(ui)` + `silentProgress` |
| `prerequisites.ts` | Composable auth gates |
| `operations.ts` | `provisionOp`, `deployOp`, `publishOp` — `defineOperation` wrappers |
| `index.ts` | Barrel exports |

## Executor envMap injection

Before each driver call, the executor:

1. Auto-injects `ctx.projectPath` as `PROJECT_PATH` if absent.
2. Temporarily syncs envMap entries into `process.env` so drivers loading external files (ARM parameter JSON, AAD manifest templates) can resolve `${{VAR}}` placeholders produced by earlier steps.
3. Runs the driver.
4. `finally`: removes injected env vars to avoid state leaking.

## Operations

| Operation | Pipeline |
|-----------|----------|
| `provisionOp` | loadEnv → parseYAML → analyzeSteps → ensureM365Auth → ensureAzureAuth → ensureSubscription → ensureResourceGroup → confirmProvision → executeLifecycle → persistEnv |
| `deployOp` | loadEnv → parseYAML → confirmDeploy → executeLifecycle → persistEnv |
| `publishOp` | loadEnv → parseYAML → executeLifecycle → persistEnv |

Each is a `defineOperation()` wrapper composing prerequisites + executor + env persistence. Returns `Result<{ postActions: PostAction[] }, AtkError>`.

## Composable prerequisites

`prerequisites.ts`:

- `ensureM365Auth(ctx)` → acquires M365 token, extracts tenantId from JWT.
- `ensureAzureAuth(ctx)` → triggers Azure login.
- `ensureSubscription(ctx, envMap)` → auto-pick if 1, prompt if many.
- `ensureResourceGroup(ctx, envMap, subId, projectName, envName)` → prompts with `rg-{safeName}{suffix}-{env}` default; also triggered when `AZURE_RESOURCE_GROUP_NAME` is present but empty.
- `ensureResourceSuffix(envMap)` → generate / reuse 6-char random.
- `confirmProvision(ctx, envName, m365Info?, azureInfo?)` → consent dialog.
- `confirmDeploy(ctx, envName)` → skipped for local/testtool/playground/sandbox envs.

## Driver introspection

`analyze.ts`:

- `analyzeSteps(steps, envMap?)` → `LifecycleAnalysis { needsM365, needsAzure, driverIds, unresolvedVars }`.
- Uses `M365_DRIVERS` set (14 driver IDs in v4) and `AZURE_DRIVERS` set (5 driver IDs).
- Collects `${{VAR}}` placeholders, checks against envMap.

## Progress

`progress.ts`:

- `createProgressAdapter(ui, title?)` — bridges `LifecycleProgress` to platform `createProgressBar`.
- `silentProgress` — no-op for CI/testing.

## Post-actions

Operations return `PostAction[]` for the consumer to render:

- `{ type: "openUrl", message, url }` — Azure portal link, etc.
- `{ type: "showMessage", message }` — completion summary.

## Tests

`tests/unit/lifecycle/` — 60 tests across `analyze.test.ts`, `progress.test.ts`, `prerequisites.test.ts`, `operations.test.ts`. Integration: `tests/integration/lifecycleExecution.test.ts` — YAML parse → driver execution → env-var chaining between steps.
