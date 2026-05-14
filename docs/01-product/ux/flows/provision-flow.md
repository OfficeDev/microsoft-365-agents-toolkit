# Provision flow

```
User → "Provision" (tree view / palette / atk provision --env dev)
  ↓
Engine entry: FxCore.provisionResources (v3) OR provisionOp (v4)
  ↓
loadEnv(envName) → parseProjectYaml → resolveLifecycle → analyzeSteps
  ↓
analyzeSteps inspects driver IDs:
  needsM365 := any step in M365_DRIVERS  (14 driver IDs in v4)
  needsAzure := any step in AZURE_DRIVERS (5 driver IDs in v4)
  unresolvedVars := all ${{VAR}} referenced minus resolved
  ↓
Prerequisites (composable, stop on first failure):
  1. ensureM365Auth      → token, tenantId from JWT
  2. ensureAzureAuth     → identity credential
  3. ensureSubscription  → auto-pick if 1, else prompt
  4. ensureResourceGroup → prompt with default rg-{safeName}{suffix}-{env}
  5. ensureResourceSuffix → reuse or generate 6-char random
  6. confirmProvision    → consent dialog (env, tenant, subscription, RG)
  ↓
executeLifecycle(steps, envMap):
  for each step:
    inject envMap into process.env
    inject ctx.projectPath as PROJECT_PATH
    driver.execute(input, ctx) → outputs
    merge outputs into envMap
    finally: remove injected env vars
  ↓
persistEnv(envName, envMap) → write env/.env.{envName}
  ↓
PostAction[]: "Open Azure portal", "Resources provisioned successfully"
```

## What appears in env after a typical bot provision

```ini
TEAMS_APP_ID=...
TEAMS_APP_TENANT_ID=...
BOT_ID=...
BOT_DOMAIN=mybot{suffix}-dev.azurewebsites.net
BOT_AZURE_APP_SERVICE_RESOURCE_ID=/subscriptions/.../sites/mybot{suffix}-dev
RESOURCE_SUFFIX=ab12cd
AZURE_SUBSCRIPTION_ID=...
AZURE_RESOURCE_GROUP_NAME=rg-mybot{suffix}-dev
```

Secrets (`BOT_PASSWORD`, `AAD_APP_CLIENT_SECRET`, ...) go to `env/.env.{envName}.user`, gitignored.

## Failure modes

| Failure | Likely cause | Mitigation surfaced |
|---------|-------------|---------------------|
| `MissingRequiredOptionError` (CI) | Required env var unset | Error names the missing var |
| `InvalidDriverInput` | YAML key typo or missing field | Zod issue path in error |
| ARM deployment failure | Azure quota / permission / template error | `arm/deploy` returns the deployment failure detail; user can re-run with `--debug` |
| 401 from TDP / Graph | Stale token | `ensureM365Auth` triggers re-auth |
| 403 from Azure | Insufficient role | Surfaced with help link to RBAC docs |
