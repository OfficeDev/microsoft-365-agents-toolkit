# Env files

Per-environment configuration storage. Loaded by [`environment/envManager.ts`](../../../packages/core-next/src/environment/envManager.ts) and persisted after each lifecycle stage.

## Files

| File | Contents | In source control |
|------|----------|---------------------|
| `env/.env.{envName}` | Non-secret values: resource IDs, domains, tenant IDs | yes |
| `env/.env.{envName}.user` | Secret values: client secrets, bot passwords, API keys | **no** (gitignored) |

`{envName}` is typically `local`, `dev`, `prod`. Custom names are supported via `atk env add <name>`.

## Format

Standard `KEY=VALUE` `.env` syntax. No quoting required; values are read verbatim. No interpolation in env files themselves — `${{VAR}}` resolution happens at YAML parse time, not env-file load time.

```ini
TEAMS_APP_ID=12345678-aaaa-bbbb-cccc-dddddddddddd
BOT_ID=...
BOT_DOMAIN=mybot{suffix}-dev.azurewebsites.net
RESOURCE_SUFFIX=ab12cd
AZURE_SUBSCRIPTION_ID=...
AZURE_RESOURCE_GROUP_NAME=rg-mybot{suffix}-dev
BOT_AZURE_APP_SERVICE_RESOURCE_ID=/subscriptions/.../sites/mybot{suffix}-dev
```

## Read / write contract

| Operation | Reads | Writes |
|-----------|-------|--------|
| `loadEnv(envName)` | both files into `envMap` | — |
| `executeLifecycle(...)` | uses `envMap` to resolve `${{VAR}}` and inject into `process.env` per-step | merges driver outputs into `envMap` |
| `persistEnv(envName, envMap)` | — | splits into non-secret / secret files based on key prefix (`SECRET_*`) and credential keyword detection |

Secret detection uses the **secret masker** keyword list; values whose keys end with credential suffixes (`PASSWORD`, `SECRET`, `KEY`, `TOKEN`, ...) are routed to `.user`.

## Conventions

| Prefix | Convention |
|--------|-----------|
| `TEAMS_APP_*` | Teams app identity |
| `BOT_*` | Bot identity & runtime |
| `AAD_APP_*` | Backend Entra ID app |
| `AZURE_*` | Azure subscription / resource group / suffix |
| `SECRET_*` | Force routing to `.user` file |
| `M365_*` | M365 token-derived values (tenant, etc.) |

## Multi-env workflow

```bash
atk env list                   # show envs
atk env add prod               # create new env (copies dev defaults where sensible)
atk provision --env prod       # provision against prod env
atk env reset --env dev        # clear dev env (re-runs of provision will recreate)
```

## Failure modes

| Failure | Cause |
|---------|-------|
| Driver expects `${{X}}` but X is unset | env file missing or earlier step failed before producing X |
| Secret value committed to non-`.user` file | `persistEnv` keyword detection bypassed by unusual key naming → use `SECRET_` prefix |
| Stale value blocks redeploy | `atk env reset` |
