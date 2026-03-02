# Environment Configuration Reference

Use this reference when configuring env files for the `foundry-agent` workflow.

## Scope

- Configure existing Foundry values in generated env files.
- Keep local debug and cloud deployment values consistent.
- Do not invent endpoint, IDs, tenant, subscription, or secrets.

## Environment Files

| File | Purpose |
| ---- | ------- |
| `env/.env.local` | Local non-secret values (including tunnel URL for bots) |
| `env/.env.local.user` | Local secrets |
| `env/.env.dev` | Cloud/dev non-secret values |
| `env/.env.dev.user` | Cloud/dev secrets |

## Lifecycle Files

| File | Used by | Notes |
| ---- | ------- | ----- |
| `m365agents.local.yml` | `--env local` | Local provisioning and deploy |
| `m365agents.yml` | `--env dev` | Cloud provisioning and deploy |

`atk provision` and `atk deploy` run tasks from these YAML files:

```bash
# local flow
atk provision --env local -i false
atk deploy --env local -i false

# cloud flow
atk provision --env dev --resource-group <rg> --region <region> -i false
atk deploy --env dev -i false
```

## Critical Runtime Flow (`.localConfigs`)

For local runs, runtime values are typically written to `.localConfigs` by lifecycle tasks.

**Important:** The running service often reads `.localConfigs`, not `env/.env.local` directly.

If local runtime values look wrong:

1. Verify values in `env/.env.local` / `env/.env.local.user`.
2. Verify mappings in `m365agents.local.yml`.
3. Regenerate runtime config: `atk deploy --env local -i false`.
4. Re-start the local service.

## Required Value Checklist

Before local debug:

- Foundry endpoint and agent identifiers are set.
- `BOT_ENDPOINT` (for bot projects) points to current dev tunnel URL.
- Tenant/app IDs are present for Teams/M365 launch URLs.

Before cloud provision/deploy:

- `AZURE_SUBSCRIPTION_ID` exists in `env/.env.dev`.
- `atk auth list` shows valid Azure and M365 login.
- Resource group and region match the target deployment.

## Fast Failure Checks

If `atk provision` or `atk deploy` fails:

1. Run `atk auth list`.
2. Run `atk validate --env <env> -i false`.
3. Re-check missing `${{VAR_NAME}}` references from lifecycle YAML against env files.
4. Re-run the failed command with the same `--env` and `-i false`.

## Common Variables

| Variable | Usage |
| -------- | ----- |
| `TEAMS_APP_ID` | Teams sideload URL |
| `M365_APP_ID` | Microsoft 365 host URLs |
| `TENANT_ID` | Host launch URL context |
| `BOT_ENDPOINT` | Public bot endpoint for Teams reachability |
| `AZURE_SUBSCRIPTION_ID` | Cloud provisioning target |
