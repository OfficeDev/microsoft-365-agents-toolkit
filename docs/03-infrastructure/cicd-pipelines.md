# CI/CD pipelines

The toolkit ships **pipeline templates** that users can opt into for repeatable provision/deploy/publish in CI.

## What ships

| Platform | Where |
|----------|-------|
| GitHub Actions | Sample workflows in scaffolded projects under `.github/workflows/` (template-dependent) |
| Azure DevOps | Sample pipelines under `.ado/` (template-dependent) |

Both rely on the CLI in **non-interactive mode** (`CI_ENABLED=true`).

## Auth in CI

CI runs cannot do interactive MSAL flows. Two paths:

| Path | When | How |
|------|------|-----|
| Service principal | Headless build agents | Set `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` env vars (or use OIDC/federated credentials). Picked up by `cli-next/src/auth/azureLoginCI.ts`. |
| Workload identity federation | GitHub Actions, ADO | OIDC token exchange — recommended over long-lived secrets. |

The CLI auto-detects CI mode (env vars: `CI`, `TF_BUILD`, `GITHUB_ACTIONS`) and routes through `createTokenProvider()` accordingly.

## Required env vars per stage

| Stage | Required env vars |
|-------|-------------------|
| `provision` | Azure credentials · `AZURE_SUBSCRIPTION_ID` · `AZURE_RESOURCE_GROUP_NAME` (optional — defaults from `ensureResourceGroup`) · `M365_*` from M365 token |
| `deploy` | Azure credentials only (assumes provision already produced resource IDs in env files) |
| `publish` | M365 admin credentials with app catalog publish permission |

Secrets must be configured as pipeline secrets, never committed.

## Reproducibility in CI

`CI_ENABLED=true` short-circuits every interactive `CLIUserInteraction` prompt — defaults are returned, missing required values surface as `MissingRequiredOptionError` (a `UserError`) so the pipeline fails with a clear message instead of hanging.

See [`cli.instructions.md`](../../.github/instructions/cli.instructions.md) §"Interactive vs Non-Interactive".
