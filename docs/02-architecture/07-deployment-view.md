# 7 — Deployment view

## The toolkit itself

The toolkit runs **locally**, on the developer's machine or a CI runner. There is no toolkit-hosted backend service. All persistent state is local:

| Location | Contents | Lifecycle |
|----------|----------|-----------|
| Editor / shell | The toolkit binaries (extension, CLI bundle) | Reinstalled per upgrade |
| `~/.fx/account/` | MSAL token cache (AES-256-GCM with optional keytar) | Survives upgrades · shared v3↔v4 |
| Project `.fx/` (v3) | Local project state | Per project |
| Project `env/` | `.env.{envName}` (non-secret) and `.env.{envName}.user` (secret) | Per project, per environment |
| Project `node_modules/` (after deploy build) | Build outputs zipped for Kudu push | Transient |

## What the toolkit deploys *for users*

This is the developer's app, not the toolkit. Topology depends on template; common shapes:

### Bot / CEA on Azure App Service

```
                    ┌─────────────────┐
                    │  Bot Framework  │
                    │  (channel reg)  │
                    └────────┬────────┘
                             │
                    ┌────────▼────────┐
   user/Teams ──▶  │  App Service    │  ◀── Kudu zip deploy (azureAppService/zipDeploy)
                    │ (Web App)       │
                    └────────┬────────┘
                             │
                  ┌──────────┼──────────┐
                  ▼                     ▼
              ┌────────┐           ┌────────┐
              │ Entra  │           │  App   │
              │ ID app │           │Insights│
              └────────┘           └────────┘
```

### Bot / CEA on Azure Functions

Same shape, App Service replaced by a Function App. Driver: `azureFunctions/zipDeploy`.

### Declarative Agent (no backend)

No Azure required. Topology is just:

```
   user/M365 Copilot ──▶ Tenant App Catalog ◀── teamsApp/publishAppPackage (MS Graph)
```

### Declarative Agent + API plugin

DA in tenant catalog plus a backing service (App Service / Functions) for the API. The API plugin manifest references the API endpoint URL.

### Graph Connector

```
                            ┌──────────────────┐
                  ◀─ ingest ┤ Microsoft Graph  │
                            └──────────────────┘
                                    ▲
                                    │
           Function App (your code) │
              ────────────┬─────────┘
                          │
                    ┌─────▼─────┐
                    │ External  │
                    │ data src  │
                    └───────────┘
```

## Detailed catalogue

For the full Bicep template catalogue, naming conventions, networking, and identity patterns the toolkit emits, see [03-infrastructure/](../03-infrastructure/README.md).
