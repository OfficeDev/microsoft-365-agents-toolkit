# v3 infra naming conventions

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

Conventions extracted from v3 Bicep + parameter files. For archival reference.

## Resource naming

| Concept | Convention | Source |
|---------|-----------|--------|
| Base name | `resourceBaseName` parameter (4–20 chars) used as default for identity, storage, key vault, log analytics, app service, function app names | every archetype |
| Environment suffix | `RESOURCE_SUFFIX` env var combined with `resourceBaseName` in parameter file: `bot${{RESOURCE_SUFFIX}}`, `tab${{RESOURCE_SUFFIX}}`, `copilotconnect${{RESOURCE_SUFFIX}}` | `azure.parameters.json.tpl` |
| SQL resources | suffix from base name: `${resourceBaseName}-sqlserver`, `${resourceBaseName}-db` | Bot/SQL archetype |
| Bot service | defaults to `resourceBaseName` (≤ 42 chars after Mustache) | all bot archetypes |
| Bot display name | separately parameterised; max 42 chars; templates default to `{{appName}}` | all bot archetypes |

## Resource group

- Implicit — templates deploy into a pre-selected or prompted RG via `resourceGroupName` in `m365agents.yml.tpl`.
- Env var: `AZURE_RESOURCE_GROUP_NAME`.
- Default name pattern at prompt time: `rg-{appName}{suffix}-{envName}`.

## Location

| Resource family | Location handling |
|-----------------|-------------------|
| App Service / Function App | `location` parameter, default `resourceGroup().location` |
| Managed Identity | same |
| SQL Server / Database | same |
| Bot Service + channels | hardcoded `"global"` (Azure requirement) |
| Static Web Apps | hardcoded `"centralus"` (template default; service supports more regions) |

Location is **never parameterised in `azure.parameters.json.tpl`** — it flows implicitly from the resource group.

## Env-var output naming

**Bot archetypes have two real variants** (see [archetypes.md §"Bot on App Service with Identity"](archetypes.md#archetype-bot-on-app-service-with-identity) for details):

### Variant A (used by `default-bot` TS/JS, `message-extension-v2` TS, `foundry-agent-to-m365`, `teams-agent-with-data-custom-api-v2`)

```
AZURE_APP_SERVICE_RESOURCE_ID
BOT_DOMAIN
BOT_ENDPOINT
BOT_ID
BOT_TENANT_ID
```

### Variant B (used by `weather-agent`, `custom-copilot-*` family, `basic-custom-engine-agent`, `declarative-agent-with-action-from-scratch-oauth`)

```
BOT_AZURE_APP_SERVICE_RESOURCE_ID
BOT_DOMAIN
BOT_ID
BOT_TENANT_ID
# no BOT_ENDPOINT
```

### Bot/SQL archetype (Variant B + SQL extras)

```
BOT_AZURE_APP_SERVICE_RESOURCE_ID
BOT_DOMAIN
BOT_ID
BOT_TENANT_ID
SQL_SERVER_FQDN
SQL_DATABASE_NAME
```

### Tab archetype:

```
AZURE_APP_SERVICE_RESOURCE_ID
TAB_DOMAIN
TAB_ENDPOINT
```

### Function-API archetype:

```
API_FUNCTION_ENDPOINT
API_FUNCTION_RESOURCE_ID
OPENAPI_SERVER_URL
```

### Graph connector:

```
AZURE_FUNCTION_RESOURCE_ID
```

### Static web app:

```
AZURE_STATIC_WEB_APPS_RESOURCE_ID
ADDIN_DOMAIN
ADDIN_ENDPOINT
```

## What v4 *may* preserve

- The **env-var output names** above (envMap back-compat for existing user projects) — but v4 must accept BOTH bot variants.
- The **shape of the prompt defaults** (`rg-{appName}{suffix}-{envName}` is reasonable for any Azure tool).

## What v4 *must not* import

- The hardcoded locations (`centralus`, `global` is required, others are template artefacts).
- The 2021-era API versions.
- The lack of managed identity in non-bot archetypes.
- The `Standard_LRS` storage default.
- The implicit-only `location` handling — surface region selection in the question model instead.

Re-derive from current Azure best practice (Well-Architected Framework, deployment-guidance docs) for any v4 Bicep work.
