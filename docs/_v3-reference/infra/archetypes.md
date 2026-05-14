# v3 infra archetypes

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

The six recurring Bicep patterns extracted from `templates/vsc/{lang}/{template}/infra/`. Every shipped template maps to one of these.

For the template → archetype mapping, see [template-inventory.md](template-inventory.md).
For naming and location conventions, see [naming-conventions.md](naming-conventions.md).

---

## Archetype: Bot on App Service with Identity

**Used by:** 11 templates — `default-bot`, `basic-custom-engine-agent`, `weather-agent`, `custom-copilot-basic`, `custom-copilot-rag-*` (3 variants), `message-extension-v2`, `declarative-agent-with-action-from-scratch-oauth`, `foundry-agent-to-m365`, `teams-agent-with-data-custom-api-v2`.

### Output schema variants — important

This archetype has **two real output-schema variants** that v4 must handle if it wants envMap back-compat. The split is real, not a documentation artefact:

#### Variant A: `AZURE_APP_SERVICE_RESOURCE_ID` + with `BOT_ENDPOINT`

**Used by:** `default-bot` (TS+JS), `message-extension-v2` (TS), `foundry-agent-to-m365`, `teams-agent-with-data-custom-api-v2`. (The non-`.tpl` `azure.bicep` variants — typically the templates that don't use Mustache LLM-provider conditionals.)

```bicep
output AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
```

#### Variant B: `BOT_AZURE_APP_SERVICE_RESOURCE_ID` + no `BOT_ENDPOINT`

**Used by:** `weather-agent`, `custom-copilot-basic`, `custom-copilot-rag-*` (3 variants), `basic-custom-engine-agent`, `declarative-agent-with-action-from-scratch-oauth`. (The Mustache `.tpl` variants with LLM-provider conditionals.)

```bicep
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
// no BOT_ENDPOINT
```

**Implication for v4.** Both variants must be supported by deploy drivers (`azureAppService/zipDeploy` reads either `BOT_AZURE_APP_SERVICE_RESOURCE_ID` or `AZURE_APP_SERVICE_RESOURCE_ID` depending on which template the user scaffolded). v4 may **converge on one schema for new templates** but cannot break either existing schema for users with existing projects.

### Resources (both variants)

**Representative file:** [`templates/vsc/ts/weather-agent/infra/azure.bicep.tpl`](../../../templates/vsc/ts/weather-agent/infra/azure.bicep.tpl) (Variant B); [`templates/vsc/ts/default-bot/infra/azure.bicep`](../../../templates/vsc/ts/default-bot/infra/azure.bicep) (Variant A).

| Type | Symbolic name | Key properties |
|------|---------------|----------------|
| `Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31` | `identity` | location-bound |
| `Microsoft.Web/serverfarms@2021-02-01` | `serverfarm` | `kind: "app"`, SKU parameterised |
| `Microsoft.Web/sites@2021-02-01` | `webApp` | `kind: "app"`, `httpsOnly: true`, `alwaysOn: true`, `WEBSITE_RUN_FROM_PACKAGE: "1"` |
| `Microsoft.BotService/botServices@2021-03-01` | `botService` | via `botRegistration` submodule, `location: "global"`, `kind: "azurebot"`, `msaAppType: "UserAssignedMSI"` |
| `Microsoft.BotService/botServices/channels@2021-03-01` | `botServiceMsTeamsChannel` | via `botRegistration` submodule, `MsTeamsChannel` |

### Params

| Name | Type | Notes |
|------|------|-------|
| `resourceBaseName` | string | 4–20 chars; default for all resource names |
| `webAppSKU` | string | App Service Plan SKU (e.g. `B1`) |
| `botDisplayName` | string | max 42 chars |
| `openAIKey` | secure | conditional `{{#useOpenAI}}` (Variant B only) |
| `azureOpenAIKey` | secure | conditional `{{#useAzureOpenAI}}` (Variant B only) |
| `azureOpenAIEndpoint` | string | conditional `{{#useAzureOpenAI}}` (Variant B only) |
| `azureOpenAIDeploymentName` | string | conditional `{{#useAzureOpenAI}}` (Variant B only) |
| `azureSearchKey` | secure | RAG variants only |
| `azureSearchEndpoint` | string | RAG variants only |
| `location` | string | default `resourceGroup().location` |

### Sub-modules

- `./botRegistration/azurebot.bicep` — registers bot with Bot Framework. Inputs: `resourceBaseName`, `identityClientId`, `identityResourceId`, `identityTenantId`, `botAppDomain`, `botDisplayName`.

---

## Archetype: Bot on App Service with SQL

**Used by:** 1 template — `teams-collaborator-agent`.

**Representative file:** [`templates/vsc/ts/teams-collaborator-agent/infra/azure.bicep`](../../../templates/vsc/ts/teams-collaborator-agent/infra/azure.bicep).

### Resources

| Type | Symbolic name | Key properties |
|------|---------------|----------------|
| `Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31` | `identity` | — |
| `Microsoft.Sql/servers@2023-05-01-preview` | `sqlServer` | `version: "12.0"`, admin login/password parameterised |
| `Microsoft.Sql/servers/databases@2023-05-01-preview` | `sqlDatabase` | sku `Basic`, capacity 5, max 2 GB |
| `Microsoft.Sql/servers/firewallRules@2023-05-01-preview` | `sqlFirewallRule` | `AllowAllWindowsAzureIps` (0.0.0.0–0.0.0.0) |
| `Microsoft.Web/serverfarms@2021-02-01` | `serverfarm` | — |
| `Microsoft.Web/sites@2021-02-01` | `webApp` | with SQL connection string in `appSettings` |
| `Microsoft.BotService/botServices@2021-03-01` | `botService` | via submodule |
| `Microsoft.BotService/botServices/channels@2021-03-01` | `botServiceMsTeamsChannel` | via submodule |

### Params

| Name | Type | Notes |
|------|------|-------|
| `resourceBaseName` | string | 4–20 chars |
| `webAppSKU` | string | — |
| `botDisplayName` | string | max 42 chars |
| `AOAI_ENDPOINT` | string | — |
| `AOAI_API_KEY` | string | — |
| `AOAI_MODEL` | string | — |
| `sqlAdminLogin` | secure | default `"sqladmin"` |
| `sqlAdminPassword` | secure | — |
| `sqlServerName` | string | default `${resourceBaseName}-sqlserver` |
| `sqlDatabaseName` | string | default `${resourceBaseName}-db` |
| `location` | string | default `resourceGroup().location` |

### Outputs

| Name | Source |
|------|--------|
| `BOT_AZURE_APP_SERVICE_RESOURCE_ID` | `webApp.id` |
| `BOT_DOMAIN` | `webApp.properties.defaultHostName` |
| `BOT_ENDPOINT` | `https://${webApp.properties.defaultHostName}` |
| `BOT_ID` | `identity.properties.clientId` |
| `BOT_TENANT_ID` | `identity.properties.tenantId` |
| `SQL_SERVER_FQDN` | `sqlServer.properties.fullyQualifiedDomainName` |
| `SQL_DATABASE_NAME` | `sqlDatabaseName` |

### Sub-modules

- `./botRegistration/azurebot.bicep`

---

## Archetype: Serverless API on Azure Functions

**Used by:** 2 templates — `declarative-agent-with-action-from-scratch`, `declarative-agent-with-action-from-scratch-bearer`.

**Representative file:** [`templates/vsc/ts/declarative-agent-with-action-from-scratch/infra/azure.bicep`](../../../templates/vsc/ts/declarative-agent-with-action-from-scratch/infra/azure.bicep).

### Resources

| Type | Symbolic name | Key properties |
|------|---------------|----------------|
| `Microsoft.Web/serverfarms@2021-02-01` | `serverfarms` | no SKU default (template comment points users at SKU tutorial) |
| `Microsoft.Web/sites@2021-02-01` | `functionApp` | `kind: "functionapp"`, `FUNCTIONS_EXTENSION_VERSION: "~4"`, `FUNCTIONS_WORKER_RUNTIME: "node"`, `WEBSITE_RUN_FROM_PACKAGE: "1"` |

### Params

| Name | Type | Notes |
|------|------|-------|
| `resourceBaseName` | string | 4–20 chars |
| `functionAppSKU` | string | no default |
| `serverfarmsName` | string | default `resourceBaseName` |
| `functionAppName` | string | default `resourceBaseName` |
| `location` | string | default `resourceGroup().location` |

### Outputs

| Name | Source |
|------|--------|
| `API_FUNCTION_ENDPOINT` | `https://${functionApp.properties.defaultHostName}` |
| `API_FUNCTION_RESOURCE_ID` | `functionApp.id` |
| `OPENAPI_SERVER_URL` | `https://${functionApp.properties.defaultHostName}` |

### Sub-modules

None.

---

## Archetype: Graph Connector on Azure Functions

**Used by:** 1 template — `graph-connector`.

**Representative file:** [`templates/vsc/ts/graph-connector/infra/azure.bicep`](../../../templates/vsc/ts/graph-connector/infra/azure.bicep).

The heaviest archetype. Includes Key Vault, Log Analytics, App Insights, and role assignments.

### Resources

| Type | Symbolic name | Key properties |
|------|---------------|----------------|
| `Microsoft.Web/serverfarms@2021-02-01` | `appServicePlan` | — |
| `Microsoft.OperationalInsights/workspaces@2023-09-01` | `logAnalytics` | sku `PerGB2018`, retention 30 days |
| `Microsoft.Storage/storageAccounts@2022-05-01` | `storageAccount` | `kind: "Storage"`, `supportsHttpsTrafficOnly: true`, `defaultToOAuthAuthentication: true` |
| `Microsoft.Web/sites@2021-02-01` | `functionApp` | `kind: "functionapp"`, SystemAssigned identity, App Insights instrumentation key, AzureWebJobsStorage connection string |
| `Microsoft.Insights/components@2020-02-02` | `appInsights` | `Application_Type: "web"`, `WorkspaceResourceId: logAnalytics.id` |
| `Microsoft.Web/sites/siteextensions@2020-06-01` | `appServiceSiteExtension` | `Microsoft.ApplicationInsights.AzureWebSites` |
| `Microsoft.Web/sites/config@2020-06-01` | `appServiceAppSettings` | logs config (applicationLogs, httpLogs, failedRequestsTracing) |
| `Microsoft.KeyVault/vaults@2021-06-01-preview` | `keyVault` | sku family `A`, name `standard`, RBAC enabled |
| `Microsoft.KeyVault/vaults/secrets` | `storageNameSecret`, `connectorReposAccessTokenSecret` | nested secrets |
| `Microsoft.Authorization/roleAssignments@2020-04-01-preview` | `keyVaultFunctionAppPermissions` | Key Vault Secrets User role for the Function App |

### Params

| Name | Type | Notes |
|------|------|-------|
| `resourceBaseName` | string | 4–20 chars |
| `storageAccountType` | string | default `Standard_LRS` |
| `functionAppSKU` | string | — |
| `clientId` | secure | — |
| `clientSecret` | secure | — |
| `tenantId` | string | — |
| `connectorId` | string | — |
| `connectorName` | string | — |
| `connectorDescription` | string | — |
| `connectorRepos` | string | — |
| `connectorReposAccessToken` | secure | default `""` |
| `teamsfxEnv` | string | — |
| `location` | string | default `resourceGroup().location` |

### Outputs

| Name | Source |
|------|--------|
| `AZURE_FUNCTION_RESOURCE_ID` | `functionApp.id` |

### Sub-modules

None — KeyVault secrets are nested resources inside the same file.

---

## Archetype: Static Web App (Office Add-in)

**Used by:** 6 templates — `office-addin-config`, `office-addin-excel-cfshortcut`, `office-addin-outlook-taskpane`, `office-addin-wxpo-taskpane`, `office-addin` (common), `declarative-agent-meta-os-new-project`.

**Representative file:** [`templates/vsc/ts/office-addin-config/infra/azure.bicep`](../../../templates/vsc/ts/office-addin-config/infra/azure.bicep).

### Resources

| Type | Symbolic name | Key properties |
|------|---------------|----------------|
| `Microsoft.Web/staticSites@2022-09-01` | `swa` | `location: "centralus"` (hardcoded), sku `{name, tier}` parameterised |

### Params

| Name | Type | Notes |
|------|------|-------|
| `resourceBaseName` | string | 4–20 chars |
| `staticWebAppSku` | string | typically `"Free"` |
| `staticWebAppName` | string | default `resourceBaseName` |

### Outputs

| Name | Source |
|------|--------|
| `AZURE_STATIC_WEB_APPS_RESOURCE_ID` | `swa.id` |
| `ADDIN_DOMAIN` | `swa.properties.defaultHostname` |
| `ADDIN_ENDPOINT` | `https://${siteDomain}` |

### Sub-modules

None.

---

## Archetype: Tab on App Service

**Used by:** 1 template — `basic-tab` (TS + Python variants).

**Representative file:** [`templates/vsc/ts/basic-tab/infra/azure.bicep`](../../../templates/vsc/ts/basic-tab/infra/azure.bicep).

### Resources

| Type | Symbolic name | Key properties |
|------|---------------|----------------|
| `Microsoft.Web/serverfarms@2021-02-01` | `serverfarm` | `kind: "app"` |
| `Microsoft.Web/sites@2021-02-01` | `webApp` | `kind: "app"`, **NO identity**, `WEBSITE_RUN_FROM_PACKAGE: "1"` |

### Params

| Name | Type | Notes |
|------|------|-------|
| `resourceBaseName` | string | 4–20 chars |
| `webAppSku` | string | — |
| `serverfarmsName` | string | default `resourceBaseName` |
| `webAppName` | string | default `resourceBaseName` |
| `location` | string | default `resourceGroup().location` |

### Outputs

| Name | Source |
|------|--------|
| `AZURE_APP_SERVICE_RESOURCE_ID` | `webApp.id` |
| `TAB_DOMAIN` | `webApp.properties.defaultHostName` |
| `TAB_ENDPOINT` | `https://${webApp.properties.defaultHostName}` |

### Sub-modules

None.

---

## Cross-archetype observations

- **Bot output schemas do NOT converge.** The Bot/AppService archetype has two real variants (see above) that differ in the resource-ID name (`AZURE_APP_SERVICE_RESOURCE_ID` vs `BOT_AZURE_APP_SERVICE_RESOURCE_ID`) and in whether `BOT_ENDPOINT` is emitted. v4 envMap design must accept both.
- The Bot/SQL archetype uses the Variant B naming (`BOT_AZURE_APP_SERVICE_RESOURCE_ID`) plus `SQL_SERVER_FQDN` and `SQL_DATABASE_NAME`.
- API versions cluster around 2021-02 / 2021-03 / 2021-06. Several years stale.
- Identity is **only** on bot archetypes. Tab and Functions/API don't get one by default.
- Static Web App archetype hardcodes `centralus`. Functions archetype has no SKU default.
- Graph Connector is the only archetype that uses Key Vault and managed identity end-to-end. v4 should consider promoting that pattern across other archetypes — but design the promotion from Azure best-practice docs, not from this single example.
