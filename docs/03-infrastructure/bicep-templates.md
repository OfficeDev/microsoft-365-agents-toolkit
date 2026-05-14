# Bicep templates

What lives under a scaffolded project's `infra/` folder.

## Standard layout

```
infra/
  azure.bicep                # main template; entry of arm/deploy
  azure.parameters.json      # parameter file with ${{VAR}} placeholders resolved at provision time
  botRegistration/           # bot channel module (if bot)
  appInsights/               # app insights module
```

## `azure.bicep` responsibilities

- Declares all resources for the topology archetype.
- Surfaces output values that flow back into `envMap` via `arm/deploy` driver outputs.
- Uses `param` blocks for everything controlled by the lifecycle (location, name suffix, env name).

## Output → envMap conventions

| Output name | Consumed by | Typical value |
|-------------|------------|---------------|
| `BOT_AZURE_APP_SERVICE_RESOURCE_ID` | `azureAppService/zipDeploy` | `/subscriptions/.../resourceGroups/.../providers/Microsoft.Web/sites/{name}` |
| `AZURE_APP_SERVICE_RESOURCE_ID` | `azureAppService/zipDeploy` (alternate name in some bot templates) | same shape |
| `BOT_DOMAIN` | Manifest `validDomains` | `{name}.azurewebsites.net` |
| `BOT_ENDPOINT` | Manifest bot URL | `https://{name}.azurewebsites.net/api/messages` |
| `FUNCTION_APP_RESOURCE_ID` | `azureFunctions/zipDeploy` | similar |

> **v3 inconsistency.** The bot/App-Service archetype ships in two output-name variants — some templates emit `AZURE_APP_SERVICE_RESOURCE_ID` (and `BOT_ENDPOINT`), others emit `BOT_AZURE_APP_SERVICE_RESOURCE_ID` (and omit `BOT_ENDPOINT`). Both work because `azureAppService/zipDeploy` accepts either. See [`_v3-reference/infra/archetypes.md §Bot on App Service with Identity`](../_v3-reference/infra/archetypes.md#archetype-bot-on-app-service-with-identity) for the per-template breakdown.

The naming convention `{COMPONENT}_{PROPERTY}` uppercase-snake mirrors the env-var style consumed by deploy drivers.

## Placeholder resolution

`azure.parameters.json` uses `${{ENV_VAR}}` placeholders (e.g. `${{AZURE_RESOURCE_GROUP_NAME}}`, `${{RESOURCE_SUFFIX}}`). The lifecycle executor injects envMap entries into `process.env` immediately before the `arm/deploy` driver runs, then cleans up.

See [`packages/core-next/src/lifecycle/executor.ts`](../../packages/core-next/src/lifecycle/executor.ts).

## Customising

Users own their `infra/` post-scaffold. Common edits:

- SKU upsizing — change `sku.name` in the App Service plan.
- VNet integration — add `Microsoft.Web/sites/networkConfig` and a VNet module.
- Managed identity — add `identity: { type: 'SystemAssigned' }` and role assignments. See [networking-and-identity.md](networking-and-identity.md).
- Key Vault references — switch `appSettings` values to `@Microsoft.KeyVault(...)` references.
