# Lifecycle YAML — `m365agents.yml`

The declarative description of provision, deploy, and publish for a project. Parsed by [`lifecycle/parser.ts`](../../../packages/core-next/src/lifecycle/parser.ts) into a `RawProjectModel`, resolved against the `DriverRegistry`, and executed by [`lifecycle/executor.ts`](../../../packages/core-next/src/lifecycle/executor.ts).

## Shape

```yaml
version: 1.1.0
provision:
  - uses: aadApp/create
    with:
      name: my-bot-aad
      generateClientSecret: true
    writeToEnvironmentFile:
      clientId: AAD_APP_CLIENT_ID
      clientSecret: SECRET_AAD_APP_CLIENT_SECRET
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID

  - uses: arm/deploy
    with:
      subscriptionId: ${{AZURE_SUBSCRIPTION_ID}}
      resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}}
      templates:
        - path: ./infra/azure.bicep
          parameters: ./infra/azure.parameters.json
          deploymentName: Create-resources

deploy:
  - uses: cli/runNpmCommand
    with:
      args: install
  - uses: cli/runNpmCommand
    with:
      args: run build --if-present
  - uses: azureAppService/zipDeploy
    with:
      artifactFolder: .
      ignoreFile: ./.webappignore
      resourceId: ${{BOT_AZURE_APP_SERVICE_RESOURCE_ID}}

publish:
  - uses: teamsApp/validateManifest
  - uses: teamsApp/zipAppPackage
  - uses: teamsApp/validateAppPackage
  - uses: teamsApp/publishAppPackage
```

## Step structure (`DriverStep`)

| Field | Required | Purpose |
|-------|----------|---------|
| `uses` | yes | Driver ID — must be registered in `DriverRegistry` |
| `with` | depends | Driver-specific input; validated by Zod schema |
| `writeToEnvironmentFile` | no | Map of driver output keys → env var names |
| `env` | no | Per-step env overrides |

## Placeholder resolution

`${{VAR_NAME}}` placeholders resolve from the **env-map** (loaded env file + outputs of preceding steps). Resolution happens in two places:

1. In YAML `with:` blocks at parse time (string interpolation).
2. In external files referenced by the driver (e.g. ARM parameter JSON, AAD manifest template) at execution time, via the `process.env` injection trick documented in [02-architecture/06-runtime-views.md](../../02-architecture/06-runtime-views.md).

## Extension points

- **New driver** → add it to `src/drivers/builtin/` and register via `registerBuiltinDrivers()`. YAML can use `uses: my-namespace/myAction` immediately.
- **Conditional steps** → not natively supported. Drivers handle internal branching (e.g. `teamsApp/create` is idempotent on `existingTeamsAppId`).

## Backward compatibility

v4 driver IDs deliberately match v3 names (`aadApp/create`, `arm/deploy`, `teamsApp/publishAppPackage`, ...) so existing `m365agents.yml` files continue to work under `TEAMSFX_V4_CORE`.

## Local lifecycle

`m365agents.local.yml` follows the same shape but contains only the actions needed for local F5 — typically tunnel setup, bot AAD creation, and the local sideload via `teamsApp/extendToM365`.
