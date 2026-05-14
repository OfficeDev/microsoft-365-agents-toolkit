# Networking and identity

## Identity model

Every templated app produces **at least** one Entra ID app registration:

| Identity | Purpose | Driver |
|----------|---------|--------|
| Bot Entra ID app | Bot Framework identity (single- or multi-tenant) | `botAadApp/create` |
| Backend AAD app | Tab SSO, API plugin OAuth, Graph access | `aadApp/create` + `aadApp/update` |

Driver outputs:

| Output | Consumed by |
|--------|------------|
| `BOT_ID` | `manifest.json` (`bots[].botId`), Bicep params |
| `BOT_PASSWORD` | Web/Function App `appSettings`, secret env file |
| `AAD_APP_CLIENT_ID` | Manifest, OAuth flows |
| `AAD_APP_CLIENT_SECRET` | Backend secret env file |
| `AAD_APP_OAUTH_AUTHORITY` | Auth flows |

## AAD manifest template

`aadApp/update` resolves `${{VAR}}` placeholders in the AAD manifest template before parsing. This lets templates reference `${{BOT_DOMAIN}}` for redirect URIs without knowing the value at scaffold time.

See [`packages/core-next/src/drivers/builtin/aadApp/update.ts`](../../packages/core-next/src/drivers/builtin/aadApp/update.ts).

## Permissions surfaced

Templates declare the Graph permissions they need in the AAD manifest's `requiredResourceAccess`. Common patterns:

| Capability | Permissions (delegated unless noted) |
|-----------|--------------------------------------|
| Tab SSO | `User.Read` |
| Graph Connector | `ExternalConnection.ReadWrite.OwnedBy` (application) + `ExternalItem.ReadWrite.OwnedBy` (application) |
| DA + API plugin (OAuth) | Whatever the API requires; templates make this configurable |

## Networking

Templates default to **public endpoints** with no VNet integration. This is the right default for a getting-started experience. Production hardening is a user customisation:

- VNet integration on Web/Function Apps
- Private endpoints for Storage and downstream services
- IP restrictions on the App Service
- Front Door / API Management for ingress

The toolkit does **not** prescribe a hardened-prod topology — users compose those patterns themselves on top of the scaffolded Bicep.

## Managed identity

Not enabled by default in templates. To opt in:

```bicep
resource site 'Microsoft.Web/sites@2024-04-01' = {
  ...
  identity: { type: 'SystemAssigned' }
}
```

…then assign roles via additional Bicep modules. The toolkit's `azure-rbac` skill (`.agents/skills/...` / runtime-loaded skills) gives recommended least-privilege roles.
