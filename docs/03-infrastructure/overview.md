# Infrastructure overview

The toolkit ships **opinionated infrastructure as Bicep** in each project's `infra/` folder. Provisioning runs `arm/deploy` against a per-environment resource group whose name follows the pattern:

```
rg-{safeProjectName}{resourceSuffix}-{envName}
```

The `resourceSuffix` is a 6-char random string generated once per project and reused across environments. See `lifecycle/prerequisites.ts → ensureResourceSuffix`.

## Topology archetypes

| Archetype | When | Backbone resources |
|----------|------|-------------------|
| **No-Azure DA** | Pure Declarative Agent (no API plugin) | None — only Teams app + tenant catalog |
| **Bot on App Service** | Bot, CEA, ME with persistent connections | App Service plan + Web App + Bot Channel + Entra app + (App Insights) |
| **Bot on Functions** | Lightweight bot, event-triggered | Function App + Bot Channel + Entra app + Storage |
| **Tab on Static Web Apps** | Pure-frontend tab | Static Web App + Entra app |
| **Graph Connector** | Connector ingestion | Function App + Entra app (with Graph permissions) + Storage |
| **API plugin backend** | DA + custom API | Function App or Web App + Entra app (for OAuth flow) |

## Naming conventions

| Resource | Pattern | Source |
|----------|---------|--------|
| Resource group | `rg-{safeName}{suffix}-{env}` | `ensureResourceGroup` default |
| Web/Function App | `{safeName}{suffix}-{env}` | Bicep parameters |
| App Service plan | `{safeName}{suffix}-asp-{env}` | Bicep parameters |
| Bot service | `{safeName}{suffix}-bot-{env}` | Bicep parameters |
| Storage account | `st{safeName}{suffix}{env}` (no dashes) | Bicep parameters |

`safeName` is the URL-safe projection of `appName`, lowercased and stripped of separators where Azure resource naming forbids them.

## Ownership

The Bicep is **owned by the user once scaffolded**. The toolkit never edits the user's `infra/` files after creation. New capabilities ship by updating the *templates*; existing projects stay on whatever Bicep they were scaffolded with.
