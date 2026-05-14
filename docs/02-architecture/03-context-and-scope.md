# 3 — Context and scope

## System context (C1)

```
        ┌──────────────────────────────────────────────────────────┐
        │                       Developer                          │
        │  (VS Code · Visual Studio · CLI · CI runner)             │
        └────────────────┬─────────────────────────────────────────┘
                         │ uses
                         ▼
   ┌────────────────────────────────────────────────────────────────┐
   │              Microsoft 365 Agents Toolkit                      │
   │   (this repo — runs on the developer's machine / CI)           │
   └────┬───────────┬───────────┬───────────┬───────────┬───────────┘
        │           │           │           │           │
        ▼           ▼           ▼           ▼           ▼
   ┌────────┐  ┌────────┐  ┌────────┐  ┌────────┐  ┌──────────┐
   │ Entra  │  │ Azure  │  │  Teams │  │   MS   │  │   Bot    │
   │   ID   │  │   ARM  │  │  Dev   │  │  Graph │  │ Framework│
   │ (MSAL) │  │ + Kudu │  │ Portal │  │  /beta │  │   (ARM)  │
   └────────┘  └────────┘  └────────┘  └────────┘  └──────────┘
                                │
                                ▼
                          ┌──────────┐
                          │   M365   │
                          │ Package  │
                          │ Service  │
                          └──────────┘
```

The toolkit itself **runs entirely on the developer's machine** (or CI runner). It never hosts user code.

## External interfaces

| External system | Contract | Implementation |
|-----------------|----------|----------------|
| Entra ID (MSAL) | `@azure/msal-node` interactive auth code flow + silent refresh | [`cli-next/src/auth/`](../../packages/cli-next/src/auth/) (v4) · [`cli/src/commonlib/`](../../packages/cli/src/commonlib/) (v3) |
| Azure ARM | Subscription / Resource Group APIs · ARM/Bicep deployment · Kudu zip deploy | [`core-next/src/clients/azure/`](../../packages/core-next/src/clients/azure/) |
| Teams Developer Portal | App CRUD, manifest validation, package validation, OAuth & API key registration | [`core-next/src/clients/teamsDevPortal/`](../../packages/core-next/src/clients/teamsDevPortal/) |
| Microsoft Graph | Entra ID app reg + password creds; `/beta/appCatalogs/teamsApps` for publish | [`core-next/src/clients/graphApi/`](../../packages/core-next/src/clients/graphApi/) |
| Bot Framework | Channel registration via ARM | [`drivers/builtin/botFramework/create.ts`](../../packages/core-next/src/drivers/builtin/botFramework/create.ts) |
| M365 PackageService | V1 (classic) + V2 (declarative agent) sideloading | [`core-next/src/clients/m365/`](../../packages/core-next/src/clients/m365/) |
| Microsoft Foundry | Foundry Agent template (preview) | [`templates/descriptors/foundry.ts`](../../packages/core-next/src/templates/descriptors/foundry.ts) |
| Spec providers | `@apidevtools/swagger-parser`, `swagger2openapi` | [`core-next/src/specParser/`](../../packages/core-next/src/specParser/) |

## What is *not* in scope for the toolkit itself

- Hosting the developer's deployed agent (that's Azure)
- Storing secrets persistently outside `~/.fx/account/` token cache
- Replacing `az` / `azd` / `func` for general Azure development
- Maintaining the upstream manifest schemas (those live in the M365 platform team)
