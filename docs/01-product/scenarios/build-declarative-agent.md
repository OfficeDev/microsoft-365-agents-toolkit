# Scenario — Build a Declarative Agent

**Persona:** P1 (TypeScript / JavaScript developer)
**Outcome:** A Declarative Agent published to the user's tenant catalog and reachable from Microsoft 365 Copilot.
**Surface:** VS Code · CLI

## 1 — Scaffold

User picks **Declarative Agent** in the create-project flow.

| Question | Typical answer | Source |
|----------|----------------|--------|
| Template | `da/basic` (no backend) · `da/api-plugin-{no-auth,oauth,bearer,entra-sso}` (with backend) · `da/mcp-{remote,local}` (MCP-backed) · `da/graph-connector` · `da/typespec` · `da/existing-action` · `da/metaos[-upgrade]` | [`packages/core-next/src/templates/descriptors/declarativeAgent.ts`](../../../packages/core-next/src/templates/descriptors/declarativeAgent.ts) |
| App name | `MyAgent` | `commonQuestions.projectNameQuestion()` |
| Language | TypeScript / JavaScript / *common* | depends on template |
| Folder | OS picker | platform UI |

The DA Basic template uses `language = common` (no source code, just manifest + DA JSON).

## 2 — Run / debug locally

The DA flow does not have a local-debug runtime in the traditional sense; the app is sideloaded into M365 to be tried in Copilot. The **`m365agents.local.yml`** lifecycle runs:

- `teamsApp/zipAppPackage`
- `teamsApp/extendToM365` — sideloads the V2 declarative agent via the M365 PackageService
- `teamsApp/validateManifest`

## 3 — Provision

Provision runs the actions in `m365agents.yml`. For a no-Azure DA Basic template:

| Action | Driver | Notes |
|--------|--------|-------|
| Create Teams app | `teamsApp/create` | Idempotent via `existingTeamsAppId` |
| Configure | `teamsApp/configure` | Updates TDP record |
| Validate manifest | `teamsApp/validateManifest` | Schema 2.4 |
| Zip app package | `teamsApp/zipAppPackage` | Bundles manifest + icons |

API-plugin variants additionally register OAuth or API key:

- `oauth/register` — for OAuth-protected APIs
- `apiKey/register` — for bearer/API-key APIs

Prerequisites enforced by [`lifecycle/prerequisites.ts`](../../../packages/core-next/src/lifecycle/prerequisites.ts): `ensureM365Auth` (no Azure for pure-DA).

## 4 — Deploy

Pure DA templates have no code to deploy — `deployOp` is effectively a no-op.

API-plugin templates with a backend deploy that backend (typically `azureFunctions/zipDeploy` or `azureAppService/zipDeploy`).

## 5 — Publish

`teamsApp/publishAppPackage` posts the package to the Microsoft Graph endpoint `/beta/appCatalogs/teamsApps`, registering the agent in the org catalog.

Optional: `teamsApp/extendToM365` (V2 codepath) sideloads the DA so the user can test it before tenant publishing.

## Files produced (DA Basic)

```
appPackage/
  manifest.json                  # Teams manifest 2.4 (declarativeCopilots[])
  declarativeAgent.json          # DA definition
  icons/color.png, icons/outline.png
m365agents.yml                   # provision/publish actions
m365agents.local.yml             # sideload actions
env/.env.dev
```

## Drivers in play

`teamsApp/create`, `teamsApp/configure`, `teamsApp/update`, `teamsApp/validateManifest`, `teamsApp/zipAppPackage`, `teamsApp/publishAppPackage`, `teamsApp/extendToM365` — and for plugin variants: `oauth/register`, `apiKey/register`.

## DA-specific operations (v4)

The [`declarativeAgent/`](../../../packages/core-next/src/declarativeAgent/) module exposes operations users invoke after scaffolding to extend a DA:

- `addKnowledgeOp` — web search · OneDrive/SharePoint · Graph connector · embedded knowledge
- `addExistingPluginOp` — attach an existing OpenAPI plugin to the DA
- `addMCPActionOp` — add an MCP-server-backed action
- `setSensitivityLabelOp` — capability annotation
- `setConversationStartersOp`

## Where this is tested

- v3 integration: `packages/fx-core/tests/integration/` (feature ID `declarative-agent-basic` and plugin variants)
- v4 unit: `packages/core-next/tests/unit/templates/descriptors.test.ts`
- v4 integration: `packages/core-next/tests/integration/daPackaging.test.ts`, `daOperations.test.ts`
- v4 E2E: cli-next E2E for `da/basic`
