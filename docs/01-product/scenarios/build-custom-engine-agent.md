# Scenario — Build a Custom Engine Agent (CEA)

**Persona:** P1 / P3 (TS / JS / Python developer)
**Outcome:** A code-first agent running on Microsoft Agents SDK, hosted on Azure, sideloaded into Microsoft 365.
**Surface:** VS Code · CLI

## 1 — Scaffold

| Template ID | Description | Languages |
|-------------|-------------|-----------|
| `cea/basic` | Echo-style starting point | TS · JS · Python |
| `cea/weather` | Function-calling weather agent | TS · JS (Python folder ships but is not in `features.json`; treat as scaffold-only for Python) |
| `cea/teams-collaborator` | Multi-step collaboration sample | TS |

Question tree (TS path): `projectNameQuestion → languageQuestion → folderQuestion`. See [`commonQuestions.ts`](../../../packages/core-next/src/questions/commonQuestions.ts).

## 2 — Run / debug locally

`m365agents.local.yml` runs:

- `botAadApp/create` — bot Entra ID app + password
- `arm/deploy` (subset, often skipped for pure-local) or local registration
- `script` — start dev tunnel
- `script` — start the bot process (`npm run dev` / `func start` / `python app.py`)

User then sideloads via M365 Agents Playground or Teams.

## 3 — Provision

| Action | Driver |
|--------|--------|
| Bot AAD app | `botAadApp/create` |
| Bot channel | `botFramework/create` (ARM-backed) |
| Azure infra (App Service or Functions) | `arm/deploy` |
| Teams app | `teamsApp/create`, `teamsApp/configure` |
| Manifest validation | `teamsApp/validateManifest` |
| Package | `teamsApp/zipAppPackage` |

Prerequisites: `ensureM365Auth` + `ensureAzureAuth` + `ensureSubscription` + `ensureResourceGroup` + `ensureResourceSuffix`. Consent gate: `confirmProvision`.

## 4 — Deploy

| Stack | Drivers |
|-------|---------|
| TS / JS on App Service | `cli/runNpmCommand` (build) → `azureAppService/zipDeploy` |
| TS / JS on Functions | `cli/runNpmCommand` (build) → `azureFunctions/zipDeploy` |
| Python on Functions | `azureFunctions/zipDeploy` (no npm; project ships ready) |
| C# on App Service | `cli/runDotnetCommand` (publish) → `azureAppService/zipDeploy` |

## 5 — Publish

`teamsApp/publishAppPackage` posts the app to the org catalog. Optional `teamsApp/extendToM365` for sideload.

## Files produced (TS, App Service)

```
src/                       # bot code (Microsoft Agents SDK)
appPackage/manifest.json
infra/
  azure.bicep
  azure.parameters.json
m365agents.yml
m365agents.local.yml
env/.env.dev
.vscode/launch.json
```

## Where this is tested

- v3: `packages/fx-core/tests/integration/` — IDs `custom-copilot-basic`, `custom-copilot-rag-customize`, etc.
- v4 E2E: `cea/basic` (TS/JS/Python) and `cea/weather` (TS/JS) verified via cli-next.
