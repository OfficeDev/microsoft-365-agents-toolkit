# Scenario — Build a Bot

**Persona:** P1 / P3
**Outcome:** A Microsoft Agents SDK bot running locally and on Azure, available in Teams.
**Surface:** VS Code · CLI

## 1 — Scaffold

| Template ID | Languages |
|-------------|-----------|
| `default-bot` (Echo Bot) | TS · JS · Python |

## 2 — Run / debug

- Local sideload via M365 Agents Playground (preferred) or Teams.
- `m365agents.local.yml` provisions an Entra ID app for the bot, registers a channel via Bot Framework (or its emulator), and starts the bot process.

## 3 — Provision (Azure)

Same shape as [CEA](build-custom-engine-agent.md) minus AI-specific bits.

| Action | Driver |
|--------|--------|
| Bot AAD app | `botAadApp/create` |
| Bot channel | `botFramework/create` |
| App Service or Functions | `arm/deploy` |
| Teams app | `teamsApp/create`, `teamsApp/configure`, `teamsApp/validateManifest`, `teamsApp/zipAppPackage` |

## 4 — Deploy

`cli/runNpmCommand` (build) → `azureAppService/zipDeploy` *or* `azureFunctions/zipDeploy`.

## 5 — Publish

`teamsApp/publishAppPackage` for org catalog distribution.

## Where this is tested

- v3: `default-bot` integration suite
- v4 E2E: `default-bot` (TS / JS / Python)
