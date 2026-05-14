# Scenario — Build a Graph Connector

**Persona:** P1
**Outcome:** A connector that ingests external content into Microsoft Graph for Copilot retrieval.
**Surface:** VS Code · CLI

## 1 — Scaffold

| Template ID | Languages |
|-------------|-----------|
| `graph-connector` | TS |

Question tree adds graph-connector-specific questions (see [`packages/core-next/src/templates/descriptors/connector.ts`](../../../packages/core-next/src/templates/descriptors/connector.ts)).

## Lifecycle highlights

| Stage | Notable drivers |
|-------|----------------|
| Provision | `aadApp/create`, `aadApp/update` (Graph permissions), `arm/deploy` (Functions), `teamsApp/create`/`configure` |
| Deploy | `cli/runNpmCommand` → `azureFunctions/zipDeploy` |
| Publish | `teamsApp/publishAppPackage` |

## Where this is tested

- v4 E2E: `connector/graph-connector`
- ADO suite: 32019603
