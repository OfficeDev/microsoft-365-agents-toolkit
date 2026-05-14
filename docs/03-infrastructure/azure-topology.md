# Azure topology archetypes

## Bot / CEA on App Service

```mermaid
flowchart LR
    User --> Teams
    Teams --> BC[Bot Channel registration]
    BC --> WA[Web App on App Service]
    WA --> AI[App Insights]
    WA -.identity.-> Entra[Entra ID app]
    style WA fill:#0a84ff20
```

| Resource | Purpose |
|----------|---------|
| App Service plan | Compute tier (B1, S1, P1v3 typical) |
| Web App | Hosts the Microsoft Agents SDK process |
| App Insights | Telemetry sink |
| Bot Service | Channel registration (Teams) |
| Entra ID app | Bot identity + manifest single-tenant or multi-tenant |

## Bot / CEA on Azure Functions

Same shape, Web App replaced by Function App + Storage. Driver: `azureFunctions/zipDeploy`.

```mermaid
flowchart LR
    User --> Teams --> BC[Bot Channel]
    BC --> FA[Function App]
    FA --> Storage[Storage account]
    FA --> AI[App Insights]
    FA -.identity.-> Entra
```

## Tab on Static Web Apps

```mermaid
flowchart LR
    User --> Teams --> SWA[Static Web App]
    SWA -.SSO.-> Entra
```

## Graph Connector

```mermaid
flowchart LR
    Schedule[Timer trigger] --> FA[Function App]
    FA -- ingest --> Graph[Microsoft Graph]
    FA --> External[(External system)]
    FA -.identity.-> Entra[Entra app w/ Graph perms]
```

## DA with API plugin

```mermaid
flowchart LR
    Copilot[M365 Copilot] --> DA[DA in tenant catalog]
    DA -- API plugin --> API[Function App / Web App]
    API -.identity.-> Entra
    API -.OAuth.-> Auth[Entra OAuth]
```

## SKU defaults shipped

Templates default to inexpensive SKUs to minimise cost during first-run experimentation. Users are expected to upgrade for production via standard Bicep edits.

| Resource | Default SKU |
|----------|-------------|
| App Service plan | B1 |
| Function App plan | Y1 (Consumption) |
| Storage account | Standard_LRS |
| Static Web App | Free |
