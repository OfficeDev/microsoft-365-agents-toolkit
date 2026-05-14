# Template → infra archetype inventory (v3)

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

Source: [`templates/vsc/`](../../../templates/vsc/) (TS, JS, Python, common).

| Template | Languages | Has infra | Archetype |
|----------|-----------|-----------|-----------|
| `default-bot` | TS · JS · Python | Yes (mixed: TS/JS=`.bicep`, Python=`.bicep.tpl`) | [Bot on App Service with Identity — Variant A](archetypes.md#variant-a-azure_app_service_resource_id--with-bot_endpoint) |
| `basic-tab` | TS · Python | Yes (`.bicep`) | [Tab on App Service](archetypes.md#archetype-tab-on-app-service) |
| `basic-custom-engine-agent` | TS · JS · Python | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) |
| `weather-agent` | TS · JS · Python | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) |
| `teams-collaborator-agent` | TS | Yes (`.bicep`) | [Bot on App Service with SQL](archetypes.md#archetype-bot-on-app-service-with-sql) |
| `custom-copilot-basic` | TS · JS · Python | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) |
| `custom-copilot-rag-azure-ai-search` | TS · JS · Python | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) (+ Search params) |
| `custom-copilot-rag-custom-api` | TS · JS · Python | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) |
| `custom-copilot-rag-customize` | TS · JS · Python | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) |
| `message-extension-v2` | TS · Python | Yes (mixed: TS=`.bicep`, Python=`.bicep.tpl`) | [Bot on App Service with Identity — Variant A](archetypes.md#variant-a-azure_app_service_resource_id--with-bot_endpoint) |
| `declarative-agent-with-action-from-scratch` | TS · JS | Yes (`.bicep`) | [Serverless API on Functions](archetypes.md#archetype-serverless-api-on-azure-functions) |
| `declarative-agent-with-action-from-scratch-bearer` | TS · JS | Yes (`.bicep`) | [Serverless API on Functions](archetypes.md#archetype-serverless-api-on-azure-functions) |
| `declarative-agent-with-action-from-scratch-oauth` | TS · JS | Yes (`.tpl`) | [Bot on App Service with Identity — Variant B](archetypes.md#variant-b-bot_azure_app_service_resource_id--no-bot_endpoint) |
| `foundry-agent-to-m365` | TS | Yes (`.bicep`) | [Bot on App Service with Identity — Variant A](archetypes.md#variant-a-azure_app_service_resource_id--with-bot_endpoint) |
| `graph-connector` | TS | Yes (`.bicep`) | [Graph Connector on Functions](archetypes.md#archetype-graph-connector-on-azure-functions) |
| `office-addin-config` | TS | Yes (`.bicep`) | [Static Web App](archetypes.md#archetype-static-web-app-office-add-in) |
| `office-addin-excel-cfshortcut` | TS | Yes (`.bicep`) | [Static Web App](archetypes.md#archetype-static-web-app-office-add-in) |
| `office-addin-outlook-taskpane` | TS | Yes (`.bicep`) | [Static Web App](archetypes.md#archetype-static-web-app-office-add-in) |
| `office-addin-wxpo-taskpane` | TS | Yes (`.bicep`) | [Static Web App](archetypes.md#archetype-static-web-app-office-add-in) |
| `office-addin` (common) | common | Yes (`.bicep`) | [Static Web App](archetypes.md#archetype-static-web-app-office-add-in) |
| `declarative-agent-meta-os-new-project` | common | Yes (`.bicep`) | [Static Web App](archetypes.md#archetype-static-web-app-office-add-in) |
| `teams-agent-with-data-custom-api-v2` | Python | Yes (`.bicep`) | [Bot on App Service with Identity — Variant A](archetypes.md#variant-a-azure_app_service_resource_id--with-bot_endpoint) |
| `declarative-agent-basic` | common | No | — |
| `declarative-agent-meta-os-upgrade-project` | common | No | — |
| `declarative-agent-typespec` | common | No | — |
| `declarative-agent-with-action-from-existing-api` | common | No | — |
| `declarative-agent-with-action-from-mcp` | common | No | — |
| `message-extension-with-existing-api` | common | No | — |
| `office-xml-addin-common` | common | No | — |

## Summary

- **22 of 29 templates** ship Azure infra.
- **6 distinct archetypes** cover all of them — with the **Bot on App Service with Identity** archetype splitting into **two output-schema variants** that differ in their first output name and whether `BOT_ENDPOINT` is emitted.
- Most-used archetype: **Bot on App Service with Identity** (11 templates split across the two variants).
- Heaviest archetype: **Graph Connector on Functions** (7 resource families including Key Vault, Log Analytics, App Insights, role assignments).
- File extensions are mixed within the same template family. `default-bot` (TS+JS) and `message-extension-v2` (TS) use raw `.bicep`; their Python variants use `.bicep.tpl`. This reflects template-history rather than design intent.

For full per-archetype detail see [archetypes.md](archetypes.md).
