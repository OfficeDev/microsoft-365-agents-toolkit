# `_v3-reference/infra/` — v3 infrastructure archetypes

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

Pure extractions of the Bicep templates, resource shapes, and naming conventions that ship with v3 templates. For archival reference only.

## Pages

- [template-inventory.md](template-inventory.md) — every template → its archetype mapping
- [archetypes.md](archetypes.md) — the six recurring infra patterns (Bot/AppService, Bot/SQL, Functions/API, Functions/Connector, Static Web App, Tab/AppService) with full resource lists, params, and outputs
- [naming-conventions.md](naming-conventions.md) — `resourceBaseName`, `RESOURCE_SUFFIX`, location handling, RG conventions

## Why this is forbidden as v4 design input

v4 templates re-emit Bicep from **topology requirements** (the user's intent: "I need a bot that runs on Azure with an Entra ID identity"), not from copying the v3 `.bicep` files. Re-deriving from requirements lets v4 fix v3 issues like:

- Hardcoded `centralus` for Static Web Apps where the SKU now supports more regions.
- `Standard_LRS` storage default for connectors (no zone redundancy).
- `Microsoft.Web/sites@2021-02-01` API version (now several years out of date).
- Lack of managed identity adoption in the bot/App Service archetype.
- `WEBSITE_RUN_FROM_PACKAGE: "1"` without explicit Run-From-URL guidance.
- Missing private endpoint / VNet-integration scaffolds for production hardening.

Importing v3's exact Bicep would re-import these gaps. Use the [archetype list](archetypes.md) only to confirm "we still want a bot-on-app-service archetype" — then design the v4 Bicep from current Azure best practice.
