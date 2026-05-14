# Capabilities matrix

> For the **deep template-by-template inventory** (descriptions, optional features, integration surface, gaps informing v4), see [v3-feature-inventory.md](v3-feature-inventory.md). This page is the cross-engine summary table; that page is the full extraction.
>
> **Source of truth:** [`.dev/features.json`](../../.dev/features.json), validated by [`packages/fx-core/tests/integration/featureRegistry.ts`](../../packages/fx-core/tests/integration/featureRegistry.ts). When this page diverges from those, the JSON wins. Run `npm run test:integration` in `packages/fx-core` to regenerate the coverage report.

## Reading this matrix

- **Languages** — which languages have a working scaffold today.
- **v3** — testable through the existing fx-core engine.
- **v4** — verified end-to-end through `core-next` + `cli-next` (currently 9 templates verified).
- **ADO suite** — link to manual test plan for traceability.

## Capability matrix (current)

| Category | Template | Languages | v3 | v4 E2E | ADO |
|----------|----------|-----------|----|--------|-----|
| Bot | Echo Bot (`default-bot`) | TS · JS · Python | ✓ | ✓ | 24569101 |
| Tab | Basic Tab (`basic-tab`) | TS | ✓ | ✓ | 24569106 |
| Custom Engine Agent | Basic CEA | TS · JS · Python | ✓ | ✓ | 34834051 |
| Custom Engine Agent | Weather (function-calling) | TS · JS · Python | ✓ | ✓ | 34648283 |
| Custom Engine Agent | Teams Collaborator | TS | ✓ | — | 35527236 |
| AI Agent | AI Chat Bot (`custom-copilot-basic`) | TS · JS · Python | ✓ | ✓ | 27042287 |
| AI Agent | AI Agent + AI Search (RAG) | TS · JS · Python | ✓ | — | 27689412 |
| AI Agent | AI Agent + Custom Data (RAG) | TS · JS · Python | ✓ | — | 27689419 |
| Declarative Agent | DA Basic | common | ✓ | ✓ | 27971458 |
| Declarative Agent | DA + API Plugin (No Auth) | TS · JS | ✓ | — | 27971458 |
| Declarative Agent | DA + API Plugin (OAuth) | TS · JS | ✓ | — | 27971458 |
| Declarative Agent | DA + API Plugin (Bearer) | TS · JS | ✓ | — | 27971458 |
| Connector | Graph Connector | TS | ✓ | ✓ | 32019603 |
| Messaging Extension | ME v2 | TS · Python | ✓ | — | 34869329 |
| AI Agent *(tracked)* | RAG Custom API | — | — | — | 27588348 |
| AI Agent *(tracked)* | Foundry Proxy Agent | C# | — | — | 36750068 |

*Italic rows are tracked for ADO traceability only; no local scaffold today.*

## Lifecycle coverage by template (v4 `core-next`)

The v4 engine has **22 built-in drivers** (counted directly from [`packages/core-next/src/drivers/builtin/`](../../packages/core-next/src/drivers/builtin/)) covering most templates. The headline gap is the `typeSpec/compile` driver for `declarative-agent-typespec`. See [_v3-reference/infra/template-inventory.md](../_v3-reference/infra/template-inventory.md) for the per-template archetype → driver dependency.

| Stage | Drivers involved (top hits) |
|-------|----------------------------|
| Provision | `aadApp/create`, `aadApp/update`, `botAadApp/create`, `botFramework/create`, `arm/deploy`, `teamsApp/create`, `teamsApp/configure`, `teamsApp/validateManifest`, `teamsApp/zipAppPackage`, `oauth/register`, `apiKey/register` |
| Deploy | `cli/runNpmCommand`, `cli/runDotnetCommand`, `azureAppService/zipDeploy`, `azureFunctions/zipDeploy`, `script` |
| Publish | `teamsApp/publishAppPackage` (Graph `/beta/appCatalogs/teamsApps`), `teamsApp/extendToM365` |

See [05-engineering/cross-cutting/driver-system.md](../05-engineering/cross-cutting/driver-system.md) for the full driver catalogue.

## How to add a capability

See [07-contributing/adding-a-template.md](../07-contributing/adding-a-template.md) and [features.instructions.md](../../.github/instructions/features.instructions.md).
