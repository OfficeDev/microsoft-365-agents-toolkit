# Feature registry

The single source of truth for *what the toolkit can build*.

## Files

| File | Role |
|------|------|
| [`.dev/features.json`](../../../.dev/features.json) | Machine-readable inventory. Edit this. |
| [`packages/fx-core/tests/integration/featureRegistry.ts`](../../../packages/fx-core/tests/integration/featureRegistry.ts) | Typed wrapper. Validates `templateName` against the `TemplateNames` enum at load time. |

## Entry shape

```typescript
{
  id: "default-bot",
  name: "Echo Bot",
  description: "...",
  category: "Bot",                                  // FeatureCategory
  templateName: "default-bot",                      // must match TemplateNames enum
  languages: ["typescript", "javascript", "python"],
  lifecycles: ["scaffold", "provision", "deploy", "publish"],
  projectType: "bot-type",                          // metadata only
  capabilities: ["conversational"],
  generatedFiles: ["src/index.ts", "appPackage/manifest.json", ...],
  entryPoints: ["fx-core", "cli", "server", "vscode"],
  adoSuiteId: 24569101,                             // ADO traceability
  adoTestCaseCount: 12,
}
```

## Helper API

```typescript
import {
  getFeatures, getFeatureById, getAdoTraceability, getFeatureByAdoSuiteId,
  FeatureCategory, Lifecycle,
} from "../tests/integration/featureRegistry";

getFeatures({ category: FeatureCategory.Bot });
getFeatures({ lifecycle: Lifecycle.Provision });
getFeatures().filter(f => f.capabilities.includes("ai-powered"));
getFeatureById("default-bot");
getFeatureByAdoSuiteId(24569101);
```

## Test generation

`scaffoldTest()` and `provisionTest()` factories in `tests/integration/testBuilders.ts` consume the registry to generate Mocha tests automatically. Adding a feature entry adds tests; the coverage report (HTML at `coverage/integration/integration-coverage.html`) flags uncovered combinations.

## ADO traceability

Each feature links to its E2E test suite in the **TeamsFx Test Plan (TTK&CLI)** ADO plan (id 24569079). The HTML coverage report includes a traceability table mapping features → ADO suites → manual test case counts.

## Two engine views

| View | Source |
|------|--------|
| v3 — `TemplateNames` enum | `packages/fx-core/src/component/generator/templates/templateNames.ts` |
| v4 — `TemplateRegistry` | `packages/core-next/src/templates/registry.ts` |

`features.json` is shared. The validator in `featureRegistry.ts` cross-checks against the v3 enum; v4 descriptors are validated by `tests/unit/templates/descriptors.test.ts`.

## Adding a new feature

1. Add a folder under `templates/vsc/{lang}/{templateDir}/`.
2. Add a `TemplateDescriptor` to a file under `packages/core-next/src/templates/descriptors/`.
3. Add a `TemplateNames` enum entry (v3) and update generator metadata.
4. Add an entry to `.dev/features.json` with full metadata, ADO suite link, and capabilities.
5. Run `npm run test:integration` in `packages/fx-core` — coverage report flags it as untested until tests pass.

See [07-contributing/adding-a-template.md](../../07-contributing/adding-a-template.md) for the full playbook.
