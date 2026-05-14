# Template descriptor

A `TemplateDescriptor` is a record in the v4 `TemplateRegistry`. It declares everything the engine needs to scaffold a template and surface it through CLI / VS Code.

Source: [`packages/core-next/src/templates/registry.ts`](../../../packages/core-next/src/templates/registry.ts) and `types.ts`.

## Shape

```typescript
type TemplateDescriptor = {
  id: string;                    // unique, e.g. "da/basic", "cea/weather"
  name: string;                  // human-readable
  category: TemplateCategory;    // da | cea | ai | me | bot | tab | connector | addin
  description: string;
  languages: TemplateLanguage[]; // ["typescript", "javascript", "python", "csharp", "common"]
  templateName: string;          // folder name under templates/vsc/{lang}/
  questions?: QuestionSpec[];    // template-specific extra questions
  capabilities?: string[];       // semantic tags
  testable?: boolean;            // default true; false → not E2E-tested
  scaffoldFn?: ScaffoldFn;       // optional override for non-standard scaffolding (OpenAPI, MetaOS, ...)
};
```

## Registration

```typescript
templateRegistry.register({
  id: "cea/basic",
  name: "Basic Custom Engine Agent",
  category: "cea",
  description: "Echo-style starting point on Microsoft Agents SDK",
  languages: ["typescript", "javascript", "python"],
  templateName: "custom-copilot-basic",
  capabilities: ["conversational"],
});
```

All built-in descriptors register via `registerBuiltinTemplates()` in [`packages/core-next/src/templates/descriptors/index.ts`](../../../packages/core-next/src/templates/descriptors/index.ts). 26 descriptors are registered today across 9 files.

## Auto-generated artifacts

Adding a `TemplateDescriptor` automatically:

1. Adds it to the question tree built by `buildQuestionTree(registry)`.
2. Generates a CLI subcommand under `atk new <category>` via `buildNewCommands(parent, registry)`.
3. Includes it in `atk list templates` output.

## `templateName` must match folder

The `templateName` field is the folder under `templates/vsc/{lang}/` that the scaffold pipeline filters to. It must match the **actual** folder name on disk, not a legacy display name.

Example: `templateName: "declarative-agent-basic"` (matches `templates/vsc/common/declarative-agent-basic/`), **not** `"copilot-gpt-basic"` (legacy name).

## `testable` flag

Defaults to `true`. Set to `false` for templates that:

- Need interactive input that can't be defaulted (e.g. `apiSpecPath`).
- Lack required artefacts (e.g. missing `appPackage/manifest.json`).

Currently `false` for: `da/api-plugin-from-spec`, `ai-agent/rag-from-spec`, `me/from-spec`, `da/graph-connector` (artefact gap).

## Scaffold function override

Most descriptors don't set `scaffoldFn` — the default scaffold pipeline (download → filter → render Mustache) is enough. Overrides are used when:

- OpenAPI scaffolding (uses `makeOpenApiScaffoldFn()` with a `SpecParserAdapter`).
- MetaOS upgrade flow (`copyExistMetaOSProject()`, `extendToDA()`, `unifyProjectID()`).

See [`templates/openApi/`](../../../packages/core-next/src/templates/openApi/) and [`helpers/metaOSHelper.ts`](../../../packages/core-next/src/helpers/metaOSHelper.ts).
