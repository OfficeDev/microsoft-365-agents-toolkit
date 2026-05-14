# Adding a template

Authoritative source: [`templates.instructions.md`](../../.github/instructions/templates.instructions.md).

## Steps

### 1. Create the template folder

```
templates/vsc/{language}/{template-name}/
```

| Language | Folder |
|----------|--------|
| TypeScript | `ts/` |
| JavaScript | `js/` |
| Python | `python/` |
| C# | `csharp/` |
| Shared | `common/` |

### 2. Author the template files

Standard layout:

```
{template-name}/
├── appPackage/                # manifest.json, declarativeAgent.json, icons
├── infra/                     # azure.bicep, parameters
├── env/                       # .env.dev, .env.local
├── m365agents.yml
├── m365agents.local.yml
└── src/
```

Files with Mustache placeholders use the `.tpl` suffix (stripped on output). Use the standard placeholders documented in [05-engineering/package-reference/templates.md](../05-engineering/package-reference/templates.md).

### 3. Register in v4 (`core-next`)

Add a `TemplateDescriptor` to a file under [`packages/core-next/src/templates/descriptors/`](../../packages/core-next/src/templates/descriptors/) (group by category):

```ts
templateRegistry.register({
  id: "cea/my-template",
  name: "My CEA Template",
  category: "cea",
  description: "...",
  languages: ["typescript", "javascript"],
  templateName: "my-template",          // MUST match folder name
  capabilities: ["conversational"],
});
```

If the template needs special scaffolding (e.g. OpenAPI flow, MetaOS upgrade), provide `scaffoldFn`.

### 4. Register in v3 (`fx-core`)

Add a `TemplateNames` enum entry and update `generatorProvider.ts` so a generator activates on the new name.

Add metadata in `packages/fx-core/src/component/generator/templates/metadata/{category}.ts`:

```ts
{
  id: "my-template-ts",                  // folder name + language
  name: TemplateNames.MyTemplate,        // enum entry
  language: "typescript",
  displayName: "My Template",
  description: "...",
}
```

### 5. Update the question model (v3)

Add a node in the appropriate question tree:

- VS Code: `packages/fx-core/src/question/scaffold/vsc/createRootNode.ts`
- Visual Studio: `packages/fx-core/src/question/scaffold/vs/createRootNode.ts`

(v4 picks the new template up automatically via `buildQuestionTree(registry)`.)

### 6. Add to `features.json`

Add an entry to [`.dev/features.json`](../../.dev/features.json):

```json
{
  "id": "my-template-ts",
  "name": "My CEA Template",
  "description": "...",
  "category": "Custom Engine Agent",
  "templateName": "my-template",
  "languages": ["typescript", "javascript"],
  "lifecycles": ["scaffold", "provision", "deploy", "publish"],
  "capabilities": ["conversational"],
  "generatedFiles": ["src/index.ts", "appPackage/manifest.json"],
  "entryPoints": ["fx-core", "cli", "vscode"],
  "adoSuiteId": 0
}
```

### 7. Test

```bash
cd packages/fx-core
npm run test:integration                # validates featureRegistry + scaffold
cd packages/core-next
npm run test:unit                       # descriptors.test.ts validates registration
```

Coverage report at `packages/fx-core/coverage/integration/integration-coverage.html` will flag the new feature as covered (or not).

### 8. Build

```bash
cd templates
npm run build                           # rebuilds all template ZIPs
```

The ZIPs propagate to:

- `packages/fx-core/templates` (v3)
- `packages/core-next/templates/fallback/` (v4)

### 9. Verify E2E

If applicable, add a cli-next E2E test:

```bash
cd packages/cli-next
npm run test:e2e -- --grep "my-template"
```
