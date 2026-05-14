# Template Conventions

## Adding a New Template

### Step 1 — Create Template Folder

```
templates/vsc/{language}/{template-name}/
```

- **Platform folders:** `vsc/` (VS Code), `vs/` (Visual Studio)
- **Language folders:** `ts/`, `js/`, `csharp/`, `python/`, `common/`
- Files with Mustache placeholders **must** end in `.tpl` (stripped at output)
- Static files (e.g., `tsconfig.json`) are copied as-is

### Step 2 — Register Metadata

Add entry in `templates/src/metadata/{category}.ts`:

```typescript
{
  id: "my-template-ts",                  // Must match folder name + language
  name: TemplateNames.MyTemplate,        // Add to templateNames.ts enum
  language: "typescript",
  displayName: "My Template",
  description: "What this template creates",
}
```

Categories: `declarativeAgent.ts`, `teams.ts`, `customEngineAgent.ts`, `graphConnector.ts`, `wxp.ts`, `special.ts`

### Step 3 — Update Question Model

**VS Code:** Edit `packages/fx-core/src/question/scaffold/vsc/createRootNode.ts`
- Add option with `data: TemplateNames.YourTemplate` to the appropriate project type node

**Visual Studio:** Edit `packages/fx-core/src/question/scaffold/vs/createRootNode.ts`

### Step 4 — Custom Generator (Optional)

Only needed if the template requires pre/post-processing. Extend `DefaultTemplateGenerator`:

```typescript
export class MyGenerator extends DefaultTemplateGenerator {
  componentName = "my-generator";

  public override activate(context: Context, inputs: Inputs): boolean {
    return inputs[QuestionNames.TemplateName] === TemplateNames.MyTemplate;
  }

  public override async getTemplateInfos(...): Promise<Result<TemplateInfo[], FxError>> {
    // Build replacement map, return template info
  }

  public override async post(...): Promise<Result<GeneratorResult, FxError>> {
    // Post-processing after template download
  }
}
```

Register in `packages/fx-core/src/component/generator/generatorProvider.ts` — **order matters** (first activated wins).

## Placeholder Conventions

| Delimiter | Usage |
|-----------|-------|
| `{{ }}` | Mustache (standard) |
| `{% %}` | Legacy (still supported) |

Common placeholders:
- `{{appName}}` — App name
- `{{SafeProjectName}}` / `{{SafeProjectNameLowerCase}}` — URL-safe variants
- `{{TargetFramework}}` — e.g., `net8.0`
- `{{useOpenAI}}` / `{{useAzureOpenAI}}` — Model provider toggle
- `{{DeclarativeCopilot}}` — DA template flag

See `packages/fx-core/src/component/generator/templates/templateReplaceMap.ts` for the full map.

## Template ID Naming

Pattern: `{template-folder-name}-{language}`

- Language omitted for `common` and `none` languages
- Examples: `declarative-agent-basic`, `weather-agent-ts`, `weather-agent-js`
