---
description: "Use when creating or modifying project scaffolding templates, template metadata, Mustache placeholders, or template build scripts."
applyTo: "templates/**"
---

# Template Conventions

## Adding a New Template

### Step 1 — Create the Template Folder

```
templates/vsc/{language}/{template-name}/
```

| Platform | Folder |
|----------|--------|
| VS Code | `vsc/` |
| Visual Studio | `vs/` |

| Language | Folder |
|----------|--------|
| TypeScript | `ts/` |
| JavaScript | `js/` |
| Python | `python/` |
| C# | `csharp/` |
| Shared | `common/` |

### Step 2 — Template File Structure

Every template should include:

```
{template-name}/
  ├── appPackage/           # App manifest files
  ├── infra/                # Bicep infrastructure
  ├── env/                  # Environment files (.env.dev, etc.)
  ├── m365agents.yml        # Action definitions
  ├── m365agents.local.yml  # Local debug actions
  └── src/                  # Application source code
```

### Step 3 — Register Metadata

Add entry in `templates/src/metadata/{category}.ts`:

```typescript
{
  id: "my-template-ts",              // Must match folder name + language
  name: TemplateNames.MyTemplate,    // Add to templateNames.ts enum
  language: "typescript",
  displayName: "My Template",
  description: "What this template creates",
}
```

Categories: `declarativeAgent.ts`, `teams.ts`, `customEngineAgent.ts`, `graphConnector.ts`, `wxp.ts`, `special.ts`

### Step 4 — Update Question Model

Add to the appropriate question tree in fx-core:
- VS Code: `packages/fx-core/src/question/scaffold/vsc/createRootNode.ts`
- Visual Studio: `packages/fx-core/src/question/scaffold/vs/createRootNode.ts`

### Step 5 — Custom Generator (Optional)

Only needed for pre/post-processing. See fx-core instructions for the generator pattern.

## Placeholder Conventions

Files with Mustache placeholders **must** use the `.tpl` suffix (stripped at output).

| Delimiter | Usage |
|-----------|-------|
| `{{ }}` | Standard Mustache |
| `{% %}` | Legacy (still supported) |

Common placeholders:

| Placeholder | Value |
|-------------|-------|
| `{{appName}}` | Application name |
| `{{SafeProjectName}}` | URL-safe name |
| `{{SafeProjectNameLowerCase}}` | Lowercase URL-safe name |
| `{{TargetFramework}}` | e.g., `net8.0` |
| `{{useOpenAI}}` / `{{useAzureOpenAI}}` | Model provider toggle |
| `{{DeclarativeCopilot}}` | Declarative Agent flag |
| `{{NewProjectTypeName}}` | C# project type name (default: `M365Agent`) |
| `{{NewProjectTypeExt}}` | C# project file extension (default: `atkproj`) |
| `{{SolutionName}}` | C# solution name (defaults to `appName`) |
| `{{PlaceProjectFileInSolutionDir}}` | C# project-in-solution layout flag |

Full map: `packages/fx-core/src/component/generator/templates/templateReplaceMap.ts` (v3) and `packages/core-next/src/templates/scaffold/replaceMap.ts` (v4)

## Template ID Naming

Pattern: `{template-folder-name}-{language}`

- Language suffix omitted for `common` and `none` languages
- Examples: `declarative-agent-basic`, `weather-agent-ts`, `weather-agent-js`

## Build Process

```bash
npm run build                # Full build (VSC + VS + distribute)
npm run generate:vsc         # VS Code templates only
npm run generate:vs          # Visual Studio templates only
```

Output is distributed to `packages/fx-core/templates` via `npm run distribute`.

## v4 Template System (`packages/core-next`)

In the v4 architecture, templates are registered declaratively via `TemplateRegistry`
in `packages/core-next/src/templates/`. Each template is a `TemplateDescriptor` record
with metadata (id, name, languages, category, capabilities, generated files).

The v4 system consumes the same `features.json` source of truth but replaces the
implicit generator-activation pattern with an explicit registry lookup. Template files
still live in `templates/vsc/` and `templates/vs/` — only the registration mechanism
changes.

### Scaffold Pipeline (implemented)

The scaffold system lives in `packages/core-next/src/templates/scaffold/`:

| File | Purpose |
|------|---------|
| `scaffolder.ts` | `scaffoldTemplates()` — main entry: download ZIP → unzip with transform → render |
| `download.ts` | `downloadTemplate()` — fetch from remote URL with retry, local fallback |
| `render.ts` | Mustache renderer — processes `.tpl`-suffixed files, preserves undefined placeholders |
| `replaceMap.ts` | `getTemplateReplaceMap()` — standard placeholder map (`appName`, `SafeProjectName`, etc.) |
| `types.ts` | `TemplateInfo`, `ScaffoldContext`, `TemplateConfig` types |

### Built-in Descriptors (implemented)

`packages/core-next/src/templates/descriptors/declarativeAgent.ts` registers 11 Declarative
Agent `TemplateDescriptor` records with factory scaffold functions. Call
`registerBuiltinTemplates()` to populate the `TemplateRegistry` at startup.

### v3 / v4 Coexistence

Until migration is complete (gated by `TEAMSFX_V4_CORE` feature flag), continue following
the v3 steps above for new templates. The v4 scaffold system is operational but consumers
(CLI, VS Code extension) have not yet switched to the v4 pipeline.
