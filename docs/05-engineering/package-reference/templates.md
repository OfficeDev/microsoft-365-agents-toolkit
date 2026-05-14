# `templates/` — project scaffolds

The on-disk source for every project template the toolkit can scaffold. Built into `packages/fx-core/templates` (v3) and consumed via the `templates/fallback/*.zip` ship set in `core-next` (v4).

## Conventions source

[`.github/instructions/templates.instructions.md`](../../../.github/instructions/templates.instructions.md).

## Layout

```
templates/
├── vsc/             # VS Code templates
│   ├── ts/          # TypeScript
│   ├── js/          # JavaScript
│   ├── python/      # Python
│   └── common/      # Shared (e.g. DA without code)
├── vs/              # Visual Studio templates (C#)
├── configs/         # Shared configuration files
├── scripts/         # Build scripts
├── unused/          # Retired templates kept for reference
└── build/           # Build output (zips for download + fallback)
```

Per-template structure:

```
{template-name}/
├── appPackage/           # manifest.json, declarativeAgent.json, icons
├── infra/                # Bicep templates (azure.bicep, parameters)
├── env/                  # .env.dev, .env.local
├── m365agents.yml
├── m365agents.local.yml
└── src/                  # application source
```

## Placeholders

Files with Mustache placeholders use `.tpl` suffix (stripped at output).

| Delimiter | Usage |
|-----------|-------|
| `{{ }}` | Standard Mustache |
| `{% %}` | Legacy (still supported) |

Common placeholders:

| Placeholder | Value |
|-------------|-------|
| `{{appName}}` | Application name |
| `{{SafeProjectName}}` | URL-safe name |
| `{{SafeProjectNameLowerCase}}` | Lowercase URL-safe |
| `{{TargetFramework}}` | e.g. `net8.0` |
| `{{useOpenAI}}`, `{{useAzureOpenAI}}` | Model provider toggle |
| `{{DeclarativeCopilot}}` | DA flag |
| `{{NewProjectTypeName}}` | C# project type name |
| `{{NewProjectTypeExt}}` | C# project file extension |
| `{{SolutionName}}` | C# solution name |
| `{{PlaceProjectFileInSolutionDir}}` | C# layout flag |

Full map: `packages/fx-core/src/component/generator/templates/templateReplaceMap.ts` (v3), `packages/core-next/src/templates/scaffold/replaceMap.ts` (v4).

## Template ID naming

Pattern: `{template-folder-name}-{language}` — language suffix omitted for `common` and `none`.

Examples: `declarative-agent-basic`, `weather-agent-ts`, `weather-agent-js`.

## Build

```bash
cd templates
npm run build                # full build (VSC + VS + distribute)
npm run generate:vsc         # VS Code only
npm run generate:vs          # Visual Studio only
```

Output is distributed to `packages/fx-core/templates` via `npm run distribute`. v4 ships fallback ZIPs in `packages/core-next/templates/fallback/` (copied from `templates/build/fallback/`).

## Adding a template

See [07-contributing/adding-a-template.md](../../07-contributing/adding-a-template.md).

## v4 specifics

Templates are still authored under `templates/vsc/{lang}/` — only the **registration** mechanism differs. v4 uses `TemplateDescriptor` records in `core-next/src/templates/descriptors/` instead of generator-activation order.

`templateName` constants must match actual folder names. For example, `DATemplateNames.Basic = "declarative-agent-basic"` — not the legacy display name `"copilot-gpt-basic"`.
