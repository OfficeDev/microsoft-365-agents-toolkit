# Template system

Source: [`packages/core-next/src/templates/`](../../../packages/core-next/src/templates/).

## Layout

```
templates/
├── registry.ts         — TemplateRegistry, TemplateDescriptor
├── types.ts            — TemplateDescriptor type
├── scaffold/           — scaffold pipeline (download → render → unzip)
│   ├── scaffolder.ts       — scaffoldTemplates() entry
│   ├── download.ts         — resolveTemplateUrl(), fetchZip(), loadLocalFallback(), unzipWithTransform()
│   ├── render.ts           — Mustache rendering (.tpl files, preserves undefined vars)
│   ├── replaceMap.ts       — getTemplateReplaceMap()
│   └── types.ts            — TemplateInfo, ScaffoldContext, TemplateConfig, convertToLangKey()
├── descriptors/        — built-in template registrations (24)
│   ├── declarativeAgent.ts — 12 DA descriptors (incl. metaos-upgrade)
│   ├── bot.ts              — 1 (echo)
│   ├── tab.ts              — 1 (basic)
│   ├── aiAgent.ts          — 3 (chat, RAG x2)
│   ├── engineAgent.ts      — 3 (basic, weather, collaborator)
│   ├── connector.ts        — 1 (graph)
│   ├── messageExtension.ts — 1 (search ME)
│   ├── openApi.ts          — 3 (DA, AI agent, ME from spec)
│   ├── foundry.ts          — 1 (Foundry agent)
│   └── index.ts            — registerBuiltinTemplates()
└── openApi/            — OpenAPI scaffolding
    ├── specParserAdapter.ts     — interface + Stub + factory
    ├── realSpecParserAdapter.ts — backed by inline specParser/
    ├── scaffoldFn.ts            — makeOpenApiScaffoldFn()
    └── index.ts
```

## Scaffold pipeline

```
scaffoldTemplates(descriptor, replaceMap, ctx)
  ├─ resolveTemplateUrl()        → remote URL
  ├─ fetchZip(url)              → ZIP buffer
  │   └─ on failure: loadLocalFallback(lang) from bundled fallback ZIPs
  ├─ unzipWithTransform(zip)
  │   ├─ Zip-Slip guard (entry name + path.resolve check)
  │   ├─ filter by {templateName}/ prefix
  │   ├─ strip prefix before write
  │   └─ render Mustache for .tpl files (strip suffix)
  └─ write files to destination
```

## Fallback strategy

`resolveFallbackDir()` resolves the fallback dir in order:

1. Explicit param.
2. `TEMPLATE_FALLBACK_DIR` env var.
3. Bundled `templates/fallback/` (resolved by `getTemplatesFolder()` in `src/folder.ts`).

Fallback ZIPs ship in `packages/core-next/templates/fallback/`: `common.zip`, `ts.zip`, `js.zip`, `python.zip`, `csharp.zip`. Listed in `package.json` `files` field so they survive npm pack.

## Language ZIP shape

Each language ZIP contains **all** templates for that language as subdirectories:

```
ts.zip
├── default-bot/
├── custom-copilot-basic/
├── weather-agent/
└── ...
```

The scaffold pipeline auto-filters to the requested template folder and strips the prefix before writing.

## `convertToLangKey()`

Maps full language names to ZIP keys: `"typescript"` → `"ts"`, `"javascript"` → `"js"`, `"csharp"`, `"python"`, `"common"`. Used by every descriptor.

## Mustache rendering

- `.tpl` files are rendered then renamed to drop the suffix.
- Non-`.tpl` files copy through unchanged.
- Undefined placeholders are preserved (Mustache default) — useful when a template references env vars resolved later by the lifecycle.

## Replace map

`getTemplateReplaceMap()` returns the standard placeholder map: `appName`, `SafeProjectName`, `SafeProjectNameLowerCase`, `TargetFramework`, etc. Per-template factories (declared via `TemplateDescriptor.scaffoldFn` or default flow) extend the base map with their own keys.

## OpenAPI templates

Three OpenAPI-backed templates (`da/api-plugin-from-spec`, `ai-agent/rag-from-spec`, `me/from-spec`) use `makeOpenApiScaffoldFn()` instead of the default scaffold pipeline:

```
makeOpenApiScaffoldFn(adapter)
  → validate spec (per-project-type validator)
  → scaffold base files (default pipeline)
  → parse spec
  → generate per-operation artifacts (manifest, plugin)
  → write
```

`adapter` is a `SpecParserAdapter`. Default: `RealSpecParserAdapter` backed by inline `specParser/`. Tests use `StubSpecParserAdapter`.

## Tests

- `tests/unit/templates/descriptors.test.ts` — 43 tests covering registration + metadata + name validation.
- `tests/unit/templates/openApi.test.ts` — OpenAPI descriptor + scaffoldFn tests.
- `tests/unit/templates/registry.test.ts` — `TemplateRegistry` CRUD.

E2E scaffold tests: 9/9 verified via cli-next (bot TS/JS/Python, DA basic, AI chat, CEA basic, CEA weather Python, tab basic, connector graph).
