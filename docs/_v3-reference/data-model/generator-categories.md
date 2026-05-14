# v3 generator categories

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).

Source: [`packages/fx-core/src/component/generator/`](../../../packages/fx-core/src/component/generator/).

## Subdirectories

| Directory | Generator family |
|-----------|------------------|
| `combinedProject/` | Composite generators (e.g. project + plugin in one go) |
| `configFiles/` | Config-file scaffolds (`atk init`) |
| `declarativeAgent/` | DA-specific generators (basic DA, plugin variants, MetaOS, TypeSpec) |
| `officeAddin/` | Office Add-in generators (Word/Excel/PowerPoint/Outlook task panes, custom functions) |
| `openApiSpec/` | OpenAPI-spec-driven generators (RAG-from-spec, ME-from-spec, plugin-from-spec) |
| `other/` | Generic / shared generators (bots, tabs, samples) |
| `spfx/` | SPFx project generators |
| `templates/` | Template metadata + name registry |

## Notes

- Generators extend `DefaultTemplateGenerator` and implement `activate(context, inputs)` + `getTemplateInfos(...)` + `post(...)`.
- Registered in `generatorProvider.ts`; **first activated wins** — collisions are silent.
- Each generator owns its question subtree and post-scaffold tasks.

## Why this is forbidden as v4 design input

The v3 generator activation pattern is the single most v3-specific anti-pattern. v4 replaces it with:

- An explicit **`TemplateRegistry`** (lookup by ID, no activation order).
- **`TemplateDescriptor`** records with declared `questions: QuestionSpec[]` and optional `scaffoldFn`.
- An auto-generated CLI subcommand tree from the registry (`buildNewCommands(registry)`).
- An auto-generated question tree from the registry (`buildQuestionTree(registry)`).

If a v4 design proposes a generator class hierarchy, that's a sign someone is replicating this v3 shape instead of using the registry. Code review red flag.
