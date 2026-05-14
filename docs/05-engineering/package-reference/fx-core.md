# `packages/fx-core` — `@microsoft/teamsfx-core`

The **v3 engine**. Centralises everything shared by the v3 IDE extensions and CLI.

## Conventions source

[`.github/instructions/fx-core.instructions.md`](../../../.github/instructions/fx-core.instructions.md) and [`codebase.instructions.md`](../../../.github/instructions/codebase.instructions.md).

## Architecture

```
FxCore (entry point — src/core/FxCore.ts)
  → Coordinator (orchestrates multi-step workflows)
    → Generators (src/component/generator/)
    → Drivers (src/component/driver/)
    → Question Model (src/question/scaffold/)
  → EnvironmentManager (multi-env support)
  → FeatureFlagManager (gated rollouts)
  → globalVars.TOOLS singleton (LogProvider, TelemetryReporter, TokenProvider, UserInteraction)
```

## Generators

Located in `src/component/generator/`. Each scaffolds a project template.

```typescript
export class MyGenerator extends DefaultTemplateGenerator {
  componentName = "my-generator";
  public override activate(context, inputs) { return inputs[QuestionNames.TemplateName] === TemplateNames.My; }
  public override async getTemplateInfos(...) { ... }
  public override async post(...) { ... }
}
```

Register in `generatorProvider.ts` — **order matters** (first activated wins). v4 replaces this implicit ordering with explicit `TemplateRegistry` lookup.

## Drivers

Encapsulate external service interactions in `src/component/driver/`:

| Driver dir | Purpose |
|-----------|---------|
| `aad/` | Entra ID app registration |
| `arm/` | Azure Resource Manager |
| `teamsApp/` | Teams app packaging / publishing |
| `deploy/` | Deployment to Azure |
| `apiKey/` | API key registration |
| `oauth/` | OAuth configuration |

Drivers return `Result<T, FxError>`, use `getLocalizedString()`, log via `LogProvider`, and follow EAFP filesystem rules.

## Question model

Tree-shaped, per platform:

```
src/question/scaffold/
├── vsc/createRootNode.ts
├── vs/createRootNode.ts
└── cli/
```

## Errors

Organised by domain in `src/error/`:

```
error/
├── common.ts     # cross-cutting
├── arm.ts        # ARM deploy
├── azure.ts
├── teamsApp.ts
├── yml.ts
└── ...
```

Each file exports error factory functions or class constructors with stable `name`s.

## Feature flags

`featureFlagManager` singleton in `src/common/featureFlags.ts`. Headline flag: `TEAMSFX_V4_CORE`.

## Globals

`src/common/globalVars.ts` — `TOOLS` singleton. Set once at consumer init; access via import. v4 replaces with injected `AtkContext`.

## Exports

`index.ts` uses **selective named exports** — only what the CLI / VS Code extension / server need.

## Testing

50+ granular test scripts (`test:core`, `test:component`, `test:bot`, ...). Integration tests in `tests/integration/` driven by [`featureRegistry.ts`](../../../packages/fx-core/tests/integration/featureRegistry.ts) which loads `.dev/features.json`.

See [04-specs/data-model/entities/feature-registry.md](../../04-specs/data-model/entities/feature-registry.md).
