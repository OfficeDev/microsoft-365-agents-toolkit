# Architecture Reference

## Package Dependency Flow

```
@microsoft/teamsfx-api (contracts)
        ↓
@microsoft/teamsfx-core (engine)
        ↓
┌───────┴───────┐
CLI          VS Code Extension
```

Changes to `api` require rebuilding all downstream packages.

## Core Engine Architecture (fx-core)

### Generators

Located in `packages/fx-core/src/component/generator/`. Each generator scaffolds project templates.

**Class pattern:**
```
DefaultTemplateGenerator (base)
  ├── activate()          — Should this generator handle the input?
  ├── getTemplateInfos()  — Prepare template name, language, replacement map
  └── post()              — Post-processing after template download
```

**Registration:** `generatorProvider.ts` — ordered array, first match wins:
```
TdpGenerator → DefaultTemplateGenerator → OfficeAddinGeneratorNew →
SPFxGeneratorNew → SPFxGeneratorImport → SsrTabGenerator →
DeclarativeAgentWithExistingApiSpecGenerator → ... → CombinedProjectGenerator
```

### Drivers

Encapsulate interactions with external systems:
- **AAD Driver** — Azure AD app registration
- **ARM Driver** — Azure Resource Manager deployments
- **Teams App Driver** — Teams app packaging and publishing
- **Bot Registration Driver** — Bot Framework registration

### Question Model

Each platform (VS Code, Visual Studio, CLI) has its own question tree in:
`packages/fx-core/src/question/scaffold/{platform}/createRootNode.ts`

Flow: User selection → `TemplateName` input → Generator activation → Template scaffolding

### Project Types Supported

1. **Declarative Agents** — Agentic AI copilot extensions
2. **Custom Engine Agents** — Self-hosted agent backends
3. **Bots** — Teams conversational bots
4. **Tabs** — Teams tab applications
5. **Message Extensions** — Rich interaction extensions
6. **Office Add-ins** — Excel, Word, Outlook extensions
7. **SPFx Solutions** — SharePoint Framework solutions
8. **Copilot Connectors** — Graph connector templates

## ESLint Configuration Stack

All packages extend from `packages/eslint-plugin-teamsfx/config/`:

| Config | Rules |
|--------|-------|
| `shared.js` | Base TypeScript + Prettier + no-secrets + import cycles |
| `header.js` | Copyright header enforcement |
| `promise.js` | Async/await correctness (no floating promises) |
| `type.js` | Strict TypeScript type checking |

## Test Stack

- **Framework:** Mocha + Chai (assertions) + Sinon (mocking)
- **Coverage:** NYC
- **Reporting:** Codecov (CI integration)
- **Custom:** `packages/extra-shot-mocha` for enhanced test utilities

## Error Handling Pattern

All operations return `Result<T, FxError>` (from `@microsoft/teamsfx-api`). Never throw exceptions in core logic.

**Two error types:**

| Type | When to Use | User Impact |
|------|-------------|-------------|
| `UserError` | Recoverable — bad input, missing config, auth issues | User sees help link |
| `SystemError` | Unrecoverable — service failures, internal bugs | User sees issue link |

**Creating errors:**
```typescript
import { UserError, SystemError } from "@microsoft/teamsfx-api";

// UserError — user can fix it
new UserError({
  source: "my-component",
  name: "MissingConfigError",
  message: getDefaultString("error.missingConfig", configName),
  displayMessage: getLocalizedString("error.missingConfig", configName),
  helpLink: "https://aka.ms/...",
});

// SystemError — internal failure
new SystemError({
  source: "my-component",
  name: "ServiceCallFailed",
  message: getDefaultString("error.serviceFailed"),
  displayMessage: getLocalizedString("error.serviceFailed"),
});
```

**Returning results:**
```typescript
import { ok, err } from "@microsoft/teamsfx-api";

async function doWork(): Promise<Result<string, FxError>> {
  if (invalid) return err(new UserError(...));
  return ok("success");
}
```

## Localization

User-facing strings are centralized in `package.nls.json` files:
- `package.nls.json` — English (default, always required)
- `package.nls.{locale}.json` — Translations (fr, de, ja, etc.)

**Usage:**
```typescript
import { getLocalizedString, getDefaultString } from "../common/localizeUtils";

// For display messages (shown to user in their locale)
const displayMsg = getLocalizedString("error.fileNotFound", fileName);

// For telemetry/logs (always English)
const logMsg = getDefaultString("error.fileNotFound", fileName);
```

String keys support `util.format()` parameter substitution: `"error.fileNotFound": "File %s not found"`
