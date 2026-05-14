---
description: "Use when editing, creating, or reviewing code in any package: manifest, api, fx-core, core-next, server, vscode-extension, cli, cli-next, or templates. Covers monorepo conventions, error handling, testing, formatting, imports, and commit rules."
applyTo: "packages/**/*.ts, templates/**/*.ts"
---

# Microsoft 365 Agents Toolkit — Codebase Conventions

## Source-of-Truth Workflow (v3 vs v4)

The two engines have **different change workflows**. Get this right before editing.

| Engine | Workflow | Source of truth |
|--------|----------|-----------------|
| **v3** (`api`, `manifest`, `fx-core`, `cli`, `vscode-extension`, `server`) | **Code-first.** Existing shipping product; the code is the running truth. Update docs *after* (or in the same PR as) a code change. | Code → instructions → skills → `.dev/features.json` → `docs/` prose |
| **v4** (`core-next`, `cli-next`) | **Design-first.** All non-trivial AI-driven changes start in `docs/`. The design page is authoritative until the code lands. | `docs/` design pages and ADRs → code → instructions → skills |

### v4 design-first procedure

For any non-trivial v4 change (new operation, new driver, new template descriptor, new lifecycle stage, API change, refactor, behaviour change):

1. **Locate or write the design page** under [`docs/02-architecture/`](../../docs/02-architecture/README.md) (architecture or runtime view) or [`docs/04-specs/data-model/`](../../docs/04-specs/data-model/README.md) (entity / contract).
2. **For architectural decisions**, file an ADR under [`docs/02-architecture/09-architecture-decisions/`](../../docs/02-architecture/09-architecture-decisions/README.md) using the [ADR template](../../docs/02-architecture/09-architecture-decisions/adr-template.md).
3. **For new capabilities**, update [`docs/01-product/capabilities-matrix.md`](../../docs/01-product/capabilities-matrix.md) and add an entry to [`.dev/features.json`](../../.dev/features.json).
4. **Land the docs change in the same PR as the code** — never as a follow-up.
5. If the design contradicts an existing instruction file, update both in the same PR.

Use the [`vibe-coding`](../skills/vibe-coding/SKILL.md) skill for the full step-by-step workflow (design → implement → test → docs → PR) and the gates that must be green.

For strategic context on why v4 is shaped this way — cli-next-first, inverted test pyramid, vibe-coding-friendly architecture — read [`docs/05-engineering/v4-design-strategy.md`](../../docs/05-engineering/v4-design-strategy.md).

Trivial v4 changes (bug fix, dependency bump, test-only edit, doc-only edit) do **not** require a forward design page.

### Forbidden inputs to v4 design

When designing v4, you may NOT use the following as inputs — they are quarantined under [`docs/_v3-reference/`](../../docs/_v3-reference/README.md) explicitly to enforce this:

- The `FxCore` class signature.
- The v3 lifecycle YAML in-memory model (`RawProjectModel`, `ProjectModel`, `DriverDefinition`, `ILifecycle`).
- The v3 `StepDriver` interface and `DriverContext` shape.
- The v3 generator activation pattern (first-activated-wins ordering).
- The full v3 error catalogue's organising structure (individual error *names* may be reused for back-compat — the catalogue's grouping must not).
- Existing v3 Bicep templates (re-design from topology requirements, not from existing `.bicep` files).
- The `TOOLS` global singleton pattern.

What you MAY use as v4 design input: [`docs/01-product/`](../../docs/01-product/README.md) (PRD, personas, capabilities), [`docs/01-product/ux/`](../../docs/01-product/ux/README.md) (surfaces, flows, question model, errors), Microsoft 365 platform contracts (manifest schemas, Graph, TDP), upstream SDK contracts. See [`docs/_v3-reference/README.md`](../../docs/_v3-reference/README.md) §"The hard rule".

> **Why the asymmetry:** v3 has years of shipping behaviour that the docs are still catching up to — pretending otherwise would silently lie to contributors. v4 is being built from scratch with the spec in `docs/`, so writing code without first writing the design forfeits the value of the new structure. See [docs/05-engineering/v3-to-v4-migration.md](../../docs/05-engineering/v3-to-v4-migration.md) §"Workflow asymmetry".

## Monorepo Structure

PNPM workspaces + Lerna. Dependency flow:

```
# Current (v3)
api → manifest → fx-core → cli / vscode-extension / server
                         → templates (build output copied into fx-core)

# Next generation (v4) — gated by TEAMSFX_V4_CORE feature flag
core-next → cli-next
         → (future) vscode-extension / server
```

Changing `api` impacts all current downstream packages — rebuild before testing.
Changing `core-next` impacts `cli-next` and future v4 consumers.

## Package Roles

| Package | npm name | Role |
|---------|----------|------|
| `packages/api` | `@microsoft/teamsfx-api` | Contracts, interfaces, error types — the public API surface |
| `packages/manifest` | `@microsoft/app-manifest` | Manifest types, JSON schema validation, converters |
| `packages/fx-core` | `@microsoft/teamsfx-core` | Core engine — generators, drivers, coordinators |
| `packages/server` | `@microsoft/teamsfx-server` | JSON-RPC server bridge to fx-core |
| `packages/vscode-extension` | `ms-teams-vscode-extension` | VS Code extension UI, handlers, tree views |
| `packages/cli` | `@microsoft/m365agentstoolkit-cli` | CLI tool (`atk` binary) — current v1.x |
| `packages/core-next` | `@microsoft/teamsfx-core-next` v4.0.0 | **Next-gen** core engine — merged API contracts, AtkContext, Operation pipeline, TemplateRegistry + **E2E-verified** scaffold system (auto-fallback, per-template filter/prefix-strip, `convertToLangKey()`, bundled fallback ZIPs) + 24 built-in descriptors, question model (buildQuestionTree + traverseQuestionTree + createProjectInteractive), DriverRegistry + createDriver factory + 22 built-in drivers, service clients (TeamsDevPortal, GraphApi, Azure ARM, M365 PackageService), DA module (knowledge/actions/auth/capabilities), lifecycle engine + operations (provisionOp, deployOp, publishOp with composable prerequisites, driver introspection, progress), project creation, environment management, teamsApp operations |
| `packages/cli-next` | `@microsoft/m365agentstoolkit-cli-next` v4.0.0 | **Next-gen** CLI — Commander.js-based, registry-driven command factory, action layer wired to core-next operations (createProject, provision, deploy, publish, env, teamsapp validate/package, list templates). **Bundled with esbuild** (`esbuild.mjs`) for fast startup. |
| `templates/` | — | Scaffolding templates (TS/JS/Python/C#), built into fx-core |

## Bundlers

| Package | Bundler | Config file | Notes |
|---------|---------|-------------|-------|
| `vscode-extension` | **esbuild** | `esbuild.mjs` | Single-file CJS, native MSAL handling |
| `cli-next` | **esbuild** | `esbuild.mjs` | Single-file CJS, externals: keytar, applicationinsights, msal-node-extensions |
| `cli` (v1) | webpack | `webpack.config.js` | Legacy — needs 4GB heap |
| `fx-core` | webpack | `webpack.config.js` | Legacy |
| `vscode-ui` | webpack | — | Legacy |
| `mcp-server` | webpack | — | Legacy |
| `spec-parser` | rollup | `rollup.config.js` | — |
| `core-next` | none | — | Pure `tsc` (bundled inline by cli-next's esbuild) |

## Change Placement

| Change Type | Package |
|-------------|---------|
| New interface or contract | `packages/api` |
| App manifest schema or validation | `packages/manifest` |
| New command, generator, or driver | `packages/fx-core` |
| CLI command handling or prompts (v3) | `packages/cli` |
| New v4 CLI action function (createProject, lifecycle, env, teamsapp) | `packages/cli-next/src/actions/` |
| New v4 CLI command factory or slug mapping | `packages/cli-next/src/commands/factory.ts` |
| New v4 core interface or contract | `packages/core-next/src/api/` |
| New v4 operation, driver, or template descriptor | `packages/core-next` |
| New v4 built-in driver implementation | `packages/core-next/src/drivers/builtin/` |
| New v4 service client (TeamsDevPortal, GraphApi, Azure ARM) | `packages/core-next/src/clients/` |
| New v4 Declarative Agent feature (knowledge, action, auth) | `packages/core-next/src/declarativeAgent/` |
| New v4 lifecycle, YAML action, or lifecycle operation | `packages/core-next/src/lifecycle/` |
| New v4 scaffold or template registration | `packages/core-next/src/templates/` |
| New v4 question factory or traversal logic | `packages/core-next/src/questions/` |
| New v4 project creation or environment operation | `packages/core-next/src/project/` or `src/environment/` |
| New v4 Teams app packaging or publishing | `packages/core-next/src/teamsApp/` |
| New v4 telemetry helper or correlation | `packages/core-next/src/telemetry/` |
| New v4 CLI telemetry transport or sanitisation | `packages/cli-next/src/telemetry/` |
| New v4 secret masking rule | `packages/core-next/src/secretMasker/` |
| New v4 feature flag | `packages/core-next/src/featureFlags/` |
| New v4 localization string or bundle | `packages/core-next/src/localization/` |
| New v4 HTTP client helper or interceptor | `packages/core-next/src/http/` |
| VS Code UI, commands, tree views | `packages/vscode-extension` |
| New project template | `templates/` + metadata in `fx-core` |

## File Headers

Every `.ts` source file must start with:

```typescript
/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */
```

## Error Handling

Use the `Result<T, FxError>` pattern from `neverthrow` (re-exported by `@microsoft/teamsfx-api`):

- `UserError` — recoverable, user-fixable (bad input, missing config)
- `SystemError` — infrastructure/service failures
- Return `ok(value)` or `err(new UserError({...}))`, never throw for expected failures
- Always set `source` to the component/plugin name
- User-facing messages: `getLocalizedString("key", ...params)`, never raw strings
- Propagate errors with context — wrap inner errors via `innerError` property
- Keep error names stable (they appear in telemetry); change `message` for clarity, not `name`

```typescript
// Good
return err(new UserError({
  source: "AppStudioPlugin",
  name: "ManifestValidationError",
  message: getLocalizedString("plugins.appStudio.validationFailed", details),
  helpLink: "https://aka.ms/...",
}));

// Bad — throwing, raw string, no source
throw new Error("manifest validation failed");
```

## Type Safety

- Enable `strict: true` in `tsconfig.json` — all packages already do this
- TypeScript 6.0 (`~6.0.0`) in core-next and cli-next; `moduleResolution: "bundler"` (the `node10` deprecation is resolved)
- Prefer `unknown` over `any` for untyped external data; narrow with type guards
- Use discriminated unions over type assertions when branching on shape
- Avoid `as` casts — use type predicates or `satisfies` operator instead
- Export interfaces from `api` package for cross-package contracts; concrete types stay internal

```typescript
// Good — type predicate
function isUserError(e: FxError): e is UserError {
  return e instanceof UserError;
}

// Bad — unsafe cast
const ue = e as UserError;
```

## Formatting (Prettier)

- Double quotes, semicolons, 2-space indent
- 100-character print width, LF line endings
- Config: `packages/prettier-config`

## Linting (ESLint)

All packages use ESLint flat config (`eslint.config.mjs`) extending `../eslint-plugin-teamsfx/config/shared.mjs` plus `header.mjs`.
Packages may also add `promise.mjs` and/or `type.mjs` overrides.

| Package | ESLint configs used | Special rules |
|---------|--------------------|--------------|
| fx-core | shared + header + promise + type | — |
| cli | shared + header + promise | — |
| core-next | shared + header + promise | `import-x/no-unresolved: off` |
| cli-next | shared + header | `import-x/no-unresolved: off` |

`import-x/no-unresolved` is disabled in core-next and cli-next because the node resolver
cannot follow pnpm `workspace:*` links. TypeScript compilation catches real import errors.

Key rules:
- License header required on every `.ts` file
- No floating promises — `await` or `return` every promise
- No hardcoded secrets (`no-secrets` rule, tolerance 4.5 entropy)
- No import cycles
- Unused variables must be prefixed with `_` (e.g., `_unused`, `_err`) — enforced by `@typescript-eslint/no-unused-vars` with `argsIgnorePattern: "^_"`

Each package also has `tsconfig.eslint.json` (extends `tsconfig.json`, `noEmit: true`, includes `src` + `tests`) and `.prettierrc.js` (extends `../prettier-config`).

## Naming Conventions

| Element | Style | Examples |
|---------|-------|---------|
| Classes / Interfaces | PascalCase | `FxCore`, `UserError`, `LogProvider` |
| Functions / methods | camelCase | `loadFromPath`, `askSubscription` |
| Constants | SCREAMING_SNAKE_CASE | `AppStudioScopes`, `SharePointAppId` |
| Enums | PascalCase with PascalCase members | `FeatureFlagName.MultiEnv` |
| Files | camelCase or PascalCase matching main export | `FxCore.ts`, `featureFlags.ts` |
| Test files | `<sourceFile>.test.ts` | `FxCore.test.ts`, `error.test.ts` |

## Imports

- Relative paths (`./`, `../`) — no path aliases
- Order: external packages → workspace packages → relative imports
- Use workspace protocol references (`workspace:*`) in `package.json`
- Use `import type { ... }` for type-only imports to avoid runtime overhead

```typescript
// Good
import type { FxError } from "@microsoft/teamsfx-api";
import { ok, err } from "neverthrow";
import { featureFlagManager } from "../common/featureFlags";

// Bad — unordered, value import for types
import { featureFlagManager } from "../common/featureFlags";
import { FxError } from "@microsoft/teamsfx-api";
import { ok } from "neverthrow";
```

## Testing

- **Framework:** Mocha + Chai + Sinon (all packages)
- **Coverage:** NYC (Istanbul) with `@istanbuljs/nyc-config-typescript`
- **Test file naming:** `*.test.ts` (all packages: fx-core, core-next, cli-next)
- **Test location:** `tests/` directory mirroring `src/` structure
- **Test types:** `tests/unit/` for unit tests, `tests/integration/` for integration tests
- **Test counts:** core-next: 550 unit + 48 integration; cli-next: 78 unit + 62 integration
- **Lint status:** core-next: 175 warnings (all `no-explicit-any`); cli-next: 32 warnings (all `no-explicit-any`)
- Every new feature or bug fix must include tests
- Run: `cd packages/<pkg> && npm run test:unit`

### Test Config Files (per package)

| File | Purpose |
|------|---------|
| `.mocharc.js` | Mocha settings: `ts-node/register`, spec reporter, `no-experimental-strip-types` |
| `.nycrc` | NYC coverage config: extends `@istanbuljs/nyc-config-typescript`, coverage thresholds |

### Test Scripts

```bash
npm run build               # TypeScript compile + postbuild (eslint --fix + prettier --write)
npm run test:unit           # Unit tests only (with NYC coverage)
npm run test:integration    # Integration tests only (no coverage)
npm run test                # Default — same as test:unit
npm run lint                # ESLint check (0 errors required; warnings acceptable)
npm run format              # Prettier auto-format
npm run format:check        # Prettier check (CI gate)
```

> **`postbuild` hook:** Both core-next and cli-next run `eslint --fix` + `prettier --write`
> automatically after `build`. No separate format step needed after building.

### Test Structure

```typescript
describe("ComponentName", () => {
  // Group by method or behavior
  describe("methodName", () => {
    afterEach(() => {
      sinon.restore(); // Always clean up stubs
    });

    it("should return ok when input is valid", async () => {
      // Arrange
      const input = { ... };
      // Act
      const result = await component.method(input);
      // Assert
      expect(result.isOk()).to.be.true;
    });

    it("should return UserError when config is missing", async () => {
      const result = await component.method({});
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().name).to.equal("MissingConfigError");
    });
  });
});
```

### core-next Test Helper

`tests/unit/testHelper.ts` exports `createMockContext()` — creates a fully-stubbed `AtkContext`
with sinon stubs for `telemetry`, `logger`, `ui`, and `auth`. Use it in every core-next test.

### Integration Tests

Integration tests verify cross-module interactions without external services:
- **core-next:** Operation pipeline → telemetry sequence, cross-cutting module composition
- **cli-next:** Commander program parsing, error class integration, command group registration

### Test Best Practices

- Use `sinon.stub()` / `sinon.sandbox` for external dependencies; always `sinon.restore()` in `afterEach`
- Test both `ok` and `err` paths for every `Result`-returning function
- Avoid testing implementation details — test behavior and outputs
- Mock I/O boundaries (file system, HTTP, Azure SDK clients), not internal logic
- Name tests as `should <expected behavior> when <condition>`

## Async Code

- Always `async`/`await` — no raw `.then()` chains
- No floating promises — every promise must be awaited or returned
- Use `Promise.all()` for independent concurrent operations, not sequential awaits
- Handle cleanup in `finally` blocks when acquiring resources (file handles, connections)

## Localization

- User-facing strings: `getLocalizedString("teamsfx.key")` from `common/localizeUtils`
- Default English strings via `getDefaultString("key")`
- Translation files in `Localize/loc/`
- Never concatenate translated strings — use parameterized keys: `getLocalizedString("key", param1, param2)`

## Security

- Never log secrets, tokens, or credentials — use `maskSecret()` from `common/stringUtils`
- Validate all external input at system boundaries (CLI args, API responses, file contents)
- Use `validator` library for URL/string validation; no hand-rolled regex for security checks
- Prefer `fs-extra` over raw `fs` for atomic file operations

### Filesystem Access — EAFP Pattern

Use **EAFP (Easier to Ask for Forgiveness than Permission)** for filesystem operations.
Do NOT check existence before operating — this causes TOCTOU race conditions (CodeQL `js/toctou-race-condition`).

```typescript
// BAD — TOCTOU race condition
if (await fs.access(filePath).then(() => true).catch(() => false)) {
  const data = await fs.readFile(filePath, "utf-8");
}

// GOOD — EAFP
try {
  const data = await fs.readFile(filePath, "utf-8");
} catch (e: any) {
  if (e.code === "ENOENT") {
    // Handle missing file
  } else {
    throw e;
  }
}
```

### Archive Security

- **Zip Slip**: When extracting archives, validate each entry name with `indexOf("..")` on the
  raw name AND check `path.resolve()` stays within the target directory.
- **ZIP validation**: Before uploading ZIP files to services, verify the magic bytes `PK\x03\x04`
  (`Buffer.from([0x50, 0x4b, 0x03, 0x04])`).

## Telemetry

- Every public API entry point should emit start/end telemetry events
- Include `correlationId` from `common/correlator` for distributed tracing
- Error telemetry: include `error.source`, `error.name`, `error.message` — never include PII
- Use `TelemetryReporter` interface for telemetry collection

## Feature Flags

- Use `featureFlagManager` from `common/featureFlags` to check flags
- Gate new behavior behind flags for safe rollout; keep the old code path functional
- Never hardcode flag values — always read from `featureFlagManager`

## Commits (Conventional Commits)

Format: `type(scope): subject`

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `revert`, `perf`, `ci`, `build`

```
feat(fx-core): add new generator for custom agents
fix(cli): resolve crash on missing config file
test(api): add contract tests for new interface
```

## Build Commands

```bash
npm run setup              # Full monorepo install + build
npm run setup:cli          # CLI + dependencies
npm run setup:vsc          # VS Code extension + dependencies
pnpm --filter @microsoft/teamsfx-core run build   # Single package
npm run watch              # Watch mode (all packages)

# v4 packages
cd packages/core-next && npm run build    # Build core-next
cd packages/core-next && npm run test:unit        # 492 unit tests
cd packages/core-next && npm run test:integration # 24 integration tests
cd packages/cli-next && npm run build     # Build cli-next (tsc only — dev)
cd packages/cli-next && npm run bundle    # Bundle with esbuild (dev, no minification)
cd packages/cli-next && npm run package   # Full production build (tsc + esbuild --production)
cd packages/cli-next && npm run test:unit         # 78 unit tests
cd packages/cli-next && npm run test:integration  # 62 integration tests
```

## Templates

- Organized by platform and language: `vsc/(ts|js|python)/`, `vs/csharp/`
- Dynamic files use `.tpl` suffix with Mustache syntax
- Every template includes: `appPackage/`, `infra/`, `m365agents.yml`, `env/`
- Build output is distributed into `packages/fx-core/templates`
- Register new templates in generator metadata and `generatorProvider.ts`

## Common Pitfalls

- **Forgetting to rebuild dependencies:** Change `api` → rebuild before testing `fx-core`
- **Floating promises:** ESLint catches these — fix, don't suppress
- **Wrong commit format:** commitlint rejects non-conventional messages
- **Raw error strings:** Always use `getLocalizedString()` for user-facing messages
- **Wrong error type:** `UserError` for input/config issues, `SystemError` for infra failures
- **Lock file conflicts:** Run `pnpm install` to regenerate — never edit lock files manually
