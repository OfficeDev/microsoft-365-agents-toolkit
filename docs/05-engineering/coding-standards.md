# Coding standards

The authoritative source is [`.github/instructions/codebase.instructions.md`](../../.github/instructions/codebase.instructions.md). This page summarises and pointers to the rest.

## Type safety

- `strict: true` everywhere.
- Prefer `unknown` over `any`; narrow with type guards.
- Use discriminated unions over type assertions.
- Avoid `as` casts — use predicates or `satisfies`.
- Export interfaces from `api` (v3) or `core-next/src/api/` (v4); concrete types stay internal.

## Formatting

Prettier shared config in `packages/prettier-config`:

- Double quotes, semicolons, 2-space indent.
- 100-char print width, LF line endings.

## Linting

ESLint flat config (`eslint.config.mjs`) extending `eslint-plugin-teamsfx/config/shared.mjs` plus `header.mjs`.

| Package | Rules |
|---------|-------|
| `fx-core` | shared + header + promise + type |
| `cli` | shared + header + promise |
| `core-next` | shared + header + promise (`import-x/no-unresolved: off`) |
| `cli-next` | shared + header (`import-x/no-unresolved: off`) |

Key rules:

- License header on every `.ts` file.
- No floating promises (`promise/no-floating-promises`).
- No hardcoded secrets (`no-secrets`, entropy 4.5).
- No import cycles.
- Unused vars must be prefixed with `_`.

`import-x/no-unresolved` is off in v4 packages because the Node resolver cannot follow pnpm `workspace:*` symlinks; `tsc` catches real import errors.

## Naming

| Element | Style | Examples |
|---------|-------|----------|
| Classes / interfaces | PascalCase | `FxCore`, `UserError`, `LogProvider` |
| Functions / methods | camelCase | `loadFromPath`, `askSubscription` |
| Constants | SCREAMING_SNAKE_CASE | `AppStudioScopes`, `SharePointAppId` |
| Enums | PascalCase + PascalCase members | `FeatureFlagName.MultiEnv` |
| Files | camelCase or PascalCase matching main export | `FxCore.ts`, `featureFlags.ts` |
| Test files | `<sourceFile>.test.ts` | `FxCore.test.ts` |

## Imports

- Relative paths only — no path aliases.
- Order: external → workspace → relative.
- `workspace:*` references in `package.json`.
- `import type { ... }` for type-only imports.

```ts
import type { FxError } from "@microsoft/teamsfx-api";
import { ok, err } from "neverthrow";
import { featureFlagManager } from "../common/featureFlags";
```

## Async

- Always `async`/`await`.
- Every promise awaited or returned (lint enforces).
- `Promise.all` for concurrency.
- Resource cleanup in `finally`.

## Localisation

`getLocalizedString("teamsfx.key", ...params)` (v3) / `Localizer.getString(key, ...params)` (v4) for user-facing. `getDefaultString(...)` for logs. Never concatenate.

## Filesystem — EAFP

```ts
// BAD — TOCTOU
if (await exists(path)) await fs.readFile(path);

// GOOD
try {
  const data = await fs.readFile(path);
} catch (e: any) {
  if (e.code !== "ENOENT") throw e;
}
```

## Commits

Conventional Commits, enforced by commitlint:

```
feat(fx-core): add new generator for custom agents
fix(cli): resolve crash on missing config file
test(api): add contract tests for new interface
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `revert`, `perf`, `ci`, `build`.
