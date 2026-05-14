---
name: dev-workflow
description: "Development workflow for implementing features and fixing bugs in the Microsoft 365 Agents Toolkit monorepo. Use when: implementing a feature, fixing a bug, adding a test, modifying a package, creating a PR, running tests, building packages, scaffolding templates, lint errors, commit message format, code review checklist."
argument-hint: "Describe the feature or bug to implement/fix"
---

# Development Workflow

## When to Use

- Implementing a new feature in any package
- Fixing a bug across the monorepo
- Adding or modifying tests
- Creating or updating scaffolding templates
- Preparing a PR for review

## Monorepo Overview

This is a PNPM monorepo (Lerna + pnpm workspaces). Key packages and their roles:

| Package | Name | Role |
|---------|------|------|
| `packages/api` | `@microsoft/teamsfx-api` | Contracts and interfaces — change here affects everything |
| `packages/fx-core` | `@microsoft/teamsfx-core` | Core engine — generators, drivers, resources |
| `packages/cli` | `@microsoft/m365agentstoolkit-cli` | CLI tool |
| `packages/vscode-extension` | `ms-teams-vscode-extension` | VS Code extension |
| `packages/manifest` | Manifest utilities | App manifest handling |
| `packages/spec-parser` | `@microsoft/m365-spec-parser` | API spec parsing |
| `packages/mcp-server` | MCP server | MCP protocol support |
| `templates/` | Scaffolding templates | Project templates for all app types |

Dependency flow: `api` → `fx-core` → `cli` / `vscode-extension`

## Procedure

### Phase 1 — Understand Scope

1. Identify which package(s) the change affects
2. If touching `api`, expect downstream impact on `fx-core`, `cli`, and `vscode-extension`
3. If adding a new template or project type, see [template conventions](./references/templates.md)
4. Read the package-specific CONTRIBUTING.md if it exists

### Phase 2 — Setup & Build

```bash
# First-time setup
npm run setup                # pnpm install && build all

# Targeted setup (faster)
npm run setup:cli            # CLI and dependencies only
npm run setup:vsc            # VS Code extension and dependencies only

# Build specific package
pnpm --filter @microsoft/teamsfx-core run build

# Watch mode (all packages)
npm run watch
```

### Phase 3 — Implement the Change

Follow these coding standards:

1. **File header** — Every `.ts` source file must start with:
   ```typescript
   /**
    * Copyright (c) Microsoft Corporation.
    * Licensed under the MIT license.
    */
   ```

2. **Formatting** — Double quotes, semicolons, 2-space indent, 100 char print width. Run:
   ```bash
   cd packages/<pkg> && npm run format
   ```

3. **Async code** — All promises must be awaited or returned. No floating promises.

4. **No secrets** — ESLint `no-secrets` rule is active (tolerance 4.5). Never hardcode credentials.

5. **Unused variables** — Prefix intentionally unused parameters/variables with `_` (e.g., `_err`, `_unused`). ESLint enforces `argsIgnorePattern: "^_"`.

6. **Error handling** — Use `Result<T, FxError>` pattern. See [architecture reference](./references/architecture.md)
   - `UserError` for recoverable errors (bad input, missing config)
   - `SystemError` for unrecoverable errors (service failures)
   - Always use `getLocalizedString(key, ...params)` for user-facing messages

7. **Localization** — User-facing strings go in `package.nls.json` (default) with translations in `package.nls.{locale}.json`. Reference strings via `getLocalizedString("key")` or `getDefaultString("key")` from `common/localizeUtils`.

8. **Architecture patterns** — See [architecture reference](./references/architecture.md)

### Phase 4 — Test

```bash
# Run unit tests for a specific package
cd packages/fx-core && npm run test:unit

# Run all unit tests across the monorepo
pnpm run -r --stream test:unit

# Run targeted test suites (fx-core examples)
npm run test:core        # Core-specific
npm run test:component   # Component tests
npm run test:bot         # Bot feature tests

# CLI tests
cd packages/cli
npm run test:unit        # All CLI unit tests
npm run test:cmds        # Command tests
npm run test:engine      # Engine tests
```

- **Framework:** Mocha + Chai + Sinon
- **Coverage:** NYC — check with `npm run test:unit` (generates coverage report)
- Add tests for every new feature or bug fix. Match existing test patterns in the package.

**Debugging the VS Code extension:**
1. Open `packages/vscode-extension` in VS Code
2. Press F5 → launches "Run Extension" (Extension Host with dev env vars)
3. Use "Extension Unit Tests" launch config to debug tests with breakpoints
4. Watch mode: `npm run watch` (runs esbuild + tsc + vite in parallel)

Key env vars set during debug: `NODE_ENV=development`, `TEMPLATE_VERSION=local`

**E2E tests** (`packages/tests/`):
- Requires M365 + Azure accounts, `.env` with credentials
- Uses vscode-extension-tester + Playwright for UI verification
- See `packages/tests/README.md` for full setup

**E2E tests — next-gen** (`packages/cli-next/tests/e2e/`):
- For core-next / cli-next changes, see the `dev-test-next` skill
- Login locally with `atk auth login azure` + `atk auth login m365`
- Run: `cd packages/cli-next && npm run test:e2e:cli` (no creds) or `npm run test:e2e:lifecycle` (with creds)

### Phase 5 — Lint & Format

> **For core-next and cli-next packages**, see the **`lint-format` skill** for the full pipeline, gotchas, and architecture.

For core-next and cli-next, `build` auto-runs `eslint --fix` via `postbuild` hook (formatting included via `eslint-plugin-prettier`).
For other packages, run manually:

```bash
cd packages/<pkg>
npm run lint              # Check lint errors
npm run lint:fix          # Auto-fix lint errors
npm run check-format      # Check formatting
npm run format            # Auto-format
npm run check-sensitive   # Detect hardcoded secrets
```

All of these run automatically on pre-commit via Husky + lint-staged.

### Phase 6 — Commit & PR

**Commit message format** (Conventional Commits, enforced by commitlint):
```
type(scope): subject
```

Valid types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `revert`, `perf`, `ci`, `build`

Examples:
```
feat(fx-core): add new generator for custom agents
fix(cli): resolve crash on missing config file
test(api): add contract tests for new interface
docs(vscode): update README with new commands
```

**PR checklist:**
- [ ] Branch from latest `dev`
- [ ] All lint and format checks pass
- [ ] Unit tests pass locally
- [ ] New tests added for the change
- [ ] Commit messages follow conventional format
- [ ] CODEOWNERS will auto-assign reviewers

## Decision Points

### Where does my change go?

| Change Type | Package |
|-------------|---------|
| New interface / contract | `packages/api` |
| New command or generator | `packages/fx-core` |
| CLI command handling | `packages/cli` |
| VS Code UI / commands | `packages/vscode-extension` |
| New project template | `templates/` + metadata in `fx-core` |
| API spec parsing | `packages/spec-parser` |
| App manifest changes | `packages/manifest` |

### Adding a new template?

See [template conventions](./references/templates.md) for the full process including template files, metadata registration, and generator setup.

### Modifying a generator?

Generators live in `packages/fx-core/src/component/generator/`. Each extends `DefaultTemplateGenerator` with `activate()`, `getTemplateNames()`, and `post()` methods. Register new generators in `generatorProvider.ts`.

## Common Pitfalls

- **Forgetting to build dependencies:** If you change `api`, rebuild before testing `fx-core`
- **Floating promises:** ESLint will catch `await`/`return` violations — fix them, don't suppress
- **Wrong commit format:** Commitlint rejects non-conventional messages at the hook
- **Lock file conflicts:** Use `pnpm install` to regenerate; don't manually edit lock files
- **Template placeholders:** Dynamic template files use `.tpl` suffix — don't forget it
- **Missing localization:** User-facing error messages must use `getLocalizedString()`, not raw strings
- **Wrong error type:** Use `UserError` for input/config issues, `SystemError` for internal failures
