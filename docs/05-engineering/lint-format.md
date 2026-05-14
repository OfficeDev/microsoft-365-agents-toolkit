# Lint & Format

> **Workflow guide:** For fix commands, gotchas, and precommit troubleshooting, see the [`lint-format` skill](../../.github/skills/lint-format/SKILL.md).

## Architecture — Single Source of Truth

All formatting and linting is handled by **ESLint with `eslint-plugin-prettier`**. There is no separate Prettier check in CI. This eliminates conflicts between standalone Prettier and the ESLint Prettier plugin.

```
eslint-plugin-prettier (in shared ESLint config)
       │
       ├── precommit: lint-staged → eslint --fix --quiet
       ├── postbuild: eslint --fix
       └── CI lint:   eslint (no --fix, errors block merge)
```

### Why Not Standalone Prettier?

**Prettier v3 and `eslint-plugin-prettier` can disagree.** Example: Prettier v3 wraps `??` in parens inside ternaries (`? (x ?? y)`), but `eslint-plugin-prettier` says remove them. If both tools run, the last one wins — and CI may use a different order or run them independently, causing spurious failures.

**Rule:** Use `eslint --fix` for all formatting. The `format` / `format:check` scripts exist for manual use but are **not** part of CI or precommit.

## Config Chain

| File | Location | Role |
|------|----------|------|
| `shared.mjs` | `packages/eslint-plugin-teamsfx/config/` | Base ESLint config — imports `eslint-plugin-prettier/recommended` |
| `eslint.config.mjs` | Each package root | Extends `shared.mjs` with package-specific rules |
| `.prettierrc.js` | Each package root | Extends `packages/prettier-config/index.js` |

The shared config at `shared.mjs` imports:
- `typescript-eslint` recommended rules
- `eslint-plugin-prettier/recommended` (formats via ESLint)
- `eslint-plugin-import-x` (import ordering)
- `eslint-plugin-no-secrets` (credential detection)

### Prettier Config

Defined in `packages/prettier-config/index.js`:
- Double quotes, semicolons, trailing commas (es5)
- 2-space indent, 100 char print width
- LF line endings, arrow parens always

## Precommit Pipeline

```
git commit
  → .husky/pre-commit
    → npx lerna run --concurrency 1 --stream precommit --since HEAD --exclude-dependents
      → per-package "precommit" script
        → lint-staged
          → eslint --fix --quiet (on staged *.{js,jsx,css,ts,tsx} files)
```

### Requirements for a Package to Participate

A package must have ALL of these in `package.json`:

1. **`precommit` script:** `"precommit": "lint-staged"`
2. **`lint-staged` config section:**
   ```json
   "lint-staged": {
     "*.{js,jsx,css,ts,tsx}": ["npx eslint --fix --quiet"]
   }
   ```
3. **`lint-staged` in devDependencies:** `"lint-staged": "^10.5.4"`
4. **`eslint-plugin-prettier` in devDependencies:** `"eslint-plugin-prettier": "^5.5.0"`
5. **`eslint-config-prettier` in devDependencies:** `"eslint-config-prettier": "^10.1.0"`

If any are missing, lerna silently skips the package and lint issues slip through to CI.

### No `--cache` Flag

lint-staged runs `eslint --fix --quiet` **without** `--cache`. CI runs on a fresh checkout with no cache. Using `--cache` locally can hide issues that CI catches (e.g., after config changes).

## CI Pipeline (`ci-next.yml`)

```
build job:
  → pnpm build (triggers postbuild: eslint --fix)

lint job (needs build):
  → pnpm build (triggers postbuild: eslint --fix)
  → pnpm lint  (eslint without --fix — errors block merge)
```

There is **no** `format-check` job. The `lint` job handles formatting via `eslint-plugin-prettier`.

## Postbuild Script

Both `core-next` and `cli-next` run `eslint --fix` as a postbuild hook:

```json
"postbuild": "eslint --fix \"src/**/*.ts\" \"tests/**/*.ts\""
```

This ensures source files are formatted after every build. Do **not** add `prettier --write` here — it conflicts with `eslint-plugin-prettier`.
