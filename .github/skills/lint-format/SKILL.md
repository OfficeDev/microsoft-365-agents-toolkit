---
name: lint-format
description: "Lint and format pipeline for core-next and cli-next packages. Use when: fixing lint errors, fixing format errors, prettier conflicts, eslint-plugin-prettier disagreements, precommit hook issues, CI lint failures, postbuild formatting, configuring lint-staged."
argument-hint: "Describe the lint/format issue or which package is failing"
---

# Lint & Format Pipeline

For the architecture (config chain, why not standalone Prettier, precommit/CI/postbuild pipelines), see **[`docs/05-engineering/lint-format.md`](../../docs/05-engineering/lint-format.md)**.

## Common Gotchas

### 1. Standalone Prettier Disagrees with ESLint Plugin

**Symptom:** `prettier --check` fails but `eslint` passes (or vice versa).

**Cause:** `prettier` v3 and `eslint-plugin-prettier` can produce different output for edge cases (e.g., nullish coalescing in ternaries).

**Fix:** Only use `eslint --fix` for formatting. Never run `prettier --write` before `eslint` in CI or postbuild.

### 2. VS Code Editor Buffer Overwrites Terminal Fixes

**Symptom:** `eslint --fix` in terminal fixes a file, but the fix disappears or isn't staged.

**Cause:** VS Code has the file open with a stale buffer and saves over the terminal's changes.

**Fix:** After running `eslint --fix` in terminal, either:
- Close and reopen the file in VS Code, or
- Use the VS Code editor tool (replace_string_in_file) to make the edit, or
- Run `eslint --fix` via the postbuild script (`npm run build`)

### 3. New Package Not Linted on Precommit

**Symptom:** Lint errors pass locally but fail in CI.

**Cause:** The package is missing the `precommit` script, `lint-staged` config, or dependencies.

**Fix:** Add all 5 requirements listed in [`docs/05-engineering/lint-format.md` §"Requirements for a Package to Participate"](../../docs/05-engineering/lint-format.md#requirements-for-a-package-to-participate).

### 4. Stale ESLint Cache Hides Errors

**Symptom:** `eslint --cache` passes locally but CI fails.

**Cause:** ESLint cache marks a file as clean, but a config change (e.g., new rule in `shared.mjs`) means it should be re-checked.

**Fix:** Run `eslint --no-cache` to verify, or delete `.eslintcache`. The lint-staged config intentionally omits `--cache`.

## Quick Reference Commands

```bash
# Check lint (same as CI)
cd packages/core-next && npx eslint --no-cache "src/**/*.ts" "tests/**/*.ts"

# Auto-fix lint + formatting
cd packages/core-next && npx eslint --fix --no-cache "src/**/*.ts" "tests/**/*.ts"

# Manual prettier format (not used in CI)
cd packages/core-next && npm run format

# Simulate full CI flow locally
cd packages/core-next && npm run build && npm run lint
```
