# Dev setup

## Workspace commands

```bash
pnpm install                                     # install all workspaces
pnpm --filter <name|path> run <script>           # run a script in one package
npm run watch                                    # watch mode (all packages)
```

Use directory paths (`./packages/core-next`) when filtering v4 packages — their npm names collide with v3.

## Per-package scripts

Most packages expose:

| Script | Effect |
|--------|--------|
| `build` | tsc (and postbuild for v4) |
| `lint` | ESLint check (0 errors required) |
| `format` | Prettier auto-format |
| `format:check` | Prettier dry-run (CI gate) |
| `test:unit` | Mocha unit tests with NYC |
| `test:integration` | Integration tests (no coverage) |
| `test` | alias for `test:unit` |

cli-next adds:

| Script | Effect |
|--------|--------|
| `bundle` | esbuild dev bundle |
| `package` | esbuild production bundle |

## Editor setup

- VS Code with TypeScript ~6.0 — workspace `tsconfig` will be picked up.
- Recommended extensions: ESLint, Prettier, Mocha Test Explorer.
- `eslint.format.enable: false` recommended; let Prettier own formatting.

## Pre-commit

Husky + lint-staged enforce:

- ESLint --fix on staged `.ts` files.
- Prettier --write on staged files.
- commitlint on the commit message.

If the pre-commit hook misbehaves, `git commit --no-verify` is **not** the right escape hatch — diagnose and fix instead. See [`lint-format`](../../.github/skills/lint-format/SKILL.md) skill.

## Running a single test

```bash
cd packages/core-next
npx mocha --require ts-node/register tests/unit/lifecycle/operations.test.ts
```

Or with VS Code's Mocha Test Explorer.

## Debugging

- VS Code: `Run and Debug` panel → "Mocha tests" preset (per package).
- For cli-next: launch `node --inspect-brk packages/cli-next/build/index.js <args>`.

## Building user apps locally

The toolkit is what you're building, but you'll often want to verify against a scaffolded user project. Quickest path:

```bash
cd packages/cli-next
npm run package
node build/index.js new                    # interactive scaffold
cd /tmp/my-test-project
node ../../packages/cli-next/build/index.js provision --env local
```

(Adjust paths to your clone.)
