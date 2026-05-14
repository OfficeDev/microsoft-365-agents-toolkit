# Versioning

## Semver per package

Each npm package has its own semver:

- **MAJOR** — breaking change. Required for any contract change in `api` / `manifest`.
- **MINOR** — new feature, backward-compatible.
- **PATCH** — bug fix or doc-only change.

`api` and `manifest` are the most conservative — every breaking change must justify the blast radius across all downstream packages.

## v3 vs v4 line

| Engine | npm name | Line |
|--------|----------|------|
| v3 | `@microsoft/teamsfx-core` (from `packages/fx-core`) | 3.x |
| v4 | `@microsoft/teamsfx-core` (from `packages/core-next`, dist tag preview) | 4.x |

The npm package name collision is intentional — when v4 reaches GA, `latest` flips from 3.x to 4.x. Until then, v3 is `latest`, v4 is `next` (or `preview`).

## VS Code extension

Independent versioning from npm packages. Convention: `M.S.B` where `M.S` mirrors marketplace channel and `B` is build number.

## CHANGELOG

Per-package `CHANGELOG.md` (where present) plus the rolled-up [`CHANGELOG.md`](../../CHANGELOG.md) at the repo root. Generated from Conventional Commit subjects.

## Conventional Commits → version bump

| Commit type | Bump |
|-------------|------|
| `fix:` | PATCH |
| `feat:` | MINOR |
| `feat!:` or `BREAKING CHANGE:` footer | MAJOR |
| `docs:`, `style:`, `refactor:`, `test:`, `ci:`, `build:`, `chore:` | none |

Lerna's `--conventional-commits` flag automates the bump computation. Manual override available for tricky cases.

## Pre-release tags

| Tag | When |
|-----|------|
| `-alpha.N` | Internal preview |
| `-beta.N` | Public preview |
| `-rc.N` | Release candidate |
| `-preview.N` | Long-running preview track (used for v4) |

All pre-release tags are published with `--tag <name>` (never `latest`) so users on `latest` are unaffected.
