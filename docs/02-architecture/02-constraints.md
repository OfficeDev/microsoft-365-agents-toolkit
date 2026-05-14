# 2 — Constraints

## Technical constraints

| Constraint | Why | Source |
|-----------|-----|--------|
| TypeScript ~6.0 in v4 packages, strict mode everywhere | Type safety, modern bundler resolution | `packages/{core-next,cli-next}/tsconfig.json` |
| Node.js 18+ runtime | esbuild target, MSAL native plugins | `esbuild.mjs` (`target: "node18"`) |
| pnpm + Lerna monorepo | Workspace deps and per-package publish | `pnpm-workspace.yaml`, `lerna.json` |
| `Result<T, FxError>` everywhere | Errors are values, not exceptions | `neverthrow` re-exported from `@microsoft/teamsfx-api` |
| No `console.log`, no raw `Error.throw` for expected failures | Telemetry, localisation, masking | [`codebase.instructions.md`](../../.github/instructions/codebase.instructions.md) |
| EAFP filesystem access | Avoids TOCTOU (CodeQL `js/toctou-race-condition`) | Same |
| Mocha + Chai + Sinon + NYC | Single test stack across packages | `.mocharc.js`, `.nycrc` per package |

## Organisational constraints

| Constraint | Implication |
|-----------|-------------|
| Microsoft Open Source CoC + CLA | Every PR signs CLA; CoC enforced |
| Conventional Commits | `feat(scope):`, `fix(scope):` etc. enforced by commitlint |
| ESLint flat config + Prettier shared config | Centralised rules in `packages/eslint-plugin-teamsfx`, `packages/prettier-config` |
| 80% NYC coverage gate (v4) | Enforced in `ci-next.yml` |
| License header on every `.ts` source file | Enforced by ESLint header rule |

## Compatibility constraints

| Constraint | Where enforced |
|-----------|----------------|
| v3 project layout (`m365agents.yml`, `appPackage/`, `env/`) must continue to work in v4 | Driver IDs match v3 names; envMap injection mimics v3 placeholder resolution |
| Token cache at `~/.fx/account/` shared between v3 and v4 | `cli-next/src/auth/cacheAccess.ts` |
| `@microsoft/teamsfx-api` re-exports stay stable | [`api.instructions.md`](../../.github/instructions/api.instructions.md) |
| Manifest schema 1.x → 2.4 migration is non-destructive | `@microsoft/app-manifest` `TeamsManifestWrapper` |

## Platform constraints

We are downstream of:

- Microsoft 365 manifest schemas (Teams, M365 Copilot, Office Add-ins)
- Teams Developer Portal (TDP) REST API
- Microsoft Graph (Entra ID app reg, app catalog publish)
- Bot Framework (channel registration via ARM)
- Azure Resource Manager (ARM/Bicep deployments, Kudu zip deploy)
- M365 PackageService (sideloading)

When upstream changes, we either adapt drivers/clients or version-gate behaviour.
