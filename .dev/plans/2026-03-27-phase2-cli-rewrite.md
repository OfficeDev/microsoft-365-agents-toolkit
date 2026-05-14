# Plan: Phase 2 — CLI Rewrite (`packages/cli-next/`)

**Status:** Complete
**Created:** 2026-03-27
**Updated:** 2026-03-30

## Goal

Rebuild the CLI from scratch using Commander.js, consuming `@microsoft/teamsfx-core` v4 operations directly. Published under `@microsoft/m365agentstoolkit-cli` v4.0.0 with binary names `atk`, `teamsapp`, `m365agentstoolkit-cli`.

## Tasks

- [x] Scaffold `packages/cli-next/` with package.json, tsconfig.json
- [x] Create entry points: `cli.js`, `cliold.js` (bin stubs)
- [x] Create `src/program.ts` — Commander.js program setup
- [x] Create `src/ui/` — interactive prompt adapter (CLIUserInteraction)
- [x] Create `src/output/` — formatting system (colorize, table, JSON)
- [x] Create `src/telemetry/` — telemetry reporter
- [x] Create `src/error.ts` — CLI error types
- [x] Create project lifecycle commands: new, provision, deploy, publish, preview
- [x] Create account commands: auth login azure/m365, auth logout, auth show
- [x] Create environment commands: env add, env list, env reset
- [x] Create teamsapp commands: validate, package, publish, update, doctor
- [x] Create add commands: add action, add capability, add auth-config, add spfx-web-part
- [x] Create utility commands: list, list samples, list templates, upgrade, set, validate, share
- [x] Create m365 commands: m365-sideload, m365-unacquire, m365-launch-info
- [x] Create permission commands: permission grant, permission status
- [x] Create entra-app commands: entra-app update
- [x] Create regenerate commands: regenerate, regenerate action
- [x] Create `src/index.ts` entry point
- [x] Register `packages/cli-next` in pnpm-workspace.yaml
- [x] Build and verify compilation
- [x] Set up cli-next test infrastructure (.mocharc.js, .nycrc, test scripts, devDependencies)
- [x] Write cli-next unit tests (19 tests across 4 files)
- [x] Write cli-next integration tests (4 tests across 1 file)

## Notes

- Commander.js replaces custom recursive-descent parser
- @inquirer/prompts kept for interactive mode
- Same 3 binary entry points as current CLI
- All commands are thin wrappers calling core-next operations
- Non-interactive via `--non-interactive` flag or `CI_ENABLED=true`
- JSON output via `--output json`

## Log

- 2026-03-27 — Plan created, starting implementation
- 2026-03-27 — Phase 2 complete: all commands registered, build passes with zero errors (92 output files)
- 2026-03-30 — Test infrastructure added: 19 unit tests + 4 integration tests, all passing
