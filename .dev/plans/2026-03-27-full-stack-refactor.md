# Plan: Full-Stack Refactoring — Phase 1 Foundation

**Status:** Complete
**Created:** 2026-03-27
**Updated:** 2026-03-30

## Goal

Implement Phase 1 of the refactoring plan: create `packages/core-next/` with the new `@microsoft/teamsfx-core` v4.0.0 foundation — merged API contracts, AtkContext, Operation pipeline, TemplateDescriptor, DriverDescriptor, and AtkError model.

## Tasks

- [x] Scaffold `packages/core-next/` with package.json, tsconfig.json
- [x] Create `src/api/` — merge contracts from `@microsoft/teamsfx-api`
  - [x] Error types (FxError, UserError, SystemError)
  - [x] Types (Inputs, OptionItem, etc.)
  - [x] Context interface
  - [x] Question model (BaseQuestion, UI configs)
  - [x] User interaction (UserInteraction)
  - [x] Utilities (LogProvider, TelemetryReporter, TokenProvider, CryptoProvider)
  - [x] Constants (Platform, Stage, etc.)
  - [x] CLI types (CLICommand, CLIContext)
  - [x] Generator interface (IGenerator)
- [x] Define AtkContext interface (replaces TOOLS singleton)
- [x] Define Operation interface + runOperation wrapper
- [x] Define TemplateDescriptor + TemplateRegistry
- [x] Define DriverDescriptor + DriverRegistry
- [x] Define AtkError model
- [x] Create package index.ts with all exports
- [x] Register `packages/core-next` in pnpm-workspace.yaml
- [x] Add `TEAMSFX_V4_CORE` feature flag in fx-core
- [x] Build and verify compilation
- [x] Set up core-next test infrastructure (.mocharc.js, .nycrc, test scripts)
- [x] Write core-next unit tests (59 tests across 11 files)
- [x] Write core-next integration tests (11 tests across 2 files)

## Notes

- Existing npm name: `@microsoft/teamsfx-core`, bumped to v4.0.0
- New folder: `packages/core-next/` (coexists with `packages/fx-core/`)
- `@microsoft/teamsfx-api` contracts absorbed into `src/api/`
- Feature flag gates any runtime switch from old to new

## Log

- 2026-03-27 — Plan created, starting implementation
- 2026-03-27 — Phase 1 complete: all tasks done, build passes with zero errors
- 2026-03-30 — Test infrastructure added: 59 unit tests + 11 integration tests, all passing
