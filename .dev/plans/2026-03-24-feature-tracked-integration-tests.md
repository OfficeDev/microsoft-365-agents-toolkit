# Plan: Feature-Tracked Integration Test Framework

**Status:** In Progress
**Created:** 2026-03-24
**Updated:** 2026-03-24

## Goal

Build a registry-driven integration test framework with a single source of truth (`featureRegistry.ts`) that serves both as AI-readable feature documentation and test driver across all four entry points (fx-core, CLI, Server, VS Code extension). Produces JSON + HTML coverage matrix reports.

## Tasks

- [ ] Create featureRegistry.ts with types, enums, and curated features
- [ ] Create coverageTracker.ts for runtime pass/fail tracking
- [ ] Create reportGenerator.ts for JSON + HTML output
- [ ] Create rootHooks.ts (Mocha afterAll wiring)
- [ ] Create testBuilders.ts (scaffoldTest, provisionTest factories)
- [ ] Rewrite createProject.test.ts to registry-driven loop
- [ ] Rewrite provision.test.ts to builder pattern
- [ ] Create .mocharc.integration.js + update package.json
- [ ] Create features.instructions.md + update fx-core.instructions.md
- [ ] Build and verify all tests pass

## Notes

- Only 6 of 12 originally planned templates have local directories available
- Templates with local dirs: default-bot, basic-tab, basic-custom-engine-agent, custom-copilot-basic, graph-connector, weather-agent
- message-extension-v2, teams-collaborator-agent, foundry-agent-to-m365 also available
- CLI/Server/VSCode integration tests deferred to follow-up (Phase 4-6 of design)

## Log

- 2026-03-24 — Plan created, starting Phase 1 implementation
