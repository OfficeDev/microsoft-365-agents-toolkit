# Plan: Wire 5 DA Quick-Win CLI Commands

**Status:** Complete
**Created:** 2026-04-07
**Updated:** 2026-04-07

## Goal

Wire 5 stub CLI commands to their existing core-next DA operations, moving coverage from 14/39 (36%) to 19/39 (49%). Flags-only (no interactive prompts). Drop `add spfx-web-part`.

## Tasks

- [x] Phase 1: Add `getAgentManifestPath(projectPath)` to core-next `declarativeAgent/manifest/`
- [x] Phase 2: Create `extendToM365Op` operation wrapping the existing driver
- [x] Phase 3: Create 5 CLI action handlers (addAction, addCapability, addAuthConfig, setSensitivityLabel, m365Sideload)
- [x] Phase 4: Rewrite command definitions (add.ts, misc.ts, m365.ts) to use `wrapHandlerWithContext`
- [x] Phase 5: Core-next unit tests (getAgentManifestPath, extendToM365Op)
- [x] Phase 6: CLI-next unit tests (5 action handlers)
- [x] Phase 7: Build, lint, and full test verification

## Notes

- `add spfx-web-part` dropped (no core-next backing)
- `addMCPActionOp` NOT wired — too many required fields for flags-only
  - **Update:** Wired in Gap 1 (session 2026-04-15). `add action --api-plugin-type mcp` routes to `addMCPActionAction()` with full MCP options.
- `--agent-manifest-path` auto-discovered from project when not provided
- All flags-only, CI-friendly (no interactive prompts)

## Log

- 2026-04-07 — Plan created
- 2026-04-07 — All 7 phases complete. 17 new tests (6 core-next + 11 cli-next), all passing. Totals: 492 core-next unit, 78 cli-next unit. 5 commands wired, `add spfx-web-part` removed.
