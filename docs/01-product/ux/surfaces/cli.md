# CLI surface

Full v3 command reference (every flag, every example): [cli-v3-command-reference.md](cli-v3-command-reference.md).

Two CLIs are shipped under the same `atk` binary name:

| | v3 | v4 |
|-|----|----|
| Package | [`packages/cli`](../../05-engineering/package-reference/cli.md) | [`packages/cli-next`](../../05-engineering/package-reference/cli-next.md) |
| Bundler | webpack (4 GB heap) | esbuild (single CJS file) |
| Argument parser | Custom recursive descent | Commander.js |
| Auth | `commonlib/` MSAL | `auth/` MSAL (ported, refactored) |
| Engine | `fx-core` | `core-next` |
| Telemetry transport | `applicationinsights` (eager) | `applicationinsights` (lazy) |
| Status | Default until `TEAMSFX_V4_CORE=true` | Preview, opt-in |

Both implement the same surface vocabulary (`new`, `provision`, `deploy`, `publish`, `account`, `env`, `teamsapp`, ...).

## v3 command engine — 7 phases

1. **Find command** — tree traversal with alias support.
2. **Parse args** — `--key=value`, `--key value`, `-k value`.
3. **Version check** — `--version`.
4. **Help display** — `--help`.
5. **Validate options** — skipped in interactive mode (reserved options only).
6. **Version compatibility** — project-version check + phantom migration.
7. **Execute handler** — wrapped in `Correlator.run()`.

See [`cli.instructions.md`](../../../.github/instructions/cli.instructions.md).

## v4 command engine — Commander + factory

```
buildProgram()
  └─ command groups (project, account, env, teamsapp, add, list, ...)
       └─ wrapHandler(name, handler)              # telemetry + error handling
       └─ wrapHandlerWithContext(name, handler)   # also creates AtkContext + delegates to core-next
            └─ registerBuiltinDrivers()           # lazy — only when a real command runs
```

The `new` subtree is **generated from `TemplateRegistry`** by `buildNewCommands()` — adding a `TemplateDescriptor` adds a CLI subcommand. Question metadata (`cliName`, `cliShortName`, `cliDescription`, `isBoolean`) is mapped to Commander options via `mapQuestionToOption()`.

## Interactive vs non-interactive

| Mode | Trigger | Behaviour |
|------|---------|-----------|
| Interactive | TTY + no `CI_ENABLED` | CLI options are parsed but **discarded**; the question model drives prompts via `CLIUserInteraction` |
| Non-interactive | `CI_ENABLED=true` or no TTY | All options must be passed as CLI args; missing required values fail with `MissingRequiredOptionError` |

`reservedOptionNamesInInteractiveMode` is the small set always parsed even in interactive mode (e.g. `--folder`, `--debug`).

## CLI error classes

All extend `UserError` from `@microsoft/teamsfx-api`:

- `MissingRequiredOptionError`
- `MissingRequiredArgumentError`
- `InvalidChoiceError`
- `UnknownCommandError` (with edit-distance "did you mean?")
- `UnknownOptionError`
- `ArgumentConflictError`

Strings localised via `src/resource/`.

## Login flows

Singleton providers in `src/commonlib/` (v3) / `src/auth/` (v4):

- `AzureLogin` — MSAL interactive code flow.
- `M365Login` — MSAL with `CryptoCachePlugin` (AES-256-GCM cache).
- Windows: native broker plugin for WAM integration when present.
