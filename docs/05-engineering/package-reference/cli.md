# `packages/cli` — `@microsoft/m365agentstoolkit-cli` (v3)

The **v3 CLI** — `atk` binary. Wraps `fx-core` for command-line use.

## Conventions source

[`.github/instructions/cli.instructions.md`](../../../.github/instructions/cli.instructions.md).

## Architecture

```
cli.js (entry)
  → activate() in index.ts
    → CLIEngine.start(rootCommand)
      → findCommand → parseArgs → validateOptions → handler()
```

## Command model

Declarative `CLICommand` objects in `src/commands/models/`:

```ts
export const myCommand: CLICommand = {
  name: "my-command",
  description: "...",
  options: [...MyCommandOptions],
  handler: async (ctx) => {
    const result = await ctx.core.myOperation(ctx.inputs);
    return result;
  },
  telemetry: { event: TelemetryEvent.MyCommand },
};
```

Adding a command: `src/commands/models/<name>.ts` → export from `index.ts` → add to `commands` array in `root.ts` → localised strings in `src/resource/`.

## Engine — 7 phases

1. Find command (tree traversal + alias).
2. Parse args (custom parser: `--key=value`, `--key value`, `-k value`).
3. Version check (`--version`).
4. Help display (`--help`).
5. Validate options (skipped in interactive mode).
6. Version compatibility (project version + phantom migration).
7. Execute handler — wrapped in `Correlator.run()`.

## Interactive vs non-interactive

| Mode | Trigger | Behaviour |
|------|---------|-----------|
| Interactive | TTY | CLI options discarded; question model drives prompts |
| Non-interactive | `CI_ENABLED=true` or no TTY | Options must be passed as args; missing → `MissingRequiredOptionError` |

`reservedOptionNamesInInteractiveMode` is the small set always parsed.

## User interaction

`CLIUserInteraction` wraps `@inquirer/prompts`:

| Method | Prompt |
|--------|--------|
| `singleSelect` | Custom list (`customizedListPrompt.ts`) |
| `multiSelect` | Custom checkbox (`customizedCheckboxPrompt.ts`) |
| `input` | `@inquirer/prompts.input` |
| `password` | `@inquirer/prompts.password` |
| `confirm` | `@inquirer/prompts.confirm` |

`ScreenManager` pauses/resumes terminal output around prompts.

## Error classes

All extend `UserError`:

- `MissingRequiredOptionError`, `MissingRequiredArgumentError`
- `InvalidChoiceError`
- `UnknownCommandError` (with edit-distance "did you mean?")
- `UnknownOptionError`, `ArgumentConflictError`

## Telemetry

```ts
CliTelemetry.sendTelemetryEvent(eventName, properties, measurements);
CliTelemetry.sendTelemetryErrorEvent(eventName, error, properties, measurements);
```

Auto-injected: `component`, `commandName`, `projectId`, `correlationId`, `runFrom` (CI detection).

## Login

Singleton providers in `src/commonlib/`:

- `AzureLogin` — MSAL interactive code flow.
- `M365Login` — MSAL with `CryptoCachePlugin` for AES-256-GCM cache.
- Windows: native broker plugin for WAM.

## Output

- `CLILogger` implements `LogProvider` with chalk colour.
- `ScreenManager.writeLine()`.
- `colorize.ts`: red=error, green=success, cyan=links.
- `Progress`: `start()` / `end(success)`.

## Bundling

Webpack — needs `NODE_OPTIONS=--max-old-space-size=4096`. Replaced in v4 by esbuild ([`cli-next`](cli-next.md)).
