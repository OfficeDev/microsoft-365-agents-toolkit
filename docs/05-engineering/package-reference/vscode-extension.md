# `packages/vscode-extension` — `ms-teams-vscode-extension`

VS Code extension. UI + handlers; engine work delegated to `fx-core` (v3) / `core-next` (v4).

## Conventions source

[`.github/instructions/vscode-extension.instructions.md`](../../../.github/instructions/vscode-extension.instructions.md).

## Architecture layers

```
extension.ts (activate / deactivate)
  └─ CommandController                    # command registration + state
       └─ handlers/                       # domain logic; calls fx-core
            └─ core (FxCore from globalVariables)
  └─ TreeViewManager                      # sidebar tree views
  └─ ExtTelemetry                         # instrumentation
  └─ chat/                                # @teamsapp Copilot Chat participant
  └─ controls/                            # React webviews (Vite, Fluent UI v8)
```

## Command registration

```ts
registerInCommandController(
  context,
  CommandKeys.MyCommand,
  myCommandHandler,
  "myCommandRunningLabel" // optional localized blocking label
);
```

- `Correlator.run()` wraps every command with a correlation ID.
- Exclusive commands block concurrent execution; tree view shows tooltip.
- Add the command key to `package.json` contributions and `package.nls.json`.

## Handler shape

Always returns `Promise<Result<any, FxError>>`. Always:

1. Send start telemetry with `getTriggerFromProperty(args)`.
2. Build inputs from args / workspace state.
3. Delegate to fx-core.
4. On error: send error telemetry + `await showError(result.error)`.
5. On success: send success telemetry + return.

Never put business logic in handlers.

## Singletons

Lazy `getInstance()` pattern. Key: `VsCodeLogProvider`, `M365Login`, `CommandController`, `TreeViewManager`.

## Telemetry

```ts
ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MyFeature, {
  [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.TreeView,
}, { duration: elapsed });

ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MyFeature, error, extraProps);
```

Define event names in `src/telemetry/extTelemetryEvents.ts`. Auto-injected: `Component`, `IsExistingUser`, `IsSpfx`, `SettingsVersion`. Always pass `getTriggerFromProperty(args)`.

## Error display

`showError()` from `src/error/common.ts` — never `vscode.window.showErrorMessage()` directly. Adds context-aware action buttons (Debug in Playground, Troubleshoot with Agent), feature-flag-gated recommendations, sends `ShowError` telemetry.

## Tree views

`TreeViewManager.registerTreeViews()` — each view is a `TreeDataProvider<TreeViewCommand>`. `updateTreeViewsByContent()` refreshes based on project type and feature flags. `setRunningCommand()` / `restoreRunningCommand()` for exclusive command state.

## Webviews

- React + `@fluentui/react` v8 (NOT v9 `@fluentui/react-components`).
- Built with **Vite** (`vite.config.mts`); the main extension is bundled with esbuild (`esbuild.mjs`). The `package` script runs Vite first to build the webview bundles, then esbuild for the extension entry point. Both are part of the standard build.
- Routes via `react-router-dom` `MemoryRouter`.
- Localised via `react-intl` `IntlProvider`.

## Globals

`src/globalVariables.ts`:

- `core` — FxCore instance
- `tools` — the `Tools` instance (the type lives in `packages/api/src/utils/index.ts` as the `Tools` interface; v3 calls it `TOOLS` colloquially because it's the v3 ambient context)
- `workspaceUri` — current workspace URI
- `isSPFxProject`, `isDeclarativeCopilotApp` — project type flags

## v4 future

Migration gated by `TEAMSFX_V4_CORE`. Imports shift from `@microsoft/teamsfx-api` to `@microsoft/teamsfx-core` v4; `TOOLS` replaced with injected `AtkContext`. Handler shape stays the same.
