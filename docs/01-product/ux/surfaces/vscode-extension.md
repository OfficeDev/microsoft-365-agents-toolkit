# VS Code extension surface

Package: [`packages/vscode-extension`](../../05-engineering/package-reference/vscode-extension.md). Conventions: [`vscode-extension.instructions.md`](../../../.github/instructions/vscode-extension.instructions.md). **Full v3 command catalog:** [vscode-extension-commands.md](vscode-extension-commands.md).

## Layers

```
extension.ts (activate / deactivate)
  └─ CommandController             # registration + state
       └─ handlers/                # domain logic; calls fx-core
            └─ core (FxCore)
  └─ TreeViewManager               # sidebar UI
  └─ ExtTelemetry                  # instrumentation
  └─ chat/                         # @teamsapp Copilot Chat participant
  └─ controls/                     # React webviews (Vite)
```

## Command registration

Always go through `registerInCommandController(context, key, handler, runningLabel?)` in `extension.ts`. This:

- Wraps the handler in `Correlator.run(uuid, ...)` for tracing.
- Marks exclusive commands so the tree view can show "running…" tooltips.
- Centralises telemetry attribution.

## Handler shape

```typescript
async function myHandler(...args: any[]): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MyStart, getTriggerFromProperty(args));
  const inputs = getSystemInputs();
  const result = await core.myOperation(inputs);
  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.My, result.error);
    await showError(result.error);
    return result;
  }
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.My, { success: "yes" });
  return result;
}
```

Never put business logic in handlers — delegate to fx-core / core-next.

## Singleton providers

Lazy `getInstance()` pattern; export the instance as default. Examples: `VsCodeLogProvider`, `M365Login`, `CommandController`, `TreeViewManager`.

## Webviews

- Built with **Vite** (`vite.config.mts`); the main extension itself is bundled with esbuild (`esbuild.mjs`). The `package` script runs Vite first to build webview bundles, then esbuild for the extension entry point.
- `@fluentui/react` v8 (not v9 `@fluentui/react-components`).
- `react-router-dom` `MemoryRouter`.
- `react-intl` `IntlProvider` for localisation.

## v4 future

When `vscode-extension` migrates to `core-next` (gated by `TEAMSFX_V4_CORE`), the imports shift from `@microsoft/teamsfx-api` to `@microsoft/teamsfx-core` v4 and `TOOLS` is replaced with injected `AtkContext`. Handler shape stays the same.
