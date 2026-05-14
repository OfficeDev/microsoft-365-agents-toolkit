---
description: "Use when editing or creating code in the VS Code extension: command registration, handlers, tree views, telemetry, error display, React webviews, singleton providers, or Copilot chat integration."
applyTo: "packages/vscode-extension/**/*.ts, packages/vscode-extension/**/*.tsx"
---

# VS Code Extension Conventions

## Architecture Layers

```
extension.ts (activate/deactivate)
  → CommandController (registration + state)
    → handlers/ (domain logic, calls fx-core)
      → core (FxCore from globalVariables)
  → TreeViewManager (sidebar UI)
  → ExtTelemetry (instrumentation)
```

## Command Registration

Register every new command through `registerInCommandController()` in `extension.ts`:

```typescript
registerInCommandController(
  context,
  CommandKeys.MyCommand,        // defined in constants
  myCommandHandler,             // handler function
  "myCommandRunningLabel"       // optional localized blocking label key
);
```

- `Correlator.run()` wraps every command with a correlation ID for tracing
- Exclusive commands block concurrent execution — the tree view shows a tooltip
- Add the command key to `package.json` contributions
- Add the display name to `package.nls.json`

## Handler Pattern

All handlers return `Promise<Result<any, FxError>>` and follow this structure:

```typescript
export async function myHandler(...args: any[]): Promise<Result<any, FxError>> {
  // 1. Telemetry start event
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.MyCommandStart,
    getTriggerFromProperty(args)
  );

  // 2. Build inputs (from args or workspace state)
  const inputs = getSystemInputs();

  // 3. Delegate to fx-core
  const result = await core.myOperation(inputs);

  // 4. Handle result
  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MyCommand, result.error);
    await showError(result.error);
    return result;
  }

  // 5. Success telemetry + return
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MyCommand, { success: "yes" });
  return result;
}
```

- Handlers live in `src/handlers/` organized by domain
- Never put business logic in handlers — delegate to fx-core
- Always extract `triggerFrom` from args for telemetry attribution

## Singleton Providers

Use the lazy `getInstance()` pattern. Export the singleton instance as default:

```typescript
class MyProvider {
  private static instance: MyProvider;
  public static getInstance(): MyProvider {
    if (!MyProvider.instance) {
      MyProvider.instance = new MyProvider();
    }
    return MyProvider.instance;
  }
}
export default MyProvider.getInstance();
```

Key singletons: `VsCodeLogProvider`, `M365Login`, `CommandController`, `TreeViewManager`

## Telemetry

```typescript
// Success
ExtTelemetry.sendTelemetryEvent(
  TelemetryEvent.MyFeature,
  { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.TreeView },
  { duration: elapsed }
);

// Error (auto-extracts errorName, errorMessage, errorStack)
ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MyFeature, error, extraProps);
```

- Define event names in `src/telemetry/extTelemetryEvents.ts`
- Auto-injected properties: Component, IsExistingUser, IsSpfx, SettingsVersion
- Always pass `getTriggerFromProperty(args)` to attribute the source (TreeView, CommandPalette, CopilotChat, etc.)

## Error Display

Use `showError()` from `src/error/common.ts` — never call `vscode.window.showErrorMessage()` directly:

- Shows `displayMessage` or `message` from `FxError`
- Adds context-aware action buttons (Debug in Playground, Troubleshoot with Agent)
- Feature-flag gated recommendations
- Sends `ShowError` telemetry automatically

## Tree Views

- `TreeViewManager` registers all providers in `registerTreeViews()`
- Each tree view is a `TreeDataProvider<TreeViewCommand>`
- `updateTreeViewsByContent()` refreshes commands based on project type and feature flags
- `setRunningCommand()` / `restoreRunningCommand()` manages exclusive command state

## React Webview Controls

- Located in `src/controls/`, built with **Vite** (separate from main esbuild bundle)
- Uses `@fluentui/react` (v8, not v9 `@fluentui/react-components`)
- Routes via `react-router-dom` `MemoryRouter`
- Internationalized via `react-intl` `IntlProvider`
- Run `Vite build` separately from the main extension build

## Global Variables

`src/globalVariables.ts` exposes shared state:
- `core` — FxCore instance
- `tools` — TOOLS singleton (LogProvider, TelemetryReporter, TokenProvider)
- `workspaceUri` — Current workspace URI
- `isSPFxProject`, `isDeclarativeCopilotApp` — Project type flags

Access these via import; never create your own FxCore instance.

## v4 Migration (Future — Phase 8)

The VS Code extension will eventually consume `@microsoft/teamsfx-core` v4 from
`packages/core-next/` instead of `packages/fx-core/`. This migration is gated by the
`TEAMSFX_V4_CORE` feature flag. Until then, all extension code continues to use fx-core.
Key v4 changes that will affect the extension:

- `TOOLS` global singleton → `AtkContext` (injected)
- `FxCore` class methods → `Operation` pipeline
- Imports shift from `@microsoft/teamsfx-api` to `@microsoft/teamsfx-core` v4
