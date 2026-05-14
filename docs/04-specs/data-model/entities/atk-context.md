# `AtkContext`

The injected context object that replaces v3's `TOOLS` global singleton. Every v4 operation, driver, client, and helper takes it as a parameter.

Source: [`packages/core-next/src/core/AtkContext.ts`](../../../packages/core-next/src/core/AtkContext.ts).

## Shape

```typescript
interface AtkContext {
  logger: LogProvider;                         // info / warn / error / debug / verbose
  telemetry: TelemetryReporter;                // sendEvent, sendErrorEvent
  ui: UserInteraction;                         // singleSelect, input, confirm, ...
  auth: TokenProvider;                         // m365TokenProvider, azureAccountProvider
  correlationId: string;                       // uuid for this user action
  projectPath?: string;                        // set by lifecycle operations
}
```

## Construction

| Surface | Factory |
|---------|---------|
| cli-next | `createCliContext(projectPath?)` in [`packages/cli-next/src/context.ts`](../../../packages/cli-next/src/context.ts) |
| Tests | `createMockContext()` in [`packages/core-next/tests/unit/testHelper.ts`](../../../packages/core-next/tests/unit/testHelper.ts) — fully sinon-stubbed |
| Future vscode-extension | will provide a `createVsCodeContext()` |

## What lives in `ctx.auth`

```typescript
interface TokenProvider {
  m365TokenProvider: M365TokenProvider;
  azureAccountProvider: AzureAccountProvider;
}
```

CLI auth providers are real MSAL-based implementations in [`packages/cli-next/src/auth/`](../../../packages/cli-next/src/auth/). The `createTokenProvider()` factory detects CI mode and returns service-principal-backed providers when applicable.

## Usage rules

- **Always pass `ctx`** — never read auth, telemetry, or logger from a module-scoped variable.
- **Never mutate `ctx`** — except `ctx.projectPath` (set once at the top of an operation that knows the project).
- **In tests use `createMockContext()`** — never construct an `AtkContext` literal by hand; the helper keeps shape in sync with the interface.

## Why this shape

- `logger` separated from `telemetry` so log lines and metric events have different lifetimes (logs flush immediately; telemetry batches).
- `ui` separated from `logger` so non-interactive surfaces (CI) can route prompts to defaults without changing the logger sink.
- `auth` separated so tests can substitute a token provider without rebuilding the rest of the context.
- `correlationId` carried explicitly so helpers that don't have access to `AsyncLocalStorage` (e.g. logging interceptors) can still attribute events.
