# Telemetry

## Goals

- Every user action emits one start event and one terminal (success / failure) event.
- All events from a single user action share a `correlationId`.
- Errors are partitioned by stable `error.name`.
- No PII or secrets ever leave the process.

## v3

```typescript
// Success
ExtTelemetry.sendTelemetryEvent(
  TelemetryEvent.MyFeature,
  { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.TreeView },
  { duration: elapsed },
);

// Error (auto-extracts errorName, errorMessage, errorStack)
ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.MyFeature, error, extraProps);
```

| Concern | v3 implementation |
|---------|-------------------|
| Reporter | `ExtTelemetry` (vscode), `CliTelemetry` (cli), `TelemetryReporter` interface in `api` |
| Correlation | `Correlator.run(uuid, async () => { ... })` wrapping handlers |
| Auto properties | `component`, `commandName`, `projectId`, `correlationId`, `runFrom` (CI detection) |

## v4

```typescript
import { instrumentOperation, sendStartEvent, sendErrorEvent } from "../telemetry";

// Compose into Operation
const myOp = defineOperation("my-op", inputSchema, async (input, ctx) => {
  return instrumentOperation(ctx, "my-op", async () => {
    // ... logic ...
  });
});
```

| Concern | v4 implementation |
|---------|-------------------|
| Reporter | `TelemetryReporter` interface; CLI adapter wraps `applicationinsights` (lazy-loaded) |
| Correlation | `correlationScope(uuid, async () => { ... })` using `AsyncLocalStorage` + `ctx.correlationId` |
| Helpers | `sendStartEvent`, `sendSuccessEvent`, `sendErrorEvent`, `instrumentOperation`, `extractErrorProperties` |
| Auto properties | Same as v3 |

All v4 helpers take `ctx: AtkContext` as the first arg — no globals.

## Event taxonomy

| Suffix | Meaning |
|--------|---------|
| `-start` | Action initiated |
| `-end` | Action completed (with `success: "yes"` / `"no"`) |
| `driver-start` / `driver-end` | Per-driver execution |

Define event names in `TelemetryEvent` enums per package.

## Property masking

The secret masker runs over every property value before transport:

- v3: `common/secretmasker/` (SVM + BloomFilter + keywords)
- v4: `core-next/src/secretMasker/` (keywords only)

URLs in HTTP client logs are sanitised (query-string secrets stripped).

## Opt-out

End users can opt out via:

- VS Code: standard "telemetry off" settings.
- CLI: `atk config set telemetry off`.

## Local debugging

Set `TEAMSFX_TELEMETRY_TRACE=true` to log telemetry events to the console instead of (or in addition to) sending. Useful when verifying a new event shape.
