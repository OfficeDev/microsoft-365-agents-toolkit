# 8 — Cross-cutting concepts

These appear in every layer; defining them once here keeps per-package docs short.

## Result + error model

Every fallible function returns `Result<T, FxError>` (`neverthrow`).

| | v3 | v4 |
|-|----|----|
| Base interface | `FxError` from `@microsoft/teamsfx-api` | same |
| User-recoverable | `UserError` | `AtkError` with `kind: "user"` |
| Infrastructure | `SystemError` | `AtkError` with `kind: "system"` |
| Re-export | `ok`, `err`, `Result` from `@microsoft/teamsfx-api` | from `@microsoft/teamsfx-core` |

Constructor: always `{ source, name, message, displayMessage?, helpLink?, error? }`. The `name` is the telemetry identifier — keep stable across releases. See [04-specs/data-model/entities/result-and-fxerror.md](../04-specs/data-model/entities/result-and-fxerror.md).

## Telemetry

All entry points emit start / end events with shared properties:

| Property | Source |
|----------|--------|
| `correlationId` | `Correlator` (v3) or `AsyncLocalStorage` (v4) |
| `component` | Package or module name |
| `commandName` | CLI / VS Code / VS command identifier |
| `projectId` | Tracking ID from project (if present) |
| `runFrom` | CI platform detection (GitHub Actions, ADO, local) |

Errors include `error.source`, `error.name`, `error.message` — never PII. v4 helper: `instrumentOperation()`. See [05-engineering/telemetry.md](../05-engineering/telemetry.md).

## Localisation

User-facing strings: `getLocalizedString("teamsfx.key", ...params)` (v3) or `Localizer.getString(...)` (v4). Default English: `getDefaultString(...)`. Translations live in `Localize/loc/` and per-package `package.nls.*.json`. Never concatenate translated fragments.

## Feature flags

| | v3 | v4 |
|-|----|----|
| Reader | `featureFlagManager.getBooleanValue(name)` | `registry.isEnabled(name)` |
| Definition | `FeatureFlagName` enum | `FeatureFlagDescriptor` records |
| Source | `process.env` | injectable `FeatureFlagSource` (defaults to `process.env`) |

Headline flag: **`TEAMSFX_V4_CORE`** — opt-in to the v4 engine.

## Security

- **EAFP filesystem.** Never check existence before operating. Catch `ENOENT`. Avoids TOCTOU.
- **Zip Slip.** Validate every entry name on extract (`.indexOf("..")` AND `path.resolve()` containment). See `templates/scaffold/download.ts`.
- **Magic bytes.** Verify `PK\x03\x04` before uploading ZIPs to services. See `clients/teamsDevPortal/client.ts`, `drivers/builtin/azureAppService/zipDeploy.ts`.
- **Secret masking.** `secretMasker` redacts credential-like values. v4 is keyword-based (regex), v3 adds SVM + BloomFilter.
- **Token cache.** AES-256-GCM at rest under `~/.fx/account/`. `keytar` upgrades to OS keychain when available.
- **No PII in telemetry.** Error `message` is sanitised; `displayMessage` (which may contain user data) is *not* sent.

## Async + correlation

- All I/O is `async/await`; no `.then()` chains.
- ESLint `promise/no-floating-promises` enforces every promise is awaited or returned.
- Correlation ID is set once per user action and threaded through telemetry events. v3: `Correlator.run(uuid, async () => {...})`. v4: `correlationScope(uuid, async () => {...})` using `AsyncLocalStorage`.

## Logging

- Use the injected `LogProvider` (v3 `TOOLS.logger`) / `ctx.logger` (v4). Never `console.log`.
- Log levels: `verbose`, `debug`, `info`, `warning`, `error`.
- Logs are sanitised by the secret masker before reaching transports.
