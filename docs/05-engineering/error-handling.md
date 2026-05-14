# Error handling

Cross-cutting reference. For data shape see [04-specs/data-model/entities/result-and-fxerror.md](../04-specs/data-model/entities/result-and-fxerror.md). For UX see [01-product/ux/error-display-and-help.md](../01-product/ux/error-display-and-help.md). For security implications see [security.md](security.md).

## Mantras

1. **Errors are values.** Return `Result<T, FxError>`; do not throw for expected failures.
2. **Errors are help.** Every `UserError` has `displayMessage` and ideally `helpLink`.
3. **Names are stable.** `error.name` is the telemetry partition — never rename across releases.
4. **Source is required.** Set the component name; never empty-string.

## Decision: `UserError` vs `SystemError`

| Pick | When |
|------|------|
| `UserError` | The user can fix it: bad input, missing config, auth failure, missing permission |
| `SystemError` | The user cannot fix it: service outage, unexpected SDK exception, code bug |

Wrong choice misroutes UX (action buttons) and skews telemetry. Err on the side of `UserError` when in doubt and the user has *some* path to remediation.

## Wrapping

Always preserve the inner error chain:

```ts
try {
  await externalCall();
} catch (e) {
  return err(new SystemError({
    source: "Coordinator",
    name: "ExternalCallFailed",
    message: getDefaultString("teamsfx.coord.externalFailed"),
    error: e instanceof Error ? e : new Error(String(e)),
  }));
}
```

## Localisation

| Field | Localised? |
|-------|-----------|
| `message` | no — English, for logs and telemetry |
| `displayMessage` | yes — `getLocalizedString` (v3) / `Localizer.getString` (v4) |
| `helpLink` | n/a — opaque URL, typically `aka.ms/...` |

## Driver-level normalisation (v4)

`createDriver()` automatically:

- Wraps thrown exceptions into `DriverExecutionError` (`SystemError`).
- Recognises plain `AtkError` shapes (with `code`/`message`/`kind`) and returns them unwrapped — prevents `[object Object]` serialisation.
- Pre-validates input via Zod; failures surface as `InvalidDriverInput` with the Zod issue path.

## Common error names

Keep these stable; new error types should pick a fresh `name`:

- `MissingRequiredOptionError`, `MissingRequiredArgumentError`, `InvalidChoiceError` (CLI)
- `UnknownCommandError`, `UnknownOptionError` (CLI)
- `ManifestValidationError`, `PackageValidationError` (manifest drivers)
- `InvalidDriverInput`, `DriverExecutionError` (drivers)
- `AzureLoginFailed`, `M365LoginFailed` (auth)
- `ResourceGroupNotFoundError`, `SubscriptionNotFoundError` (Azure)

## Anti-patterns

- Catching everything and returning a single `UnknownError`.
- Renaming `error.name` between releases.
- Including PII in `message` (`displayMessage` may include user data; `message` must not).
- Re-throwing without `innerError`.
- Using bare strings instead of localised keys.
