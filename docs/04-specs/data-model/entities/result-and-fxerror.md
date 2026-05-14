# `Result` and `FxError` / `AtkError`

Errors are values. `Result<T, FxError>` is the universal return type for fallible operations.

## Where it comes from

- v3: `neverthrow` re-exported by `@microsoft/teamsfx-api` (`ok`, `err`, `Result`, `ResultAsync`).
- v4: same `neverthrow` re-exported by `@microsoft/teamsfx-core` (after the merge).

## Error type hierarchy

```
FxError (interface in api)
  ├── UserError      — user-fixable; carries displayMessage, helpLink, action buttons
  └── SystemError    — infrastructure; carries report-issue affordance
```

v4 `AtkError` extends the same shape; the discriminator is `kind: "user" | "system"`.

## Construction

```typescript
return err(new UserError({
  source: "AppStudioPlugin",                       // component name (REQUIRED)
  name: "ManifestValidationError",                 // stable telemetry id (REQUIRED)
  message: getDefaultString("teamsfx.appStudio.invalidManifest", details),  // English, for logs
  displayMessage: getLocalizedString(              // localised, for UI
    "teamsfx.appStudio.invalidManifest", details),
  helpLink: "https://aka.ms/teamsfx-manifest-validation",
  error: innerErr,                                 // wrapped inner error
}));
```

## Field purposes

| Field | Sent to telemetry? | Shown to user? | Localised? |
|-------|-------------------|----------------|-----------|
| `source` | yes | no | n/a |
| `name` | yes (stable id!) | no | n/a |
| `message` | yes | only if no displayMessage | no — English |
| `displayMessage` | **no** (may contain user data) | yes | yes |
| `helpLink` | yes | yes | n/a |
| `innerError` | name only | no | n/a |

## Telemetry contract

`name` appears as `error.name` in telemetry. **Never rename it across releases** — it is the partition key for failure analysis. Change `message` / `displayMessage` for clarity; keep `name` stable.

## Result patterns

```typescript
// chain
const result = await op1(input)
  .andThen(out1 => op2(out1))
  .mapErr(e => new SystemError({ ...e, source: "Coordinator" }));

// branch
if (result.isErr()) {
  await showError(result.error);
  return result;
}
return ok(result.value);
```

## Anti-patterns

- Throwing for expected failures.
- Catching and re-throwing without `innerError`.
- Renaming `name` between releases.
- Putting PII in `message`.
