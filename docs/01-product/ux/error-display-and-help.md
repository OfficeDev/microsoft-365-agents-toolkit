# Error display and help

A failure is a UX surface — not a dead end.

## The `showError` flow (VS Code)

`packages/vscode-extension/src/error/common.ts` exports `showError(error: FxError)`. **No handler ever calls `vscode.window.showErrorMessage` directly.**

`showError` does:

1. Reads `displayMessage` (localised, may include user data) — falls back to `message` (English, sanitised).
2. Appends a `helpLink` if present.
3. Adds context-aware action buttons:
   - **Debug in Playground** — when error is a `LocalDebugFailure`.
   - **Troubleshoot with Agent** — opens the Copilot Chat participant with the error pre-loaded.
   - **Report Issue** — for `SystemError`.
   - Feature-flag gated: additional recommendations.
4. Sends a `ShowError` telemetry event with `error.source`, `error.name`.

## CLI error display

`cli-next` renders errors via `wrapHandler()`:

| Field | Rendered as |
|-------|-------------|
| `name` | Bold red header |
| `displayMessage` | Body |
| `helpLink` | Cyan-coloured link |
| `innerError.message` | Indented subtext |

`SystemError` adds a "report issue" instruction with a pre-filled GitHub issue URL.

## Help links are stable

`helpLink` URLs use the `aka.ms/...` short-link service. The mapping is owned outside this repo. When you create a new error type, **request an aka.ms first** — don't ship a deep `learn.microsoft.com` URL that is likely to rot.

## Action buttons feature flags

VS Code action buttons are gated by feature flags so they can be safely rolled out:

| Button | Flag |
|--------|------|
| Troubleshoot with Agent | (always-on once Copilot Chat is available) |
| Debug in Playground | per-error |
| Recommend custom fix | per-error |

See `packages/vscode-extension/src/error/common.ts` and the per-error recommendation map.

## Anti-patterns to avoid

- Catching `FxError` and re-throwing a fresh one without `innerError` — loses the chain.
- Putting user PII in `message` — only `displayMessage` is exempt from the no-PII rule (and even there, prefer not to).
- Calling `console.error` instead of using the `LogProvider` — bypasses secret masker.
