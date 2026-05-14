# 08 — Troubleshooting

User-facing error reference. Each page lists the error class, what triggers it, and the mitigation.

## Pages

- [bot-errors.md](bot-errors.md) — bot service / channel registration errors
- [provision-errors.md](provision-errors.md)
- [deploy-errors.md](deploy-errors.md)
- [auth-errors.md](auth-errors.md)

## Pre-existing reference

Legacy bot help content lives at [`docs/_v3-reference/bot-help.md`](../_v3-reference/bot-help.md) and is **kept in place** to preserve external links from the toolkit's `helpLink` URLs. The [bot-errors.md](bot-errors.md) page in this section indexes that content alongside newer entries.

## Where errors come from

Every error surfaces with a stable `error.name` (the telemetry partition key). When troubleshooting:

1. Look up the `error.name` in the relevant page below.
2. If the error has a `helpLink`, follow it — it points at the most current mitigation.
3. If the page suggests `--debug`, re-run the command with that flag to capture verbose logs.
