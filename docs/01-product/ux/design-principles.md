# Design principles

## P1 — One mental model across surfaces

A user who learns the lifecycle in VS Code should recognise the same shape in the CLI and Visual Studio. Concretely:

- Same five stages: scaffold → run/debug → provision → deploy → publish.
- Same vocabulary in every surface (and in this site's [glossary](../00-overview/glossary.md)).
- Same per-environment behaviour (`env dev`, `env prod`).
- Same outputs (env-file edits, manifest files, Bicep).

## P2 — Progressive disclosure

Defaults work. Advanced options are reachable but not in the way. In question trees this means: surface only the questions whose answer the engine cannot infer, in the order they are needed.

## P3 — Errors are help

Every `UserError` carries a `displayMessage` (localised) and ideally a `helpLink`. The `showError` path adds context-aware action buttons (Debug in Playground, Open file, Troubleshoot with Agent). Errors are never bare strings — they are a UX surface.

## P4 — Consent before consequence

The lifecycle inserts explicit consent gates before destructive or expensive actions:

- `confirmProvision` — before resource creation in Azure / Entra.
- `confirmDeploy` — before pushing code to provisioned resources (skipped for local-ish envs).

Headless CI bypasses these via `CI_ENABLED=true`; interactive surfaces always show them.

## P5 — Make CI easy

Every interactive prompt has a non-interactive analogue (CLI flag, env var, or sensible default). `CI_ENABLED=true` switches the whole CLI to non-interactive mode; missing required values fail loudly with named errors.

## P6 — Telemetry is invisible

Telemetry must never block, slow down, or surface to the user except via opt-out (`telemetry off`). All emissions are async and best-effort; failures are swallowed.

## P7 — Localisation is uniform

Every user-facing string is a key, never a concatenation. The locale switch flips the entire surface, including driver progress messages. See [localization.md](localization.md).
