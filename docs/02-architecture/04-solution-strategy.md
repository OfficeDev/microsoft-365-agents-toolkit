# 4 — Solution strategy

The fundamental architectural decisions, distilled. Each maps to one or more ADRs in [09-architecture-decisions/](09-architecture-decisions/README.md).

## S1 — Single shared engine; thin per-surface adapters

VS Code, Visual Studio, and CLI each are presentation adapters around one engine.

- v3: `fx-core` is the engine; `vscode-extension`, `cli`, `server` adapt it.
- v4: `core-next` is the engine; `cli-next` adapts it; future `vscode-extension` will too.

This keeps lifecycle behaviour consistent across surfaces and makes E2E testing the engine sufficient to validate all surfaces.

## S2 — Two coexisting engine generations behind a feature flag

Rather than rewrite-and-replace, v4 (`core-next` + `cli-next`) ships **alongside** v3 and is opt-in via `TEAMSFX_V4_CORE`. Both engines are tested in CI. v4 lifts the patterns proven in v3 while replacing globals with DI and implicit conventions with explicit registries.

ADRs: [0001-feature-flag-v4-core](09-architecture-decisions/0001-feature-flag-v4-core.md)

## S3 — Errors as values

Every fallible operation returns `Result<T, FxError>` (`neverthrow`), with concrete `UserError` / `SystemError` (v3) or `AtkError` (v4). No throwing for expected failures. This makes telemetry and localised display uniform.

ADRs: [0003-result-pattern-neverthrow](09-architecture-decisions/0003-result-pattern-neverthrow.md)

## S4 — Declarative lifecycle (`m365agents.yml`)

Provision / deploy / publish are not hard-coded — they are interpretations of YAML steps. A **driver registry** maps action IDs (e.g. `arm/deploy`, `teamsApp/publishAppPackage`) to implementations. New capabilities ship by adding a driver, not by editing the engine.

In v4, drivers are explicit `DriverDescriptor` records produced by `createDriver()` with **Zod pre-validation**, telemetry, and error normalisation built in.

ADRs: [0004-zod-driver-validation](09-architecture-decisions/0004-zod-driver-validation.md)

## S5 — Templates as first-class data

Templates are described by `TemplateDescriptor` records in a `TemplateRegistry` (v4). Adding a template auto-generates a CLI subcommand and wires the question tree. The actual scaffold files still live under `templates/vsc/{lang}/{template-name}/`.

The scaffold pipeline downloads remote ZIPs (or falls back to bundled ones in `packages/core-next/templates/fallback/`), filters by template prefix, and renders Mustache (`.tpl` files).

## S6 — Esbuild for the modern bundles

`vscode-extension` and `cli-next` use **esbuild** (`esbuild.mjs`) for single-file CJS bundles, replacing the legacy webpack pipeline that needed 4 GB of heap. `core-next` uses pure `tsc` (it is bundled inline by `cli-next`).

ADRs: [0002-esbuild-over-webpack](09-architecture-decisions/0002-esbuild-over-webpack.md)

## S7 — Async correlation with `AsyncLocalStorage` (v4)

Correlation IDs propagate across awaits without explicit threading via `AsyncLocalStorage` (`telemetry/correlationScope`). v3 uses a `Correlator.run()` wrapper. Both produce the same observable outcome: every telemetry event for a single user action shares one correlation ID.

ADRs: [0005-async-local-storage-correlation](09-architecture-decisions/0005-async-local-storage-correlation.md)

## S8 — Schema-typed manifest layer

Manifest manipulation goes through `@microsoft/app-manifest`'s `TeamsManifestWrapper`. Engine code never edits raw JSON. This isolates the rest of the codebase from schema-version churn.

## S9 — DI everything, singletons nowhere (v4)

`AtkContext` carries `logger`, `telemetry`, `ui`, `auth`, `correlationId`. Every operation, driver, and client takes it as a parameter. No module-scoped state. This makes parallel lifecycles, isolated tests, and fan-out scenarios trivial.

## S10 — Security by default

EAFP filesystem access · Zip-Slip guards on archive extraction · magic-byte validation on ZIP uploads · keyword-based secret masking · no PII in telemetry · MSAL token cache encrypted with AES-256-GCM (with `keytar` if available).

See [08-crosscutting-concepts.md](08-crosscutting-concepts.md) and [05-engineering/security.md](../05-engineering/security.md).
