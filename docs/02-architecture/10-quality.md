# 10 — Quality

Quality goals from [01-introduction-and-goals.md](01-introduction-and-goals.md), expanded into testable scenarios.

## Quality tree

```
Quality
├── Reproducibility
│   ├── Provision N times → same envMap (modulo random suffix)
│   └── Headless CI run → matches interactive run for same inputs
├── Compatibility
│   ├── v3 project loads in v4 with no manual edits
│   └── Manifest schema 1.x project upgrades cleanly to 2.4
├── Security
│   ├── No secrets in logs / telemetry
│   ├── ZIP extraction never escapes target dir
│   └── ZIP uploads never carry non-ZIP bytes
├── Performance
│   ├── `atk --help` < 200ms cold
│   └── Provision overhead (engine only, excluding ARM) < 5s
├── Observability
│   └── Every user action has a single correlation ID across all events
└── Localisation
    └── Every user-facing string is keyed; no concatenation
```

## Quality scenarios

### Reproducibility

> Given a project with `m365agents.yml`, when `provision` is run twice with the same `--env`, then the second run is a no-op (or makes only idempotent updates).

Verified by integration tests in `tests/integration/lifecycleExecution.test.ts` (driver outputs are merged into envMap; second run sees existing IDs).

### Compatibility

> Given a v3 project, when `TEAMSFX_V4_CORE=true` is set and `provision` is run via cli-next, then the same Azure resources are produced as the v3 path.

Verified manually in v4 E2E tests; tracked in [features.instructions.md](../../.github/instructions/features.instructions.md) "v4 E2E" column.

### Security

> Given any log line containing `password`, `secret`, `token`, `key`, or 100+ other suffixes, when emitted, then the value is replaced with `<REDACTED>`.

Verified by `tests/unit/secretMasker/masker.test.ts`.

### Performance

> Given a fresh Node process, when `atk --help` is invoked, then it completes in under 200 ms p95.

Achieved through esbuild bundling and lazy-loading (`applicationinsights`, `registerBuiltinDrivers`, `node-machine-id`).

### Observability

> Given any single user action, when telemetry is captured, then all events emitted by that action share the same `correlationId`.

Verified by `tests/integration/operationPipeline.test.ts` and `tests/unit/telemetry/correlation.test.ts`.

### Localisation

> Given any user-visible string, when extracted, then it is keyed in `package.nls.json` and translated in all 13 supported locales.

Verified manually during release; the `Localize/loc/` directory is the source of truth.
