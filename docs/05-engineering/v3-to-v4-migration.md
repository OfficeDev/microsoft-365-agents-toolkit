# v3 → v4 migration

## Workflow asymmetry

v3 and v4 have **different change workflows** — this is deliberate and is the most important practical difference for contributors:

| Engine | Workflow | Source of truth |
|--------|----------|-----------------|
| **v3** (`api`, `manifest`, `fx-core`, `cli`, `vscode-extension`, `server`) | **Code-first.** The shipping engine is the running truth. Docs catch up to it. | Code → instructions → skills → `.dev/features.json` → `docs/` prose |
| **v4** (`core-next`, `cli-next`) | **Design-first.** Non-trivial changes start with a design page in `docs/`. The design is authoritative until the code lands. | `docs/` design + ADRs → code → instructions → skills |

**Why the split:** v3 ships in millions of installs; pretending the new docs are authoritative for it would silently lie to contributors. v4 is being built from scratch alongside this docs site — writing v4 code without first writing the design forfeits the value of the structure.

**What counts as non-trivial in v4:** any new operation, driver, template descriptor, lifecycle stage, contract, command group, auth provider, telemetry shape change, or refactor that changes a public surface. Trivial changes (bug fixes, dependency bumps, test-only edits, doc-only edits) do not require a forward design page.

See [codebase.instructions.md §Source-of-Truth Workflow](../../.github/instructions/codebase.instructions.md), [docs/07-contributing/docs-contributing.md §Source-of-truth precedence](../07-contributing/docs-contributing.md), and the [`dev-test-next`](../../.github/skills/dev-test-next/SKILL.md) skill §"Phase 0 — Design First".

## What's changing

| Aspect | v3 (`fx-core` + `cli`) | v4 (`core-next` + `cli-next`) |
|--------|-----------------------|------------------------------|
| API contracts | Separate `@microsoft/teamsfx-api` package | Merged into `core-next/src/api/` |
| Context | `TOOLS` global singleton | Injected `AtkContext` |
| Operations | `FxCore` class methods | `Operation` records via `defineOperation()` + `runOperation()` |
| Templates | Inline in generators (activation order) | `TemplateRegistry` + `TemplateDescriptor` records |
| Drivers | Implicit modules | `DriverRegistry` + `createDriver()` factory + Zod pre-validation |
| Errors | `FxError` | `AtkError` (extends `FxError` shape) |
| DA features | Spread across generators + drivers | Dedicated `declarativeAgent/` module |
| Lifecycle | YAML actions dispatched ad hoc | `lifecycle/` engine (parser → resolver → executor) |
| Telemetry | `TOOLS.telemetry` + scattered helpers | DI-first `telemetry/` module with `instrumentOperation` |
| Secret masking | SVM + BloomFilter + keywords | Keywords-only regex |
| Feature flags | Singleton `FeatureFlagManager` | Injectable `FeatureFlagRegistry` |
| Localisation | `getLocalizedString()` singleton | `Localizer` class |
| HTTP client | `wrappedAxiosClient` (TOOLS-bound) | `createHttpClient(ctx)` with retry + timeout helpers |
| CLI bundler | webpack (4 GB heap) | esbuild (single CJS file) |
| Auth | `commonlib/` MSAL | `auth/` MSAL — ported, refactored, same `~/.fx/account/` cache |

## What's *not* changing

- Driver IDs match (`aadApp/create`, `arm/deploy`, etc.) — `m365agents.yml` files are compatible.
- Manifest format is identical — both engines use `@microsoft/app-manifest`.
- Token cache location and format — `~/.fx/account/`.
- Project layout on disk — `appPackage/`, `infra/`, `env/`, `m365agents.yml`.
- Telemetry event names follow the same patterns.

## Feature flag

`TEAMSFX_V4_CORE=true` opts a consumer adapter into the v4 engine. Default: off.

## For consumers (cli, vscode-extension)

When migrating a consumer:

1. Switch imports from `@microsoft/teamsfx-api` and `@microsoft/teamsfx-core` to `@microsoft/teamsfx-core` v4 (or directly from `core-next` during dev).
2. Replace `TOOLS` access with an injected `AtkContext` — provide via a factory like `createCliContext()` / `createVsCodeContext()`.
3. Replace `FxCore` method calls with `runOperation(operation, input, ctx)`.
4. Render `PostAction[]` returned by operations (open URL / show message).
5. Wire telemetry through the v4 helpers; correlation propagates automatically.

## For driver authors

When porting a driver:

1. Move from implicit `src/component/driver/` module to `core-next/src/drivers/builtin/<area>/`.
2. Wrap with `createDriver({ id, name, inputSchema: zodSchema, execute, rollback? })`.
3. Replace any `existsSync` + `readFile` with EAFP.
4. Replace any direct `process.env` reads with `ctx`-injected access (or env-map injection from the executor).
5. Add unit tests using `createMockContext()`.

## For template authors

When porting a template:

1. Add a `TemplateDescriptor` to `core-next/src/templates/descriptors/<category>.ts`.
2. Ensure `templateName` matches the actual folder name in `templates/vsc/{lang}/`.
3. Set `testable: false` only if the template needs interactive-only input.
4. Run `npm run test:unit` in `core-next` — `descriptors.test.ts` will validate it.
5. Update `.dev/features.json`.

## Migration status

- **22 built-in drivers** registered in `packages/core-next/src/drivers/builtin/` cover provision/deploy/publish for the templates v4 currently scaffolds. The headline gap is the `typeSpec/compile` driver for `declarative-agent-typespec`.
- **9 templates** verified end-to-end via cli-next (default-bot, basic-tab, custom-copilot-basic, cea/basic, cea/weather, da/basic, connector/graph-connector, plus Python variants where present).
- Coverage matrix lives in [01-product/capabilities-matrix.md](../01-product/capabilities-matrix.md) (cross-engine view) and [01-product/v3-feature-inventory.md](../01-product/v3-feature-inventory.md) (full v3 catalog incl. templates not in `features.json`).

Track progress via [features.instructions.md](../../.github/instructions/features.instructions.md) §"Current features".
