# Testing strategy

> For v4 lifecycle code (`core-next/src/lifecycle/`, `core-next/src/drivers/builtin/`, `core-next/src/clients/`, `cli-next/src/actions/`, `cli-next/src/commands/`) the **inverted test pyramid** applies: integration and E2E tests carry the verification weight, with unit tests reserved for pure functions. See [ADR 0007 — Inverted test pyramid](../02-architecture/09-architecture-decisions/0007-inverted-test-pyramid-for-lifecycle.md) for the full rationale and the per-layer rules. The conventional pyramid (lots of unit tests) still applies for `secretMasker`, `featureFlags`, `localization`, and other pure-logic modules.

## Stack

| Tool | Role |
|------|------|
| Mocha | Test runner |
| Chai | Assertions |
| Sinon | Stubs / spies / fakes |
| NYC (Istanbul) | Coverage |
| `@istanbuljs/nyc-config-typescript` | TS coverage config |
| ts-node/register | Mocha TS support |

## Layers

| Layer | Where | Boundary |
|-------|-------|----------|
| Unit | `tests/unit/` | Isolate one module; mock all I/O |
| Integration | `tests/integration/` | Multiple modules together; mock only external services |
| E2E | `packages/cli-next/tests/e2e/` and `tests/` package | Real CLI invocation against real (or recorded) services |

## Counts (current)

| Package | Unit | Integration |
|---------|------|-------------|
| core-next | 606 | 48 |
| cli-next | 87 | 81 |
| fx-core | thousands (50+ granular suites) | many |

## Test conventions

- File naming: `*.test.ts`, mirroring `src/` structure under `tests/unit/` and `tests/integration/`.
- Group with `describe("<ComponentName>")` → `describe("<methodName>")` → `it("should <behaviour> when <condition>")`.
- Always `sinon.restore()` in `afterEach`.
- Test both `ok` and `err` paths for every `Result`-returning function.

```ts
describe("FxCore", () => {
  describe("createProject", () => {
    afterEach(() => sinon.restore());

    it("should return ok when input is valid", async () => {
      const result = await core.createProject({ ... });
      expect(result.isOk()).to.be.true;
    });

    it("should return UserError when language is missing", async () => {
      const result = await core.createProject({});
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().name).to.equal("MissingLanguageError");
    });
  });
});
```

## v4 test helper

`tests/unit/testHelper.ts` exports `createMockContext()` — a fully-stubbed `AtkContext` with sinon stubs for `telemetry.sendEvent`, `telemetry.sendErrorEvent`, `logger`, `ui`, and `auth`. Use in every core-next unit test.

## Configuration files (per package)

| File | Purpose |
|------|---------|
| `.mocharc.js` | `ts-node/register`, spec reporter, `no-experimental-strip-types` |
| `.nycrc` | Extends `@istanbuljs/nyc-config-typescript`; coverage thresholds; excludes `src/api/**/*` in core-next |

## Test scripts

```bash
npm run build               # tsc + postbuild
npm run test:unit           # unit tests with NYC coverage
npm run test:integration    # integration tests (no coverage)
npm run test                # alias for test:unit
npm run lint                # ESLint check (0 errors required)
npm run format              # Prettier auto-format
npm run format:check        # Prettier check (CI gate)
```

## Coverage gates

| Package | Gate |
|---------|------|
| core-next | 80% via `nyc` thresholds + CI guard in `ci-next.yml` |
| cli-next | (no hard gate yet; tracked) |
| fx-core | per-suite; codecov reports |

## Best practices

- Mock at I/O boundaries (HTTP, filesystem, Azure SDK), not internal logic.
- Avoid testing implementation details — test behaviour and outputs.
- Use `sinon.sandbox` or per-test stubs; restore in `afterEach`.
- Name tests as `should <expected behaviour> when <condition>`.
- For async sequences, prefer `await` chains over `.then()`.

## CI

| Workflow | Scope |
|----------|-------|
| `ci-next.yml` | core-next + cli-next: build → lint, format-check, unit-test (80% coverage gate), integration-test |
| `unit-test.yml` (legacy) | v3 packages |
| `e2e-test-next.yml` | Daily E2E for cli-next + on PR + manual |

E2E failures publish a GitHub Step Summary with stats, failed test table, and log tail.
