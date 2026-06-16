# Development Guideline for vscode-extension package

## Dev environment tips
- Use PowerShell scripts on Windows.
- Run `pnpm run setup` from monorepo root (`../..`) when you need a full workspace build.
- Keep code cross-platform (Windows, Linux, macOS).
- Validate changed behavior with unit tests before completion.

## Code style and testability
- Prefer namespace imports for internal modules that need stubbing in UT.
- Keep external boundaries injectable via deps objects (for example `*Deps` in handlers/utils) instead of hard-binding calls at use sites.
- Avoid side effects at module load time; initialize inside functions where possible.
- Isolate mutable global state (`globalVariables`, VS Code window/workspace objects) behind stable seams so tests can mock one place.
- Do not mix multiple mock owners for the same symbol in one test file.

## Testing instructions
- `npm run test:unit` runs unit tests **with coverage** by default.
- `npm run test:unit:vitest` runs unit tests without coverage (faster local loop).
- `npm run test:debug` runs debugger/plugin-specific UT.
- Run a focused file with:
  - `npx vitest run test/<path>.test.ts --config vitest.config.ts`
- Use Vitest-native mocks (`vi.fn`, `vi.spyOn`, `vi.mock`) for new tests.

## Test reliability and performance tips
- Keep default setup lightweight; avoid expensive per-test re-initialization in `test/setup.ts`.
- Avoid duplicate cleanup (`afterEach(vi.restoreAllMocks())`) when global setup already handles it.
- Reuse the existing compatibility utilities in `test/mocks/vitestMockUtils.ts` for legacy cases, but do not add new Sinon-style patterns unless migration requires it.
- Keep tests deterministic (no real FS/network/process side effects unless explicitly mocked).
