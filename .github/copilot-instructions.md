# Coding Style Guidelines

When generating or editing code in this repository:

- **Line Endings**: Use LF (Unix-style) line endings for all source code files. Exception: Use CRLF for localization files under `**/package.nls.*.json` only.
- **Indentation**: Use 2 spaces for TypeScript/JavaScript files
- **Quotes**: Use double quotes for strings in TypeScript/JavaScript

---

# Architecture Overview
The toolkit follows a **layered architecture** with clear separation of concerns:

┌──────────────────┐ ┌──────────────────────┐ ┌────────────────────┐
│  VS Code Ext.    │ │  CLI                 │ │  Server            │
│  (vscode-ext.)   │ │  (packages/cli)      │ │  (packages/server) │
│  - UI Commands,  │ │  - Command Line      │ │  - JSON-RPC based  │
│    TreeView,     │ │    Interface         │ │  - Used by VS ext. │
│    CodeLens      │ │                      │ │                    │
└────────┬─────────┘ └──────────┬───────────┘ └──────────┬─────────┘
         └─────────────────┬────┘                        │
                           └────────────────┬────────────┘
┌──────────────────────────▼──────────────────────────────────────────┐
│                         FX-Core Layer                               │
│   (packages/fx-core)                                                │
│   - Project Generation, Lifecycle, Drivers, Manifest Utilities      │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────────┐
│                          API Layer                                  │
│   (packages/api)                                                    │
│   - Type Definitions, Interfaces, Question Models                   │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────────┐
│                        Manifest Layer                               │
│   (packages/manifest)                                               │
│   - App Manifest Types, Converters, OOP Wrappers                    │
│   - Three manifest types: Teams, Declarative Agent, API Plugin      │
└─────────────────────────────────────────────────────────────────────┘

---

# Package-Specific Instructions

Detailed instructions for individual packages are available in `.github/instructions/`:

| Package | Instructions File | Key Exports |
|---------|-------------------|-------------|
| `@microsoft/app-manifest` (packages/manifest) | `.github/instructions/manifest.instructions.md` | `TeamsManifest`, `DeclarativeAgentManifest`, `APIPluginManifest`, `AppManifestUtils`, Wrappers |

> **Note**: When working with a specific package, read the corresponding instructions file for detailed documentation about the package's architecture, APIs, and patterns

---

# Manifest Package Quick Reference

The manifest package (`packages/manifest`) provides TypeScript types for Microsoft 365 app manifests:

| Manifest Type | Purpose | Wrapper | Latest Type |
|---------------|---------|---------|-------------|
| **Teams Manifest** | Core M365 apps (bots, tabs, extensions) | `TeamsManifestWrapper` | `TeamsManifestLatest` |
| **Declarative Agent** | AI agents with instructions & actions | `DeclarativeAgentManifestWrapper` | `DeclarativeAgentManifestLatest` |
| **API Plugin** | REST API plugin capabilities | `PluginManifestWrapper` | `APIPluginManifestLatest` |

**Key utilities:**
- `AppManifestUtils` - Read/write/validate manifests
- `*Converter` classes - JSON ↔ typed object conversion
- Wrappers provide fluent APIs for manipulation

**Deprecated types** (use generated types instead):
- `TeamsAppManifest` → use `TeamsManifest`
- `DeclarativeCopilotManifestSchema` → use `DeclarativeAgentManifest`
- `PluginManifestSchema` → use `APIPluginManifest`
- `ManifestUtil` → use `AppManifestUtils`

**Schema upgrade guardrail (manifest package):**
- Run `node download.js` in `packages/manifest` before `npm run convert` to sync new schema folders.
- After conversion, update `src/generated-types/index.ts` so new versions are registered in converter maps, unions, and `*Latest` aliases (and re-export new schema enums/types when applicable).
- Do not leave new versions relying on fallback unchecked casts; validate parity with `npx mocha test/converterMapParity.test.ts`.

---

# Unit Testing Guidelines

When fixing unit tests for a package:

1. **Navigate to the package directory** before running tests
2. **Run the full test suite** to identify failures:
   ```bash
   npm run test:unit
   ```
3. **For large test suites or targeted debugging**, run specific test files directly:
   ```bash
   npx nyc mocha --no-timeouts --require ts-node/register <path_to_test_file_or_folder>
   ```
4. **Fix errors iteratively** - run tests after each fix to verify

> **Tip**: When there are many failing tests, start with the specific file causing issues to reduce feedback loop time.

---

# Template Maintenance Guidelines

Templates in this repository (located in `templates/`) are used when scaffolding.

## Template Locations
- `templates/` - Root location for all scaffolding templates
- `templates/src/` - Template metadata, question/template name definitions, and UI assets
- `templates/vsc/` - VS Code scaffolding templates (TypeScript, JavaScript, Python, and shared common)
- `templates/vs/` - Visual Studio scaffolding templates (C#)
- `templates/unused/` - Templates not currently in use but kept for reference, no need to update them
