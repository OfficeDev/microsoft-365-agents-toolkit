# Architecture Overview
The toolkit follows a **layered architecture** with clear separation of concerns:

┌─────────────────────────────────────────────────────────────────────┐
│                     VS Code Extension Layer                         │
│   (packages/vscode-extension)                                       │
│   - UI Commands, Handlers, TreeView, CodeLens Providers             │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────────┐
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
│   - Manifest Schema Types, Converters                               │
└─────────────────────────────────────────────────────────────────────┘

---

# Package-Specific Instructions

Detailed instructions for individual packages are available in `.github/instructions/`:

| Package | Instructions File |
|---------|-------------------|
| `@microsoft/app-manifest` (packages/manifest) | `.github/instructions/manifest.instructions.md` |

> **Note**: When working with a specific package, read the corresponding instructions file for detailed documentation about the package's architecture, APIs, and patterns
