# Monorepo layout

## Tooling

- **pnpm workspaces** (`pnpm-workspace.yaml`) for dependency hoisting.
- **Lerna** (`lerna.json`) for ordered builds and per-package publish.
- **commitlint** for Conventional Commits enforcement.
- **Renovate** for upstream dependency PRs (`.github/renovate.json`).

## Top-level layout

```
microsoft-365-agents-toolkit-new/
├── .github/                # CI workflows + Copilot instructions + skills + policies
├── .agents/                # Workspace-scoped agent skills
├── docs/                   # this site
├── Localize/loc/           # 13 locales of UI strings
├── packages/               # 20 packages
├── templates/              # project scaffolds shipped to users
└── ROADMAP.md, README.md, ...
```

## Dependency flow

```
# v3 (current)
api ──▶ manifest ──▶ fx-core ──▶ cli / vscode-extension / server
                              ──▶ templates (build output copied into fx-core)

# v4 (next)
core-next ──▶ cli-next
          ──▶ (future) vscode-extension / server
```

Changing `api` impacts all v3 downstream packages — rebuild before testing.
Changing `core-next` impacts `cli-next` and future v4 consumers.

## Packages

| Package | npm name | Role |
|---------|----------|------|
| `api` | `@microsoft/teamsfx-api` | v3 public contract surface |
| `manifest` | `@microsoft/app-manifest` | Manifest types, schema, converters |
| `fx-core` | `@microsoft/teamsfx-core` | v3 engine |
| `server` | `@microsoft/teamsfx-server` | JSON-RPC server bridge to fx-core |
| `vscode-extension` | `ms-teams-vscode-extension` | VS Code extension UI + handlers |
| `cli` | `@microsoft/m365agentstoolkit-cli` | v3 CLI (`atk` binary) |
| `core-next` | `@microsoft/teamsfx-core-next` v4.0.0 | **v4** engine |
| `cli-next` | `@microsoft/m365agentstoolkit-cli-next` v4.0.0 | **v4** CLI |
| `sdk` | `@microsoft/teamsfx` | App-side SDK (auth helpers) |
| `sdk-react` | `@microsoft/teamsfx-react` | React hooks for the SDK |
| `dotnet-sdk` | NuGet `Microsoft.TeamsFx` | C# SDK |
| `function-extension` | `Microsoft.Azure.WebJobs.Extensions.TeamsFx` | Azure Functions auth-aware bindings |
| `mcp-server` | `@microsoft/m365-mcp-server` | Model Context Protocol server |
| `spec-parser` | `@microsoft/m365-spec-parser` | OpenAPI parser/validator/filter |
| `simpleauth` | — | Auth proxy for legacy Tab SSO |
| `metrics-ts` | — | Internal metrics utility |
| `adaptivecards-tools-sdk` | `@microsoft/adaptivecards-tools` | Adaptive Card helpers |
| `eslint-plugin-teamsfx` | local | Shared ESLint flat-config |
| `prettier-config` | local | Shared Prettier config |
| `vscode-ui` | local | Shared React UI components for the extension |
| `extra-shot-mocha` | local | Mocha extension utility |
| `tests/` | local | Cross-package E2E suites |

See [package-reference/](package-reference/README.md) for per-package detail.

## Change placement guide

| Change Type | Package |
|-------------|---------|
| New v3 interface or contract | `api` |
| Manifest schema or validation | `manifest` |
| New v3 generator, driver, or coordinator | `fx-core` |
| v3 CLI command | `cli` |
| New v4 CLI action | `cli-next/src/actions/` |
| New v4 CLI command factory or slug mapping | `cli-next/src/commands/factory.ts` |
| New v4 contract / interface | `core-next/src/api/` |
| New v4 operation, driver, or template descriptor | `core-next` (specific subdir) |
| New v4 built-in driver implementation | `core-next/src/drivers/builtin/` |
| New v4 service client | `core-next/src/clients/` |
| New v4 Declarative Agent feature | `core-next/src/declarativeAgent/` |
| New v4 lifecycle action / operation | `core-next/src/lifecycle/` |
| New v4 question or template registration | `core-next/src/questions/` or `templates/descriptors/` |
| VS Code UI, commands, tree views | `vscode-extension` |
| New project template | `templates/` + descriptor in `core-next` (and metadata in `fx-core` for v3) |

## License header

Every `.ts` source file must start with:

```ts
/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */
```

Enforced by ESLint header rule.
