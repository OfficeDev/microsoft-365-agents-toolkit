# Schemas

Pointers to the JSON / YAML schemas the toolkit consumes and produces. Each entry in this list points at the *authoritative* schema source; this site does not re-publish them.

## App manifest

| Schema | Source |
|--------|--------|
| Teams app manifest 1.x → 2.4 | `@microsoft/app-manifest` (typed wrappers + JSON schemas under `dist/`) |
| Declarative Agent (`declarativeAgent.json`) | `@microsoft/app-manifest` v2.4+ |
| API plugin (`aiPlugin.json`) | `@microsoft/app-manifest` v2.4+ — note `mcp_tool_description` must be `{ file: "..." }`, `auth` always required, `namespace` required |

See [entities/teamsapp-manifest.md](../entities/teamsapp-manifest.md).

## Lifecycle YAML

`m365agents.yml` is parsed by `lifecycle/parser.ts`. There is no published JSON schema today; the shape is enforced by Zod schemas inside drivers (one schema per driver). Future work: emit a unified JSON schema for editor support.

See [entities/m365agents-yml.md](../entities/m365agents-yml.md).

## Env files

`.env.{envName}` files use standard `KEY=VALUE` syntax — no schema. Convention: prefix-based (`TEAMS_APP_*`, `BOT_*`, `AAD_APP_*`, `AZURE_*`, `SECRET_*`). See [entities/env-files.md](../entities/env-files.md).

## OpenAPI specs (for API plugins)

Parsed and validated by `packages/core-next/src/specParser/` (inline) and `packages/spec-parser` (standalone). Supports OpenAPI 3.0 (via `@apidevtools/swagger-parser`) and OpenAPI 2.0 → 3.0 conversion (via `swagger2openapi 7.0.8` — exact pin).

| Validator | Project type |
|-----------|--------------|
| `CopilotValidator` | Copilot plugins |
| `SMEValidator` | Search-based message extensions |
| `TeamsAIValidator` | Teams AI Library function-calling |

## Bicep

`infra/azure.bicep` follows Azure's standard Bicep schema (validated by `az bicep build` / `arm/deploy` driver). The toolkit does not impose extra constraints.

## CLI command schema

`atk` does not publish an OpenAPI/JSON command schema today. Commands are introspectable via `--help` and the `TemplateRegistry` (for `atk new`). Future work: emit `manifest.json` for editor IntelliSense.
