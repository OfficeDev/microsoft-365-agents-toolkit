# Declarative Agent Lifecycle with WIQD

Use WIQD for the complete lifecycle of pure DAs and DAs that attach existing OpenAPI, remote MCP, or Copilot Connector backends. Do not fall back to ATK when a WIQD lifecycle command fails.

Use ATK instead for the complete lifecycle of a DA that contains project-owned backend source code and lifecycle actions that provision or deploy that backend. The `declarative-agent-action*` templates create this shape. Do not switch it to WIQD.

Also preserve ATK for TypeSpec DA projects, which WIQD does not expose. This explicit exception is not evidence that every ATK-scaffolded DA owns backend deployment. WIQD supports referencing existing Copilot Connector connections, while a project that builds and deploys its own Connector backend remains on ATK.

## Structural Routing Gate

Before any lifecycle command, inspect the project capability and structure. Creator identity is not a reliable signal because WIQD delegates operations to ATK/fx-core.

1. Confirm the Teams app manifest contains `copilotAgents.declarativeAgents`.
2. Preserve ATK when the project explicitly identifies a TypeSpec DA capability.
3. Otherwise, check for project-owned backend source, such as `src/functions`, together with its build configuration.
4. Check `m365agents*.yml` for lifecycle actions that provision or deploy that backend, such as `arm/deploy`, `azureFunctions/zipDeploy`, or `azureAppService/zipDeploy`.

Select ATK based on backend ownership only when both project-owned backend source and corresponding backend deployment lifecycle are present. A DA action or plugin that points to an existing OpenAPI service or remote MCP server stays on WIQD. `m365agents.yml` alone does not select ATK. If an unusual project does not provide enough evidence, ask whether the backend is owned and deployed by this project before running lifecycle commands.

## Detect a Declarative Agent

Treat the request or project as a DA when any of these conditions is true:

- The user explicitly says "Declarative Agent" or "DA" in a Declarative Agent context.
- `appPackage/declarativeAgent.json` exists.
- `appPackage/manifest.json` contains `copilotAgents.declarativeAgents`.

The presence of `m365agents.yml` does not make a project non-DA. Check the DA markers before selecting a lifecycle CLI.

## Read-Only Reference Requests

Answer questions about DA schemas, manifest fields, capabilities, examples, or project structure from the local references without requiring WIQD installation or authentication. Require WIQD only when executing a DA lifecycle operation routed to WIQD.

## WIQD Setup

Before the first DA lifecycle command, run:

```bash
wiqd --version
```

If WIQD is unavailable, stop the DA lifecycle workflow and direct the user to the
[official WIQD installation guide](https://microsoft.github.io/wiqd/getting-started/installation/)
for the current installation instructions. Do not duplicate platform-specific installer commands or use ATK as a fallback.

Use these diagnostics when needed:

```bash
wiqd auth status
wiqd auth login --interactive
wiqd doctor
```

Do not require login until the requested operation needs Microsoft 365 access.

## Lifecycle Commands

Use this mapping for pure DAs and DAs backed by existing OpenAPI, remote MCP, or Copilot Connector services. Do not apply it to DAs with project-owned backend deployment or TypeSpec DAs.

| DA intent or former ATK command | WIQD command                                                                                         |
| ------------------------------- | ---------------------------------------------------------------------------------------------------- |
| `atk new`                       | `wiqd agent create --name <name> --output <parent>`                                                  |
| `atk validate`                  | `wiqd agent validate --path <project> --env <env>`                                                   |
| `atk package`                   | `wiqd agent package --path <project> --env <env>`                                                    |
| `atk provision`                 | `wiqd agent provision --path <project> --env <env>`                                                  |
| Share tenant-wide               | `wiqd agent share --path <project> --env <env> --scope tenant`                                       |
| Share with users                | `wiqd agent share --path <project> --env <env> --scope users --email <comma-separated-emails>`       |
| Remove sharing                  | `wiqd agent share remove --path <project> --env <env> --users <comma-separated-emails>`               |
| `atk publish`                   | `wiqd agent publish --path <project> --env <env>`                                                    |
| `atk uninstall`                 | `wiqd agent delete --path <project> --env <env>`                                                     |
| `atk auth list`                 | `wiqd auth status`                                                                                   |
| `atk auth login m365`           | `wiqd auth login --interactive`                                                                      |
| `atk doctor`                    | `wiqd doctor`                                                                                        |

Do not execute `wiqd agent publish` unless the user explicitly asks to publish and confirms the target.

For a pure DA with no backend compute, use this sequence:

```text
validate -> package -> provision -> share or publish
```

Do not run `atk deploy` for a pure DA.

## OpenAPI Actions

Inspect the OpenAPI document and ask the user to resolve any missing operation selection. Pass operations as one comma-separated value:

```bash
wiqd agent add action \
  --folder <project-directory> \
  --openapi-spec <path-or-url> \
  --operations "GET /resource,POST /resource"
```

Do not add ATK-only flags such as `--api-plugin-type`, `--openapi-spec-type`, `--openapi-spec-location`, `--api-operation`, or `-i false` to WIQD commands.

## MCP Actions

No authentication:

```bash
wiqd agent add action --folder <project> --mcp-server-url <url> --mcp-auth-type none
```

Dynamic OAuth:

```bash
wiqd agent add action --folder <project> --mcp-server-url <url> --mcp-auth-type oauth-dynamic
```

Static OAuth:

```bash
wiqd agent add action \
  --folder <project> \
  --mcp-server-url <url> \
  --mcp-auth-type oauth \
  --mcp-client-id <id> \
  --mcp-client-secret <secret> \
  --mcp-scopes <space-separated-scopes>
```

Entra SSO:

```bash
wiqd agent add action \
  --folder <project> \
  --mcp-server-url <url> \
  --mcp-auth-type entra-sso \
  --mcp-client-id <id>
```

For static OAuth, require a real client ID and client secret; scopes are optional and must come from the user or provider documentation. For Entra SSO, require a real client ID. Never invent client IDs, client secrets, or scopes.

## DA Backend Routing

Classify the project by structure, then keep one toolchain for its complete lifecycle:

| Scenario | Routing basis | Lifecycle |
| --- | --- | --- |
| Pure DA | No project-owned backend deployment | WIQD |
| DA that references an existing Copilot Connector connection | Existing connection | WIQD |
| DA that attaches an existing OpenAPI API or remote MCP server | Existing external backend | WIQD for the DA; the existing backend keeps its own deployment toolchain |
| DA with project-owned backend source and corresponding backend deployment actions, including a new Connector ingestion backend | Project-owned backend deployment | ATK for the DA manifest, backend, provisioning, deployment, packaging, publishing, and deletion |
| TypeSpec DA | Explicit TypeSpec capability | Follow the generated ATK lifecycle |

Do not switch tools after a lifecycle failure. Apply the structural routing gate above to every existing project.
