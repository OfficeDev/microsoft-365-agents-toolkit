# Add MCP Action To Declarative Agent (draft)

## Metadata

- Created: 2026-05-20T00:00:00Z
- Last updated: 2026-05-20T00:00:00Z
- PM owner: summzhan
- Engineer owner: HuihuiWu-Microsoft, Alive-Fish
- Scenario group: da
- Scenario ID: SCN-DA-ADD-MCP-ACTION-TO-DA
- Visual/state reference: add-mcp-action-to-da.html

> **Draft note:** This draft redesigns the live [`../add-mcp-action-to-da.md`](../add-mcp-action-to-da.md). VS Code Add action now runs the full flow end-to-end (URL &rarr; auth type &rarr; conditional follow-up &rarr; write action manifest + DA manifest + yml), the same way CLI does. Static tool fetching is dropped; tool discovery is dynamic at runtime. The separate VS Code fetch-from-CodeLens scenario [`../fetch-mcp-tools.md`](../fetch-mcp-tools.md) is absorbed into this flow and will be archived when the redesign ships. Companion redesign: [`create-da-with-mcp-server.md`](create-da-with-mcp-server.md).

## Scenario

A developer has an existing Declarative Agent project and wants to wire a Microsoft 365 Copilot action that calls an MCP server. Both surfaces ask the same questions in the same order: MCP server URL, authentication type, and at most a couple of follow-up fields determined by the chosen authentication type. When the user confirms, the toolkit creates a new action manifest, updates the declarative agent manifest, and updates `m365agents.yml` to wire OAuth provisioning when the authentication type requires it. No static tool list is captured; the agent host discovers the MCP server's tools at runtime.

Success means that after the flow completes, the project's declarative agent manifest references a new MCP-backed action, the action manifest references the MCP server URL, and `m365agents.yml` has the OAuth provisioning step injected when the authentication type requires it. For `OAuth (with static registration)`, env files (`env/.env.dev` and `env/.env.dev.user`) are populated with placeholders for the developer-provided client id, client secret, and optional scopes.

## Dependencies

- Requires: an existing Declarative Agent project with `appPackage/manifest.json` and a declarative agent manifest.
- VS Code precondition: the project is open, and the MCP-for-DA preview is enabled so `Start with a MCP server` shows up in the action-type picker.
- CLI precondition: the DA project path is supplied with `-f` / `--folder` or is the current project folder, and the manifest path is supplied with `-t` / `--manifest-file` or defaults to `./appPackage/manifest.json`.
- Produces: an updated declarative agent manifest, a new action manifest (`ai-plugin.json`), and an updated `m365agents.yml`. For `OAuth (with static registration)`, also updated env files with placeholders for the developer-provided credentials.
- Does not produce: a static MCP tools JSON file. Tool discovery is deferred to the agent host at runtime.

## Surfaces

- VS Code: the `Add action` command (tree view, Command Palette, or right-click on the DA project) runs the full add-action flow. The user picks `Start with a MCP server`, enters the server URL, picks an authentication type, fills any conditional follow-up fields, and the toolkit writes the action manifest, DA manifest update, and yml wiring in one shot. The legacy `.vscode/mcp.json`-then-CodeLens-fetch handoff is gone.
- CLI interactive: prompt-driven `atk add action` behavior. Same questions in the same order: action type, MCP server URL, authentication type, conditional follow-up fields, Teams manifest path.
- CLI non-interactive: flag-driven `atk add action` behavior. Requires `--api-plugin-type mcp`, `--mcp-da-server-url`, `--mcp-da-auth-type`, `--manifest-file`, `--folder`, and `--interactive false`. Conditional follow-up flags are required only when the chosen authentication type asks for them.
- Visual Studio and chat: not covered by this draft scenario.

## States

- Entry: an existing DA project is available.
- Action-type pick (VS Code and CLI interactive): single-select titled `Add an Action`. When the MCP-for-DA preview is enabled the list contains `Start with an OpenAPI Description Document` and `Start with a MCP server`. The user picks `Start with a MCP server`.
- Server URL input: text input titled `MCP Server URL` with placeholder `Enter your MCP server URL(e.g. https://example-mcp.com)`. Required.
- Authentication type pick: single-select titled `Select Authentication Type` with four options shown in this order:
  - `OAuth (with dynamic client registration)` &mdash; the MCP server registers a client at runtime; no follow-up fields.
  - `OAuth (with static registration)` &mdash; follow-up: required `Client ID`, required `Client Secret`, optional `Scopes` (space-separated).
  - `Entra SSO` &mdash; follow-up: required `Client ID` of an Entra application that will be used for provisioning. No tenant id or scopes are asked here.
  - `None` &mdash; no follow-up.
- Manifest pick (CLI only when needed): select Teams `manifest.json` file. VS Code uses the project's `appPackage/manifest.json` without asking.
- Write step (all auth types): the toolkit creates the action manifest (`ai-plugin.json`) with the MCP server URL, updates the declarative agent manifest to reference the new action, and updates `m365agents.yml` with the standard MCP action wiring.
- Write step (`OAuth (with static registration)`): in addition, `m365agents.yml` includes an `oauth/register` step, and `env/.env.dev` plus `env/.env.dev.user` contain placeholders for the client id, client secret, and scopes the developer entered. The secret is written into `env/.env.dev.user` and is not committed.
- Write step (`Entra SSO`): in addition, `m365agents.yml` references the developer-provided Entra application client id for the SSO action.
- Success notification: VS Code shows an info notification (`Action added to your Declarative Agent`); CLI prints a success message that names the updated DA manifest and the new action manifest.
- Recoverable error: missing MCP server URL, missing required auth follow-up field (static OAuth client id/secret, Entra SSO client id), invalid manifest path (CLI), or invalid project path is shown with a same-flow recovery path.
- Cancellation: the user can cancel at any pick or input; cancellation must not leave any half-written file in the project.

## Flow

### VS Code add action flow

```mermaid
flowchart TD
  ProjectReady([Existing DA project is open]) --> RunAdd[Run Add action command]
  RunAdd --> PickType[Single-select 'Add an Action': pick 'Start with a MCP server']
  PickType --> EnterUrl[Text input 'MCP Server URL']
  EnterUrl --> UrlValid{URL provided?}
  UrlValid -- No --> MissingUrl[Show 'MCP Server URL is required' error and stay on the input]
  UrlValid -- Yes --> SelectAuth["Select Authentication Type: OAuth (dynamic) / OAuth (static) / Entra SSO / None"]
  SelectAuth -- "OAuth (dynamic)" --> WriteFiles
  SelectAuth -- None --> WriteFiles
  SelectAuth -- "OAuth (static)" --> StaticFields["Enter Client ID, Client Secret, optional Scopes"]
  SelectAuth -- "Entra SSO" --> SsoFields[Enter Entra app Client ID]
  StaticFields --> WriteFiles[Write action manifest, update DA manifest, update m365agents.yml, and inject OAuth env placeholders when needed]
  SsoFields --> WriteFiles
  WriteFiles --> Success([Show 'Action added' notification])
  PickType --> Cancel([Cancel without changing project])
  EnterUrl --> Cancel
  SelectAuth --> Cancel
  StaticFields --> Cancel
  SsoFields --> Cancel
```

### CLI interactive add-action flow

```mermaid
flowchart TD
  Start([Run atk add action in an existing DA project]) --> ChooseType[Choose Action type: mcp]
  ChooseType --> EnterUrl[Enter MCP Server URL]
  EnterUrl --> SelectAuth["Select Authentication Type: OAuth (dynamic) / OAuth (static) / Entra SSO / None"]
  SelectAuth -- "OAuth (dynamic)" --> SelectManifest
  SelectAuth -- None --> SelectManifest
  SelectAuth -- "OAuth (static)" --> StaticFields["Enter Client ID, Client Secret, optional Scopes"]
  SelectAuth -- "Entra SSO" --> SsoFields[Enter Entra app Client ID]
  StaticFields --> SelectManifest[Select Teams manifest.json file]
  SsoFields --> SelectManifest
  SelectManifest --> WriteFiles[Write action manifest, update DA manifest, update m365agents.yml, and inject OAuth env placeholders when needed]
  WriteFiles --> Complete([Show CLI add action success message])
  ChooseType --> Cancel([Cancel without changing project])
  EnterUrl --> Cancel
  SelectAuth --> Cancel
  StaticFields --> Cancel
  SsoFields --> Cancel
  SelectManifest --> Cancel
```

### CLI non-interactive add-action flow

```mermaid
flowchart TD
  Start([Run atk add action --interactive false]) --> ValidateFlags{Required flags present?}
  ValidateFlags -- No --> MissingOption[Return validation error, including missing MCP server URL or auth type]
  ValidateFlags -- Yes --> AuthSwitch{mcp-da-auth-type}
  AuthSwitch -- "oauth-dynamic / none" --> WriteFiles
  AuthSwitch -- "oauth (static)" --> StaticFlags{Static-OAuth follow-up flags present?}
  AuthSwitch -- entraSSO --> SsoFlags{Entra-SSO follow-up flags present?}
  StaticFlags -- No --> MissingStatic[Return validation error for missing client id or client secret]
  StaticFlags -- Yes --> WriteFiles
  SsoFlags -- No --> MissingSso[Return validation error for missing Entra app client id]
  SsoFlags -- Yes --> WriteFiles[Write action manifest, update DA manifest, update m365agents.yml, and inject OAuth env placeholders when needed]
  WriteFiles --> Complete([Command succeeds with action added])
```

Example non-interactive command (no auth):

```bash
atk add action --api-plugin-type mcp --mcp-da-server-url <server-url> --mcp-da-auth-type none \
  -t ./appPackage/manifest.json -f <project-path> --interactive false
```

Example non-interactive command (static OAuth):

```bash
atk add action --api-plugin-type mcp --mcp-da-server-url <server-url> --mcp-da-auth-type oauth \
  --mcp-da-client-id <client-id> --mcp-da-client-secret <client-secret> --mcp-da-scopes "<space-separated-scopes>" \
  -t ./appPackage/manifest.json -f <project-path> --interactive false
```

Example non-interactive command (Entra SSO):

```bash
atk add action --api-plugin-type mcp --mcp-da-server-url <server-url> --mcp-da-auth-type entraSSO \
  --mcp-da-entra-client-id <entra-app-client-id> \
  -t ./appPackage/manifest.json -f <project-path> --interactive false
```

Example non-interactive command (dynamic OAuth):

```bash
atk add action --api-plugin-type mcp --mcp-da-server-url <server-url> --mcp-da-auth-type oauth-dynamic \
  -t ./appPackage/manifest.json -f <project-path> --interactive false
```

## Validation notes

- VS Code UI test intent should trace to `SCN-DA-ADD-MCP-ACTION-TO-DA` and cover the `Add action` command, the `Add an Action` action-type pick (with and without the MCP-for-DA preview enabled), the `MCP Server URL` input including empty-input recovery, the new `Select Authentication Type` pick for each of the four options, and each conditional follow-up path. The success state is the `Action added` notification and the matching writes to action manifest, DA manifest, and `m365agents.yml`; static-OAuth additionally verifies env file placeholders.
- The legacy `.vscode/mcp.json`-then-CodeLens-fetch flow (live scenario `SCN-DA-FETCH-MCP-TOOLS`) is replaced by this end-to-end flow; UI tests for the old fetch CodeLens should retire when the redesign ships.
- CLI E2E test intent should trace to `SCN-DA-ADD-MCP-ACTION-TO-DA` for interactive and non-interactive `atk add action --api-plugin-type mcp` paths and cover each auth-type value (`oauth-dynamic`, `oauth`, `entraSSO`, `none`).
- CLI non-interactive validation should cover missing `--mcp-da-server-url`, missing `--mcp-da-auth-type`, missing static-OAuth follow-up flags (`--mcp-da-client-id`, `--mcp-da-client-secret`), missing Entra-SSO `--mcp-da-entra-client-id`, invalid manifest path, and invalid DA project path.
- The static MCP tools file flag (`--mcp-tools-file-path`) and the static "Select Operation(s) Copilot can interact with" pick are explicitly out of scope for this scenario in the new flow; tool discovery is dynamic at runtime.
