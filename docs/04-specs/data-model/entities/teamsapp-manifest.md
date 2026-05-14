# Teams app manifest

The Microsoft 365 / Teams app manifest is the contract between the toolkit's output and the platform. The toolkit **never edits raw JSON** — all manipulation goes through `@microsoft/app-manifest`'s `TeamsManifestWrapper`.

## Schema versions

| Version | Notes |
|---------|-------|
| 1.x | Legacy Teams manifests |
| 1.16+ | Adds `composeExtensions`, `bots`, `staticTabs`, `configurableTabs` |
| 1.17 | Adds `webApplicationInfo` (SSO) |
| 2.x | Microsoft 365 unified app schema |
| **2.4** | Required for declarative agents · `RemoteMCPServer` runtime · `RuntimeAuthenticationObject` (auth always required) |

## v2.4 specifics

The v2.4 typed converter (`APIPluginManifestWrapper`) enforces:

- `mcp_tool_description` must be `{ file: "mcp-tools.json" }` (object), not a bare string.
- `auth` is always required, even `{ type: "None" }` for unauthenticated APIs (`RuntimeAuthenticationObject` is non-optional in v2.4).
- `namespace` is required in the top-level manifest.
- `schema_version` must be `"v2.4"` to enable `RemoteMCPServer`.

See [`fx-core.instructions.md`](../../../.github/instructions/fx-core.instructions.md) for inline notes.

## Wrapper API

`TeamsManifestWrapper` exposes typed read/write per capability section:

| Section | Wrapper method |
|---------|---------------|
| `bots[]` | `getBots()`, `addBot()`, `updateBot(id, patch)` |
| `composeExtensions[]` | `getComposeExtensions()`, `addComposeExtension()` |
| `webApplicationInfo` | `getWebApplicationInfo()`, `setWebApplicationInfo()` |
| `declarativeCopilots[]` | `getDeclarativeCopilots()`, `addDeclarativeCopilot()` |
| `validDomains` | `getValidDomains()`, `addValidDomain()` |

Drivers consume the wrapper, not raw JSON. This means a future schema bump (2.5, 3.x) propagates through one library upgrade rather than a sweep of the codebase.

## Where it lives

- Authored at: `appPackage/manifest.json`
- Mustache-rendered at scaffold time using values from the question tree.
- Re-rendered with env-resolved placeholders during `teamsApp/zipAppPackage`.

## Validation

| Driver | Layer |
|--------|-------|
| `teamsApp/validateManifest` | Schema validation via `TeamsManifestWrapper` (offline) |
| `teamsApp/validateAppPackage` | Full package validation via TDP API (online) |

## DA-specific manifest files

For Declarative Agent templates, the manifest is paired with:

- `appPackage/declarativeAgent.json` — the DA definition.
- `appPackage/aiPlugin.json` — the API plugin definition (when applicable).

Both are typed by `@microsoft/app-manifest` and manipulated through the same wrapper layer.
