# `packages/manifest` — `@microsoft/app-manifest`

Typed wrapper layer over the Microsoft 365 / Teams app manifest.

## Role

- Exposes `TeamsManifestWrapper` — typed read/write per capability section.
- Exposes `APIPluginManifestWrapper` — typed read/write for API plugins (v2.4+).
- Bundles JSON schemas for every supported manifest version.
- Ships converters for schema upgrades (e.g. 1.x → 2.x).

## Why a wrapper

Engine code never edits raw manifest JSON. Schema-version churn (1.x → 2.x → 2.4 → next) propagates through one library upgrade rather than a sweep of every driver / generator.

## Schema versions

| Version | Notes |
|---------|-------|
| 1.x | Legacy Teams |
| 2.x | Unified M365 schema |
| **2.4** | DA + RemoteMCPServer + RuntimeAuthenticationObject (auth always required) + `namespace` required |

## v2.4 specifics

The `APIPluginManifestWrapper` enforces:

- `mcp_tool_description` must be `{ file: "mcp-tools.json" }` (object), not bare string.
- `auth` is **always** required (even `{ type: "None" }`).
- `namespace` is required at the top level.
- `schema_version` must be `"v2.4"` to enable `RemoteMCPServer`.

## Consumers

- `fx-core` generators / drivers (v3).
- `core-next` `declarativeAgent/` module + `teamsApp/*` drivers (v4).

Manipulating manifest from any other code path is out-of-pattern.
