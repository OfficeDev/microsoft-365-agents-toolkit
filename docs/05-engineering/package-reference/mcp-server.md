# `packages/mcp-server` — `@microsoft/m365-mcp-server`

A [Model Context Protocol](https://modelcontextprotocol.io) server exposing toolkit operations as MCP tools.

## Role

Lets MCP-aware clients (e.g. AI agents, IDE chat panels) drive scaffold / provision / deploy / publish via the standard MCP protocol instead of CLI invocations.

## Tools exposed

Mirrors the high-level surface of the CLI:

- `create_project`
- `provision`
- `deploy`
- `publish`
- `list_templates`
- `validate_manifest`
- `package_app`

(Exact tool names live in the package source.)

## Bundler

Webpack (legacy). Migration tracked but not on the immediate roadmap.

## Status

Independent release cadence. Consumed by experimental client integrations.
