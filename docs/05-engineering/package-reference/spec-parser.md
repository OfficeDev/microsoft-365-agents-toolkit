# `packages/spec-parser` — `@microsoft/m365-spec-parser`

OpenAPI parser, validator, filter, and optimizer used by API-plugin templates.

## Role

Wraps two upstream libraries with a Result-oriented API:

- `@apidevtools/swagger-parser ^10.1.1` — parses OpenAPI 3.x.
- `swagger2openapi 7.0.8` (exact pin) — converts OpenAPI 2.x to 3.x.

## Public API

| Function | Purpose |
|----------|---------|
| `parseSpec(uri)` | Returns `Result<ParsedSpec>`; resolves env vars; checks circular refs |
| `filterSpec(spec, operations)` | Keeps only selected operations; runs optimizer |
| `optimizeSpec(spec)` | Removes unused components, tags, security schemes, vendor extensions |
| `createValidator(projectType)` | Returns `CopilotValidator`, `SMEValidator`, or `TeamsAIValidator` |

## Project types

| `ProjectType` | Validator | Used by |
|--------------|-----------|---------|
| `Copilot` | `CopilotValidator` | Copilot plugin templates |
| `SME` | `SMEValidator` | Search-based message extensions |
| `TeamsAI` | `TeamsAIValidator` | Teams AI Library function calling |

## Inline copy in `core-next`

The same module exists inline at `packages/core-next/src/specParser/` (merged from this package). The standalone `spec-parser` continues to ship for v3 consumers; the inline copy is what v4 templates use.

## Bundler

Rollup — library shape (CJS + ESM + d.ts).

## Why pin `swagger2openapi`

Newer versions break the `RemoteMCPServer` runtime type used in v2.4 manifests. Pinned at exactly `7.0.8` until a fix lands upstream. Tracked as a risk in [02-architecture/11-risks.md](../../02-architecture/11-risks.md).
