# Open Plugin Converter (`atk convert openplugin`)

Converts an [Open Plugin Spec v1.0](https://open-plugins.com/) plugin directory
into a scaffolded Microsoft 365 Agents Toolkit project with a devPreview manifest
containing `agentSkills` and `agentConnectors`.

Accepts plugins using any of the three manifest locations:
- `.plugin/plugin.json` (vendor-neutral, recommended)
- `.claude-plugin/plugin.json` (Claude Code)
- `.cursor-plugin/plugin.json` (Cursor)

## Usage

```bash
# Minimal — skills + MCP servers, with required developer URLs
atk convert openplugin \
  --path ./my-plugin \
  --privacy-url https://contoso.com/privacy \
  --terms-url https://contoso.com/terms

# Explicit output directory and website URL override
atk convert openplugin \
  --path ./my-plugin \
  --output ./converted-project \
  --privacy-url https://contoso.com/privacy \
  --terms-url https://contoso.com/terms \
  --website-url https://contoso.com

# Override auth type for all MCP connectors
atk convert openplugin \
  --path ./my-plugin \
  --privacy-url https://contoso.com/privacy \
  --terms-url https://contoso.com/terms \
  --default-auth-type None

# Provide a specific app ID
atk convert openplugin \
  --path ./my-plugin \
  --privacy-url https://contoso.com/privacy \
  --terms-url https://contoso.com/terms \
  --app-id 00000000-0000-0000-0000-000000000001
```

After conversion, build and validate the package:

```bash
cd ./my-plugin          # or --output path
atk teamsapp package
atk teamsapp validate --package-file ./appPackage/build/appPackage.dev.zip
```

## CLI options

| Flag | Required | Description |
|---|---|---|
| `--path / -p` | yes | Path to the Open Plugin directory. |
| `--output / -o` | no | Destination project folder. Defaults to `./<plugin-name>`. |
| `--privacy-url` | yes | `developer.privacyUrl` — the Open Plugin spec has no equivalent. |
| `--terms-url` | yes | `developer.termsOfUseUrl` — the Open Plugin spec has no equivalent. |
| `--website-url` | no | `developer.websiteUrl`. Falls back to `plugin.json` `homepage` then `author.url`. |
| `--app-id` | no | Override the deterministic UUIDv5 manifest id. |
| `--default-auth-type` | no | `Auto` (default), `None`, `OAuthPluginVault`, or `ApiKeyPluginVault`. |

## What gets mapped

| Open Plugin component | Manifest field | Notes |
|---|---|---|
| `skills/<name>/SKILL.md` | `agentSkills[].folder` | Copied verbatim; sorted alphabetically. |
| `.mcp.json` HTTP servers | `agentConnectors[].toolSource.remoteMcpServer` | Auth auto-detected: HTTPS non-localhost → OAuthPluginVault, else None. |
| `.mcp.json` stdio servers | *(skipped)* | Warning emitted; requires manual `localMcpServer` setup. |
| `commands/*.md` | *(copied alongside, inert)* | Not yet in MOS3 manifest; kept for forward compatibility. |
| `hooks/`, `agents/`, `rules/`, `lspServers/`, `outputStyles/` | *(dropped)* | Warning emitted per field. Not representable in MOS3 today. |

## Module structure

```
openPlugin/
  types.ts            # TypeScript interfaces for parsed plugin data and inputs
  parser.ts           # Reads plugin dir: manifest probe, .mcp.json, skills/, commands/
  authorParser.ts     # Parses author field (object or "Name <email> (url)" string)
  textUtils.ts        # Word-boundary truncation, kebab-to-title-case
  deterministicId.ts  # UUIDv5 (SHA-1) for stable manifest id generation
  mapper.ts           # Pure transform: parsed plugin → devPreview manifest + copy operations
  iconStrategy.ts     # Resolves color.png / outline.png from plugin icons or logo field
  placeholderPng.ts   # Generates solid-color RGB PNGs using Node zlib (no native deps)
  generator.ts        # Orchestrator: parse → map → scaffold project tree (baseline files come from the open-plugin-convert template)
```

## Feature flags

| Flag | Default | Purpose |
|---|---|---|
| `TEAMSFX_OPENPLUGIN_CONVERT` | `true` | Gates CLI command registration. |
| `TEAMSFX_AGENT_SKILLS` | `true` | Gates `agentSkills` emission and `createAppPackage` folder walk. |
