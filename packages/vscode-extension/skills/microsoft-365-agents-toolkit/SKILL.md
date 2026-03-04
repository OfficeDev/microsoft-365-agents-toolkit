---
name: microsoft-365-agents-toolkit
description: Builds, tests, and deploys Microsoft 365 apps or agents working on Teams and Copilot. Used when working with Teams agent, bot, tab, message extension, Declarative Agents, Custom Engine Agents using Microsoft 365 Agents Toolkit. It helps to do local testing, Azure resource provision, remote deployment, or any Microsoft 365 application development workflow.
---

# Microsoft 365 Agents Toolkit Skill

Build Microsoft 365 agents and Teams apps using the ATK CLI.

## AI Behavior Guidelines

1. **Testing Strategy:** Recommend Agents Playground first (faster, no M365 needed). Use Teams workflow only if user explicitly requests it.

2. **Environment Variables:** NEVER hardcode secrets or make up placeholder values. Always ask users for real values.

3. **Error Handling:** Read error messages carefully. Check `env/.env.local`, `.localConfigs`, and `atk auth list`.

4. **Long-Running Commands — WAIT for completion:**
   - `atk new`, `atk provision`, `atk deploy` can take several minutes
   - Always wait for completion before running the next step (timeout 120000ms+)

5. **Local Service Startup — Hangs terminal (expected):**
   - `npm run dev`, `npm start`, `python app.py`, etc. will hang — the server keeps running
   - Run as a background process (`isBackground=true`)
   - Do NOT wait for it to "finish" — verify startup by checking output for "listening on port" or similar
   - If errors appear, read logs, diagnose, fix, restart
   - Use a **NEW terminal** to launch Agents Playground or open Teams sideloading URL

6. **Monitor App Logs:** Periodically check background terminal output for runtime errors. If the app crashes, read the error, fix the root cause, and restart.

## ATK CLI Setup

```bash
atk --version  # Must be > 1.1.5-alpha
```

If ATK is not found or version is too old:
```bash
npm i -g @microsoft/m365agentstoolkit-cli@alpha
```

## Start from a Template

ALWAYS scaffold from a template when building new Teams/M365 apps.

**Default project folder:** `~/AgentsToolkitProjects`

### Template Selection Guide

| User Wants | Capability |
|------------|------------|
| Extend M365 Copilot with custom instructions | `copilot-gpt-basic` |
| Declarative Agent with new API | `api-plugin-from-scratch` |
| Declarative Agent with existing OpenAPI spec | `api-plugin-from-existing-api` |
| Connect MCP Server to Copilot | `declarative-agent-with-action-from-mcp` |
| Agent with custom LLM (Azure OpenAI, etc.) | `basic-custom-engine-agent` |
| Agent using Azure AI Foundry | `foundry-agent` |
| Teams chatbot with AI | `custom-copilot-basic` |
| Teams bot with RAG/knowledge base | `custom-copilot-rag-customize` |
| Simple Teams echo bot | `default-bot` |
| Teams tab app | `non-sso-tab` |
| Teams message extension | `default-message-extension` |

See [templates.md](references/templates.md) for complete list with language support.

### Creating Projects

```bash
# Declarative Agent (no backend, no -l flag needed)
atk new -c copilot-gpt-basic -n my-agent -f ~/AgentsToolkitProjects -i false

# Declarative Agent with new API
atk new -c api-plugin-from-scratch -l typescript -n my-api-agent -f ~/AgentsToolkitProjects -i false

# Declarative Agent with existing OpenAPI spec (requires -a and -o with operation IDs)
atk new -c api-plugin-from-existing-api -n my-agent -a <openapi-spec-url-or-path> -o "GET /repairs" -o "POST /repairs" -f ~/AgentsToolkitProjects -i false

# Custom Engine Agent
atk new -c basic-custom-engine-agent -l typescript -n my-cea -f ~/AgentsToolkitProjects -i false

# Teams Agent with RAG
atk new -c custom-copilot-rag-customize -l typescript -n my-rag-agent -f ~/AgentsToolkitProjects -i false
```

**Notes:**
- `copilot-gpt-basic` does NOT require `-l` language flag
- `api-plugin-from-existing-api` requires `-a` (OpenAPI spec) and `-o` (operation IDs like `"GET /path"`)
- Always use `-i false` for non-interactive scripted creation

## Quick Reference

| Task | Command |
|------|---------|
| Install/Update ATK | `npm i -g @microsoft/m365agentstoolkit-cli@alpha` |
| Install Agents Playground | `winget install agentsplayground` |
| Login M365 / Azure | `atk auth login m365` / `atk auth login azure` |
| Check login | `atk auth list` |
| **Test in Playground** | `agentsplayground -e http://localhost:3978/api/messages -c msteams` |
| Provision local | `atk provision --env local -i false` |
| Deploy local | `atk deploy --env local -i false` |
| Provision cloud | `atk provision --env dev --resource-group <rg> --region <region> -i false` |
| Deploy cloud | `atk deploy --env dev -i false` |
| Open Teams app | `https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}` |
| Add API action | `atk add action --api-plugin-type api-spec --openapi-spec-type enter-url-or-open-local-file -a <spec> -i false` |

## Local Testing

### Option 1: Agents Playground (Recommended)

Faster, no M365 account or provisioning needed.

```bash
# Install (Windows)
winget install agentsplayground

# Start bot service (background — hangs terminal, expected)
npm run dev

# NEW terminal: launch playground
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

Verify bot started successfully by checking terminal output before launching playground.

### Option 2: Run on Teams

For bot projects, start devtunnel first:
```bash
devtunnel host -p 3978 --allow-anonymous
# Set BOT_ENDPOINT in env/.env.local with the tunnel URL
```

Then provision, deploy, start service, and open Teams:
```bash
atk provision --env local -i false
atk deploy --env local -i false

# Start service (background — hangs terminal, expected)
# Check package.json for start command: npm run dev:teamsfx, npm run dev, npm start, etc.

# NEW terminal: open Teams sideloading URL
# https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}
```

**For declarative agents (no backend service):** Just provision/deploy and open the M365 Copilot sideloading URL.

### Opening Declarative Agents in M365 Copilot

Declarative agents use `M365_APP_ID` (not `TEAMS_APP_ID`), acquired after `teamsApp/extendToM365` runs during provisioning.

**Sideloading URL format:**
```
https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&developerMode=Basic
```

Where `${agent-hint}` is Base64-encoded JSON:
```json
{"id": "${M365_APP_ID}", "scenario": "launchcopilotextension", "properties": {"clickTimestamp": "2/6/2026, 10:30:45 AM"}, "version": 1}
```

See [localdebug.md](references/localdebug.md) for complete local testing details.

## Key Project Files

| File | Purpose |
|------|---------|
| `appPackage/manifest.json` | App metadata and capabilities |
| `appPackage/declarativeAgent.json` | Agent instructions (Declarative Agents) |
| `env/.env.local` | Local environment variables |
| `env/.env.dev` | Dev/cloud environment variables |
| `.localConfigs` | Runtime config generated by `atk deploy --env local` |
| `m365agents.yml` | Lifecycle config for dev/cloud deployment |
| `m365agents.local.yml` | Lifecycle config for local development |

**Critical:** Backend service reads from `.localConfigs`, NOT from `env/.env.local`. See [environment-config.md](references/environment-config.md) for details.

## Troubleshooting

**Port already in use:**
```powershell
Get-NetTCPConnection -LocalPort 3978 | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

**Missing environment variables at runtime:** Check `.localConfigs` exists and has required values. Run `atk deploy --env local -i false` to regenerate.

**App not loading:** Verify `M365_APP_ID` (for declarative agents) or `TEAMS_APP_ID` (for bots/tabs) exists in `env/.env.local`.

**General diagnostics:** `atk doctor`, `atk validate --env <env>`, `atk auth list`

## Detailed References

- **[templates.md](references/templates.md)** — All templates with descriptions and language support
- **[localdebug.md](references/localdebug.md)** — Local testing with Agents Playground and Teams
- **[environment-config.md](references/environment-config.md)** — YAML lifecycle files, .localConfigs, Azure OpenAI config, cloud deployment
- **[commands.md](references/commands.md)** — Package, validate, share, collaborate, environment management, adding actions
