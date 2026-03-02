---
name: foundry-agent-to-m365
description: Onboard an existing Azure AI Foundry agent to Microsoft 365 with Microsoft 365 Agents Toolkit CLI (atk). Use when users already have a deployed Foundry agent and need to scaffold a foundry-agent project, configure environment values, run local debug, provision/deploy, and validate in Teams or Microsoft 365 Copilot.
---

# Azure Foundry Agent to Microsoft 365 Skill

Use this skill only for connecting an existing Foundry agent to Microsoft 365.

## AI Behavior Guidelines

1. **Stay in scope:** Handle only existing Foundry agent onboarding to Microsoft 365.
2. **Use one capability:** Use only `foundry-agent` in `atk new`; do not suggest other templates.
3. **Never invent values:** Ask user for real endpoint, IDs, tenant/subscription, secrets.
4. **Run non-interactive commands:** Use `-i false` for `atk` operations.
5. **Diagnose before retry:** Read command output first, then check auth and env files.

## ATK CLI Setup

```bash
atk --version
```

Require `atk` newer than `1.1.5-alpha`.

If not installed or outdated:

```bash
npm i -g @microsoft/m365agentstoolkit-cli@alpha
```

## Required Workflow

### 1. Bootstrap Project

Ask for project name and folder (default current folder), then run:

`atk new -c foundry-agent -n <project-name> -f <project-folder> -i false`

Switch to the new project directory before running next commands:

`cd <project-folder>`

Open the folder if it's in VS Code.

### 2. Configure Environment

- Update generated env files (`env/.env.dev`, optionally `env/.env.local`) with existing Foundry settings.
- Keep tenant, subscription, resource group, and endpoint aligned with the deployed Foundry agent.
- Do not continue until required values are present.

### 3. Local Debug

#### Option 1: Agents Playground (Recommended)
Use this first when possible (fastest path, no M365 provisioning).

```bash
# Install (Windows)
winget install agentsplayground
# Start bot service (background — hangs terminal, expected)
npm run dev:teamsfx:playground
# NEW terminal: launch playground
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

Confirm bot service starts successfully before launching playground.

#### Option 2: Run on Teams
Use this when Teams sideload validation is required.

Ensure Microsoft 365 login and sideloading are enabled:

```bash
atk auth list # If needed: atk auth login m365
```

Start dev tunnel first in a separate terminal:

```bash
devtunnel host -p 3978 --allow-anonymous
# Set BOT_ENDPOINT in env/.env.local with the tunnel URL
```

Then run local provision/deploy, start service, and open Teams URL:

```bash
atk provision --env local -i false
atk deploy --env local -i false

# Start service (background — hangs terminal, expected)
# Check package.json for start command: npm run dev:teamsfx, npm run dev, npm start, etc.

# NEW terminal: open Teams sideloading URL
# https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}
```

### 4. Provision and Deploy

Ensure Azure login:

```bash
atk auth list # If needed: atk auth login azure
```

```bash
atk provision --env dev --resource-group <rg> --region <region> -i false
atk deploy --env dev -i false
```

### 5. Validate in Microsoft 365

- Open Microsoft 365 Copilot or Teams and run a prompt expected to hit the connected Foundry agent.
- If behavior is incorrect, re-check env values, then redeploy.

## Quick Reference

| Task               | Command                                                                    |
| ------------------ | -------------------------------------------------------------------------- |
| Install/Update ATK | `npm i -g @microsoft/m365agentstoolkit-cli@alpha`                          |
| Login              | `atk auth login m365` / `atk auth login azure`                             |
| Verify auth        | `atk auth list`                                                            |
| Bootstrap project  | `atk new -c foundry-agent -n <project-name> -f <project-folder> -i false`  |
| Local debug        | `atk provision --env local -i false` / `atk deploy --env local -i false`   |
| Provision          | `atk provision --env dev --resource-group <rg> --region <region> -i false` |
| Deploy             | `atk deploy --env dev -i false`                                            |

## Key Project Files

| File                       | Purpose                                   |
| -------------------------- | ----------------------------------------- |
| `appPackage/manifest.json` | Microsoft 365 app metadata                |
| `env/.env.dev`             | Remote environment configuration          |
| `env/.env.local`           | Local environment configuration (if used) |
| `m365agents.yml`           | Remote lifecycle config                   |
| `m365agents.local.yml`     | Local lifecycle config                    |
| `.localConfigs`            | Generated local runtime config            |

## Troubleshooting

**Authentication**

- Run `atk auth list`; if needed, re-run login commands.

**Provision/deploy failures**

- Re-check resource group/region/tenant/subscription context.
- Confirm environment variables map to the already deployed Foundry agent.

**Port already in use:**

```powershell
Get-NetTCPConnection -LocalPort 3978 | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

**General diagnostics**

- `atk doctor`
- `atk validate --env dev -i false`

## References

- Read **[environment-config.md](references/environment-config.md)** when updating env files or lifecycle settings.
- Read **[localdebug.md](references/localdebug.md)** when troubleshooting local debug paths.
