# Local Debugging Guide

Use this reference to choose and run local debug paths for `foundry-agent` projects.

## Decision Rule

- Use **Agents Playground** first for fastest validation.
- Use **Teams direct launch** when user needs a more realistic environment and have Microsoft 365 account sideloading permission.

## Path A: Agents Playground (Preferred)

### Preconditions

1. Local service start command is known.
2. Service listens on `http://localhost:3978/api/messages` (or equivalent project endpoint).

### Steps

```bash
# install (Windows)
winget install agentsplayground

# terminal 1: start local service (long-running)
npm run dev:teamsfx:playground

# terminal 2: launch playground
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

### Success Checks

- Service terminal shows successful startup (listening/started log).
- Playground opens and returns bot responses.

## Path B: Teams Direct Launch

### Preconditions

1. Microsoft 365 sign-in is valid: `atk auth list`.
2. For bot projects, dev tunnel is running and public.
3. `BOT_ENDPOINT` points to current tunnel URL.

### Steps (Bot Projects)

```bash
# terminal 1: start dev tunnel first
devtunnel host -p 3978 --allow-anonymous

# update env/.env.local with BOT_ENDPOINT=<tunnel-url>

# provision and deploy local resources
atk provision --env local -i false
atk deploy --env local -i false

# terminal 2: start local service (long-running)
# use project-appropriate script, e.g. npm run dev:teamsfx

# terminal 3: open Teams sideload URL
# https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}
```

## Troubleshooting

If service does not respond:

1. Confirm local service is still running.
2. Confirm endpoint and port match playground/tunnel configuration.
3. For Teams path, confirm tunnel URL is current and mapped to `BOT_ENDPOINT`.
4. Re-run `atk deploy --env local -i false` after env changes.

If port `3978` is in use (Windows):

```powershell
Get-NetTCPConnection -LocalPort 3978 | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

Useful diagnostics:

```bash
atk auth list
atk doctor
atk validate --env local -i false
agentsplayground --help
```
