# Local Debugging Guide

## Contents
- Debugging Options (Playground vs Teams)
- Option 1: Agents Playground (installation, quick start, CLI options, features, limitations)
- Option 2: Teams Direct Launch (devtunnel, provisioning, sideloading URLs)
- Comparison table
- Configuration File (.m365agentsplayground.yml)
- Environment Variables
- Troubleshooting

Debug and test your agents locally using two options:

## Debugging Options

| Option | Use When |
|--------|----------|
| **Agents Playground** | Recommended first for testing (faster, no M365 needed) |
| **Teams Direct Launch** | User explicitly asks to run on Teams |

---

## Option 1: Agents Playground (Recommended)

**Recommend Agents Playground first for testing**, unless user explicitly asks to run on Teams.

The Microsoft 365 Agents Playground is a web-based sandbox for testing and debugging agents locally without requiring:
- Microsoft 365 Developer tenant
- Azure tunneling (ngrok/dev tunnel)
- HTTPS endpoint
- App registration

### Installation

**Windows:**
```powershell
winget install agentsplayground
```

**Linux:**
```bash
curl -LO https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/microsoft-365-agents-playground%400.2.23/agentsplayground-linux-x64.zip
unzip agentsplayground-linux-x64.zip agentsplayground
chmod +x agentsplayground
sudo mv agentsplayground /usr/local/bin/
```

**npm:**
```bash
npm install -g @microsoft/m365agentsplayground
```

### Quick Start

```bash
# 1. Start your bot service (this will HANG the terminal — expected!)
# Run as a background process since the server keeps running
cd my-bot
npm run dev

# 2. Use a NEW/separate terminal to start Agents Playground
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

**Note:** The bot service start command keeps running and will not return to the prompt. This is expected — the server must stay running. Always start the service in a background terminal, then verify it started successfully by checking the output for messages like "listening on port" or "server started". If errors appear, read the logs, fix the issue, and restart. Use a **new terminal** for Agents Playground.

### CLI Options

| Option | Short | Required | Description |
|--------|-------|----------|-------------|
| `--app-endpoint` | `-e` | Recommended | Bot endpoint URL (e.g., http://localhost:3978/api/messages) |
| `--channel-id` | `-c` | Optional | Channel to emulate: msteams, emulator, webchat, directline |
| `--port` | `-p` | Optional | Server port (default: 56150, auto-fallback if occupied) |
| `--client-id` | `--cid` | Optional | Azure app client ID (for authenticated agents) |
| `--client-secret` | `--cs` | Optional | Azure app client secret (for authenticated agents) |
| `--tenant-id` | `--tid` | Optional | Azure tenant ID (for authenticated agents) |
| `--enable-events-recording` | `--er` | Optional | Enable events recording (default: false) |

### Examples

```bash
# Basic start with Teams channel
agentsplayground -e http://localhost:3978/api/messages -c msteams

# With authentication
agentsplayground -e http://localhost:3978/api/messages -c emulator \
  --client-id <CLIENT_ID> \
  --client-secret <CLIENT_SECRET> \
  --tenant-id <TENANT_ID>

# Test different channels
agentsplayground -e http://localhost:3978/api/messages -c webchat
agentsplayground -e http://localhost:3978/api/messages -c emulator
```

### Features

- **No Setup Required**: Works with HTTP localhost endpoints
- **Adaptive Card Preview**: See how cards render in Teams
- **Chat Interface**: Simulate user messages and bot responses
- **Context Mocking**: Mock Teams APIs (team members, channels, etc.)
- **Message Inspection**: View request/response payloads in real-time

### Limitations

- Application manifest not processed (command menus unavailable)
- Some Adaptive Card features unsupported (people picker, user mentions, stage view)
- SSO not supported
- Only Adaptive Cards supported (not Hero/Thumbnail cards)

---

## Option 2: Teams Direct Launch

**Use this when user explicitly asks to run on Teams.**

Test your agent in the actual Teams environment. Requires M365 account and HTTPS endpoint.

### Requirements

- Microsoft 365 account with sideloading enabled
- HTTPS endpoint (for bots, dev tunnels must be started first)

### Setup Steps

**For bot projects: Start devtunnel first (must be public/anonymous):**
```bash
devtunnel host -p 3978 --allow-anonymous
# Copy the tunnel URL and set BOT_ENDPOINT in env/.env.local before provisioning
```

### Quick Start

```bash
# Step 1: Provision and deploy
atk provision --env local -i false
atk deploy --env local -i false

# Step 2: Start your local service (this will HANG the terminal — expected!)
# Run as a background process (isBackground=true) since the server keeps running
# Check package.json scripts or project configuration for the appropriate start command
# - If project uses .localConfigs: use `npm run dev:teamsfx` or equivalent that loads .localConfigs
# - If project uses .env directly: use `npm run dev` or `npm start`
# Common patterns: npm run dev:teamsfx, npm run dev, npm start, python app.py, dotnet run

# Step 3: Open Teams with your app (use a NEW/separate terminal!)
# Get TEAMS_APP_ID and TENANT_ID from env/.env.local
# Open: https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}
```

### Opening in Different Hosts

Get your app IDs from `env/.env.local`, then open:

| Host | URL |
|------|-----|
| Teams web | `https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}` |
| Outlook web | `https://outlook.office.com/host/${{M365_APP_ID}}` |
| Office web | `https://www.office.com/m365apps/${{M365_APP_ID}}` |

### For Declarative Agents (No Backend Service)

```bash
# Just provision/deploy and open directly
atk provision --env local -i false
atk deploy --env local -i false
# Then open Teams and find your agent in the app list
```

### Dev Tunnels for Bots

**IMPORTANT**: For bot projects, you must start a public devtunnel BEFORE provisioning.

The tunnel must be public/anonymous so Teams can reach your bot:
```bash
# Start a public devtunnel
devtunnel host -p 3978 --allow-anonymous
```

Then set `BOT_ENDPOINT` in `env/.env.local` with the tunnel URL before running `atk provision`.

---

## Comparison

| Feature | Agents Playground | Teams Direct Launch |
|---------|-------------------|---------------------|
| Setup complexity | Simple | Requires provisioning |
| M365 account needed | No | Yes |
| HTTPS required | No | Yes (for bots) |
| Real Teams environment | No (simulated) | Yes |
| SSO testing | No | Yes |
| Speed | Fast | Slower (tunnel setup) |
| Recommended for | Testing first (recommended) | When user explicitly asks to run on Teams |

---

## Configuration File

Create `.m365agentsplayground.yml` in project root to mock Teams context:

```yaml
version: "0.1.1"
tenantId: 00000000-0000-0000-0000-0000000000001
bot:
  id: 00000000-0000-0000-0000-00000000000011
  name: Test Bot
currentUser:
  id: user-id-0
  name: Alex Wilber
  email: alexw@example.com
users:
  - id: user-id-1
    name: Megan Bowen
    email: meganb@example.com
personalChat:
  id: personal-chat-id
groupChat:
  id: group-chat-id
team:
  id: team-id
  name: My Team
  channels:
    - id: channel-announcements-id
      name: Announcements
```

---

## Environment Variables

For Agents Playground, you can use environment variables instead of CLI options:

| Variable | Description |
|----------|-------------|
| `BOT_ENDPOINT` | Bot endpoint URL |
| `DEFAULT_CHANNEL_ID` | Channel type (emulator, webchat, msteams) |
| `AUTH_CLIENT_ID` | Azure app client ID for authentication |
| `AUTH_CLIENT_SECRET` | Azure app client secret for authentication |
| `AUTH_TENANT_ID` | Azure tenant ID for authentication |

---

## Troubleshooting

### Playground won't start

Check if port is in use:
```powershell
# Windows
netstat -ano | findstr :56150

# The playground will automatically find an available port if 56150 is in use
```

### Bot not responding

1. Verify bot is running on specified endpoint
2. Check bot logs for errors
3. Ensure your bot endpoint is accessible:
   ```bash
   curl http://localhost:3978/api/messages
   ```

### Teams shows "app not available"

This usually means BOT_ENDPOINT requires HTTPS. Use Agents Playground instead, or ensure dev tunnel is running and BOT_ENDPOINT is properly configured.

### View help

```bash
agentsplayground --help
atk provision --help
atk deploy --help
```
