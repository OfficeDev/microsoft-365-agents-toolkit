# Test with Agents Playground

Test your bot locally — interactively or with automated smoke tests — using the Microsoft 365 Agents Playground toolset. No M365 account, Azure tunnel, or app registration required.

**Recommend Agents Playground first for testing**, unless user explicitly asks to run on Teams.

## Intent Router

| User Intent | Read |
|---|---|
| Manual/interactive testing, chat with the bot, explore responses in a UI | → [playground.md](playground.md) |
| Automated tests, CI pipeline, smoke tests, programmatic testing, `TestClient` | → [playground-cli.md](playground-cli.md) |
| "Test my bot" / "check if it works" / no specific preference | → **Combined Workflow** below |

---

## Combined Workflow (Default Recommendation)

> Use this when the user says "test my bot", "run some tests", or does not specify manual vs automated.

Run automated smoke tests first to get instant pass/fail results, then open Agents Playground for interactive exploration.

### Step 1: Start the bot service

Start the bot as a background process. **This will hang the terminal — expected.**

```bash
# ATK project
npm run dev:teamsfx:playground

# Custom project
npm run dev
```

Verify startup by checking output for `"listening on port"` or `"server started"`.

### Step 2: Inspect the bot and generate smoke tests

Read [playground-cli.md — Smoke Test Workflow](playground-cli.md) for full instructions. In summary:

1. **Read the bot's entry point** (`app.ts`, `index.ts`, or equivalent):
   - Find message handlers (`app.message`, `app.command`, etc.)
   - Find AI system prompt (if using `ChatPrompt`) to understand the bot's purpose
   - Identify known commands or keywords

2. **Design 5–7 test cases** covering:
   - Greeting (`"Hello"`)
   - Main feature (1–2 inputs that exercise the bot's primary purpose)
   - Known command (e.g., `"/help"`)
   - Follow-up in same conversation (tests memory/context)
   - Unknown/edge input (e.g., `"xyzzy"`)

3. **Write and run `_smoke-test.mjs`** in the project root (see template in [playground-cli.md](playground-cli.md)).

### Step 3: Present results as a table

After running the smoke tests, present results to the user in this format:

| # | Test | Input | Status | Response (first 100 chars) |
|---|------|-------|--------|---------------------------|
| T1 | Greeting | `Hello` | ✅ PASS | Hello! I'm your assistant... |
| T2 | Main feature | `What can you do?` | ✅ PASS | I can help you with tasks... |
| T3 | Help command | `/help` | ✅ PASS | [Card] |
| T4 | Follow-up | `Tell me more` | ✅ PASS | Sure! Here are the details... |
| T5 | Unknown input | `xyzzy` | ✅ PASS | I didn't understand that. |

If any test fails (`❌`), diagnose the bot output before proceeding.

### Step 4: Open Agents Playground for interactive exploration

In a **new terminal** (bot service must stay running):

```bash
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

The user can now chat with the bot, inspect Adaptive Card rendering, and test scenarios beyond the smoke tests. See [playground.md](playground.md) for full CLI options and configuration.

---

## Installation

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

## Quick Start

```bash
# 1. For ATK projects, deploy playground config first
atk deploy --env playground -i false

# 2. Start your bot service (this will HANG the terminal — expected!)
# Run as a background process since the server keeps running
cd my-bot
npm run dev:teamsfx:playground  # For ATK projects
# npm run dev                   # For customized projects

# 3. Use a NEW/separate terminal to start Agents Playground
agentsplayground -e http://localhost:3978/api/messages -c msteams
```

**Note:** The bot service start command keeps running and will not return to the prompt. This is expected — the server must stay running. Always start the service in a background terminal, then verify it started successfully by checking the output for messages like "listening on port" or "server started". If errors appear, read the logs, fix the issue, and restart. Use a **new terminal** for Agents Playground.

## CLI Options

| Option | Short | Required | Description |
|--------|-------|----------|-------------|
| `--app-endpoint` | `-e` | Recommended | Bot endpoint URL (e.g., http://localhost:3978/api/messages) |
| `--channel-id` | `-c` | Optional | Channel to emulate: msteams, emulator, webchat, directline |
| `--port` | `-p` | Optional | Server port (default: 56150, auto-fallback if occupied) |
| `--client-id` | `--cid` | Optional | Azure app client ID (for authenticated agents) |
| `--client-secret` | `--cs` | Optional | Azure app client secret (for authenticated agents) |
| `--tenant-id` | `--tid` | Optional | Azure tenant ID (for authenticated agents) |
| `--enable-events-recording` | `--er` | Optional | Enable events recording (default: false) |

## Examples

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

## Features

- **No Setup Required**: Works with HTTP localhost endpoints
- **Adaptive Card Preview**: See how cards render in Teams
- **Chat Interface**: Simulate user messages and bot responses
- **Context Mocking**: Mock Teams APIs (team members, channels, etc.)
- **Message Inspection**: View request/response payloads in real-time

## Limitations

- Application manifest not processed (command menus unavailable)
- Some Adaptive Card features unsupported (people picker, user mentions, stage view)
- SSO not supported
- Only Adaptive Cards supported (not Hero/Thumbnail cards)

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

## Environment Variables

You can use environment variables instead of CLI options:

| Variable | Description |
|----------|-------------|
| `BOT_ENDPOINT` | Bot endpoint URL |
| `DEFAULT_CHANNEL_ID` | Channel type (emulator, webchat, msteams) |
| `AUTH_CLIENT_ID` | Azure app client ID for authentication |
| `AUTH_CLIENT_SECRET` | Azure app client secret for authentication |
| `AUTH_TENANT_ID` | Azure tenant ID for authentication |

## References

- For project file details → [../toolkit/manifest-and-yaml.md](../toolkit/manifest-and-yaml.md)
- If something goes wrong → [../troubleshoot/troubleshoot.md](../troubleshoot/troubleshoot.md)
- To test on real Teams instead → [../test-teams/test-teams.md](../test-teams/test-teams.md)
- Manual testing sub-skill → [playground.md](playground.md)
- Automated testing sub-skill → [playground-cli.md](playground-cli.md)

> **Applies to: code-based Teams bots/agents only.** Declarative agents and API plugins do not run in Agents Playground — they must be tested in M365 Copilot via [test-teams](../test-teams/test-teams.md).
