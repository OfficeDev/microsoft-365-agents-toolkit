# Microsoft 365 Agents Toolkit Skill

[![npm version](https://badge.fury.io/js/%40microsoft%2Fm365agentstoolkit-skill.svg)](https://www.npmjs.com/package/@microsoft/m365agentstoolkit-skill)

Install the **Microsoft 365 Agents Toolkit** skill into your project so that AI coding tools can help you build Teams apps, Declarative Agents, Custom Engine Agents, and Microsoft 365 integrations.

Supports:

| AI Tool | Config file(s) created |
|---|---|
| [Claude Code](https://claude.ai/code) | `CLAUDE.md` + `.m365-atk-skill/` |
| [Cursor](https://cursor.com) | `.cursor/rules/m365-atk.mdc` + `.cursor/m365-atk-skill/` |
| [GitHub Copilot](https://github.com/features/copilot) CLI / Coding Agent | `.github/copilot-instructions.md` + `.github/skills/microsoft-365-agents-toolkit/` |
| [OpenAI Codex](https://platform.openai.com/docs/codex) | `AGENTS.md` + `.m365-atk-skill/` |

## Quick Start

```bash
# Interactive — choose your AI tool
npx @microsoft/m365agentstoolkit-skill

# Non-interactive
npx @microsoft/m365agentstoolkit-skill --tool claude-code
npx @microsoft/m365agentstoolkit-skill --tool cursor
npx @microsoft/m365agentstoolkit-skill --tool copilot
npx @microsoft/m365agentstoolkit-skill --tool codex
npx @microsoft/m365agentstoolkit-skill --tool all
```

Run this command in the **root of the project** you're building with ATK. The installer copies the skill markdown files to the appropriate location and creates (or updates) the entry-point config file for your AI tool.

## What Gets Installed

### Claude Code

```
CLAUDE.md                ← references @./.m365-atk-skill/SKILL.md
.m365-atk-skill/
  SKILL.md               ← main skill (auto-loaded via @import)
  create-project/
  test-playground/
  test-teams/
  provision-deploy/
  troubleshoot/
  slack-to-teams/
  toolkit/
  experts/               ← 100+ micro-expert files
  docs/
```

Claude Code auto-loads `CLAUDE.md` on every session. The `@./.m365-atk-skill/SKILL.md` import pulls in the full skill.

### Cursor

```
.cursor/rules/m365-atk.mdc          ← Cursor rule (description + globs)
.cursor/m365-atk-skill/             ← full skill files for reference
```

Cursor loads `.cursor/rules/*.mdc` automatically. The rule is scoped to `m365agents*.yml`, `teamsapp*.yml`, and TypeScript/JavaScript/Python files.

### GitHub Copilot

```
.github/copilot-instructions.md     ← references the skill location
.github/skills/microsoft-365-agents-toolkit/
  SKILL.md
  ...
```

Copilot Coding Agent picks up `.github/copilot-instructions.md` and `.github/skills/` automatically.

### OpenAI Codex

```
AGENTS.md                           ← contains full SKILL.md content
.m365-atk-skill/                    ← full skill files for reference
```

## Re-running / Updating

Re-run the installer at any time to update the skill to the latest version.
Existing blocks (identified by HTML marker comments) are replaced rather than duplicated.

```bash
npx @microsoft/m365agentstoolkit-skill --tool claude-code
```

## What the Skill Does

Once installed, ask your AI tool to:

- `"Create a new Teams bot"` → scaffolds an ATK project
- `"Test my agent locally with Agents Playground"` → sets up local dev
- `"Provision Azure resources and deploy"` → runs `atk provision` + `atk deploy`
- `"Fix my 401 error"` → runs the troubleshoot workflow
- `"Migrate my Slack bot to Teams"` → runs the slack-to-teams workflow

## Requirements

- Node.js ≥ 18
- ATK CLI: `npm i -g @microsoft/m365agentstoolkit-cli@beta` (version > 1.1.5-beta)

## Source

This skill ships with the [Microsoft 365 Agents Toolkit VS Code extension](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension).
The canonical source is in the [`OfficeDev/microsoft-365-agents-toolkit`](https://github.com/OfficeDev/microsoft-365-agents-toolkit) repository under `packages/vscode-extension/skills/microsoft-365-agents-toolkit/`.

## License

MIT — see [LICENSE](LICENSE)
