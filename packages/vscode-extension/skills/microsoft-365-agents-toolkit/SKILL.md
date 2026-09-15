---
name: microsoft-365-agents-toolkit
description: "Builds, tests, and deploys Microsoft 365 apps and agents for Teams and Copilot. Uses WIQD for pure Declarative Agents and DAs backed by existing APIs, MCP servers, or Copilot Connector connections; uses ATK for hybrid Declarative Agents with project-owned backend code and for ATK-only specialized DA capabilities. Uses ATK for Custom Engine Agents, Teams bots, tabs, message extensions, Agents Playground, and Teams runtime testing. USE FOR: Declarative Agents, Custom Engine Agents, Teams apps, local testing, deployment, troubleshooting, and Slack-to-Teams migration. DO NOT USE FOR: general web development or non-Microsoft-365 projects."
---

# Microsoft 365 Agents Toolkit Skill

Build pure Declarative Agents and DAs that use existing APIs, MCP servers, or Copilot Connector connections with WIQD. Use ATK for hybrid DAs that own backend compute, including projects that build and deploy a new Copilot Connector backend, ATK-only specialized DA capabilities such as MetaOS and TypeSpec, non-DA Teams apps, code-based agents, and runtime testing.

## Declarative Agent Routing

Treat a request or project as a DA when the user explicitly says Declarative Agent, `appPackage/declarativeAgent.json` exists, or `appPackage/manifest.json` contains `copilotAgents.declarativeAgents`. The presence of `m365agents.yml` does not override these markers.

Determine DA routing from project structure, not from which CLI created it. WIQD delegates scaffolding and lifecycle work to ATK/fx-core, so creator identity is not a reliable project-type signal.

Treat a DA as hybrid only when it has project-owned backend source code and its lifecycle configuration provisions or deploys that backend. A DA action that calls an existing OpenAPI service or remote MCP server is not hybrid. The presence of `m365agents.yml`, a DA action, or an API plugin alone is insufficient.

Use ATK for an explicitly identified ATK-only specialized DA capability or for a hybrid DA. Use WIQD for pure DAs and DAs that use existing OpenAPI, remote MCP, or Copilot Connector backends. Referencing an existing Copilot Connector connection is not a hybrid signal; owning and deploying the connector implementation is. If the structure is incomplete or unusual, inspect the backend source and lifecycle deployment actions; ask whether the backend is owned and deployed by this project only when that cannot be determined. Do not switch tools as a runtime fallback. If WIQD is required but unavailable, provide its installation guidance and stop.

Answer read-only DA schema, manifest, capability, example, and project-structure questions from local references without requiring WIQD installation or login.

## AI Behavior Guidelines

1. **Testing Strategy:** Test Declarative Agents in Microsoft 365 Copilot. For code-based agents and Teams apps, recommend Agents Playground first (faster, no M365 needed) and use the Teams workflow only if the user explicitly requests it.

2. **Environment Variables:** NEVER hardcode secrets or make up placeholder values. Always ask users for real values.

3. **Error Handling:** Read error messages carefully. Check `env/.env.local`, `.localConfigs`, and `atk auth list`. Common pitfalls:
   - **`AADSTS7000229`** → `aadApp/create` missing `generateServicePrincipal: true` in YAML — add it and re-provision
   - **Missing `TENANT_ID`** in `.localConfigs` → SDK uses wrong token authority → 401 from Bot Connector
   - **401 persists after auth fix** → devtunnel URL may be blacklisted — create a fresh tunnel
   - See [troubleshoot/troubleshoot.md](troubleshoot/troubleshoot.md) for full diagnostic steps

4. **Long-Running Commands — WAIT for completion:**
   - `atk new`, `atk provision`, `atk deploy` can take several minutes
   - Always wait for completion before running the next step (timeout 120000ms+)

5. **Local Service Startup — Hangs terminal (expected):**
   - `npm run dev`, `npm start`, `python app.py`, `devtunnel host`, etc. will hang — the process keeps running indefinitely
   - ALWAYS run as a background process (`isBackground=true`) — NEVER use `isBackground=false` for these commands
   - Do NOT wait for it to "finish" — verify startup by checking output for "listening on port" or tunnel URL
   - If errors appear, read logs, diagnose, fix, restart
   - Use a **NEW terminal** to launch Agents Playground or open Teams sideloading URL

6. **Monitor App Logs:** Periodically check background terminal output for runtime errors. If the app crashes, read the error, fix the root cause, and restart.

7. **Telemetry Tagging:** Before running any `atk` CLI commands, set the session environment variable so all CLI invocations are tagged as skill-initiated:
   ```bash
   export ATK_CLI_SKILL=true
   ```
   Run this once at the start of the session. All subsequent `atk` commands in the same terminal will inherit it.

## ATK CLI Setup

```bash
atk --version  # Must be > 1.1.5-beta
```

If ATK is not found or version is too old:
```bash
npm i -g @microsoft/m365agentstoolkit-cli@beta
```

## CLI Global Options

| Option | Meaning | Recommendation |
| --- | --- | --- |
| `-i` | Interactive mode | Always use `-i false` in automation to avoid hanging |
| `-f` | Project folder | Default to be current directory, used when specifying a custom folder. When scaffolding a new project, this is the parent folder where the project folder will be created under. |
| `-h` | Command help | Use `atk <command> -h` for quick syntax checks |

## Sub-Skills

| Sub-Skill | When to Use | Reference |
|-----------|-------------|-----------|
| **create-project** | Scaffold new project from template, choose template, `atk new` | [create-project/create-project.md](create-project/create-project.md) |
| **test-playground** | Test locally with Agents Playground, `agentsplayground`, quick testing | [test-playground/test-playground.md](test-playground/test-playground.md) |
| **test-teams** | Run on Teams, devtunnel, sideload, Teams testing, test in Copilot | [test-teams/test-teams.md](test-teams/test-teams.md) |
| **provision-deploy** | Provision Azure resources, deploy to cloud, `atk provision`, `atk deploy` | [provision-deploy/provision-deploy.md](provision-deploy/provision-deploy.md) |
| **troubleshoot** | Fix errors, 401, port conflicts, YAML errors, stale bots | [troubleshoot/troubleshoot.md](troubleshoot/troubleshoot.md) |
| **slack-to-teams** | Migrate Slack bot to Teams, cross-platform bridging, Block Kit to Adaptive Cards | [slack-to-teams/SKILL.md](slack-to-teams/SKILL.md) |

> **MANDATORY:** Before executing any workflow, read the corresponding sub-skill document.

## Shared References

- [manifest-and-yaml.md](toolkit/manifest-and-yaml.md) — Project files, YAML config, env vars, .localConfigs flow
- [commands.md](toolkit/commands.md) — ATK CLI commands: package, validate, share, collaborate
- [templates.md](toolkit/templates.md) — Complete template catalog with language support
- [experts/](experts/index.md) — 100+ micro-expert files: Teams SDK, Slack SDK, cross-platform bridging, deploy, AI models, security, language conversion
- [docs/](docs/README.md) — Platform comparison guides: UI, messaging, identity, infrastructure, feature gaps

## Workflow Chains

Match user intent to the smallest valid workflow.

| User Intent | Workflow (read in order) |
|---|---|
| Build a pure DA or attach an existing API/MCP/Copilot Connector backend | create-project (WIQD) → test-teams (M365 Copilot) |
| Build a hybrid DA with project-owned backend source and deployment, including a new Connector backend | create-project (ATK) → provision-deploy → test-teams (M365 Copilot) |
| Build an ATK-only specialized DA such as MetaOS or TypeSpec | create-project (ATK) → follow generated ATK lifecycle |
| Build a code-based agent or Teams app from scratch | create-project → test-playground |
| Test existing project locally | test-playground (recommended) or test-teams |
| Deploy to Azure | provision-deploy |
| Fix broken bot | troubleshoot → re-test |
| Migrate Slack bot to Teams | slack-to-teams |

> **MANDATORY:** Before executing any slack-to-teams workflow, read [slack-to-teams/SKILL.md](slack-to-teams/SKILL.md) first. The sub-skill contains a routed expert system with 100+ micro-expert files for cross-platform bot development.

## ATK Project Context Resolution

Resolve config values only when missing. If a value is already known in the session, reuse it.

### Step 1: Detect ATK Project

Check the Declarative Agent markers above first. If none match and `m365agents*.yml` exists in the current folder, treat it as an ATK project and parse configuration.

### Step 2: Resolve Common Configuration

Resolve variables referenced in `m365agents*.yml`. Common variables:
AZURE_OPENAI_API_KEY
AZURE_OPENAI_ENDPOINT
AZURE_OPENAI_DEPLOYMENT_NAME

### Step 3: Collect Missing Values

If required values are missing, ask the user for only the missing ones.

Refer to [manifest-and-yaml.md](toolkit/manifest-and-yaml.md) for full config-file details.