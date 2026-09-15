<!--
SOURCE OF TRUTH: microsoft/microsoft-365-agents-toolkit
PATH: packages/vscode-extension/skills/microsoft-365-agents-toolkit/toolkit/
This folder is the canonical reference for the M365 Agents Toolkit toolchain.
Other repos may sync this folder via their knowledge-sync workflow — do NOT
edit copies downstream; open a PR upstream instead.
-->

# M365 Agents Toolkit — Toolchain Knowledge

Reference material for the **toolchain** itself: the `atk` CLI, the `m365agents.yml` lifecycle, environment files, project templates, manifests, the Agents Playground, and publishing. Use [declarative-agent-lifecycle.md](declarative-agent-lifecycle.md) as the single source of truth for DA routing.

ATK lifecycle content applies to non-DA projects and DAs selected by that routing gate. Read-only manifest reference content can also answer DA schema questions without requiring WIQD setup.

For SDK code patterns (handlers, AI prompts, Adaptive Cards, MCP, OAuth, etc.), see the sibling [../experts/](../experts/) folder. For Slack-vs-Teams platform comparison, see [../docs/](../docs/).

## Files

| File | Scope |
|---|---|
| [declarative-agent-lifecycle.md](declarative-agent-lifecycle.md) | Structural routing for all DAs and WIQD lifecycle commands for DAs without project-owned backend deployment |
| [templates.md](templates.md) | Full `atk new -c` capability catalog: declarative agents (8 variants), Copilot connectors, Office add-ins, Teams bots/tabs/message extensions, custom engine agents, RAG agents |
| [commands.md](commands.md) | `atk` CLI reference outside the lifecycle: `add action`, `add auth-config`, `regenerate action`, `share`, `collaborator`, `env`, `install/uninstall`, `upgrade`, `doctor` |
| [lifecycle-cli.md](lifecycle-cli.md) | Lifecycle CLI commands (`provision`, `deploy`, `package`, `validate`, `preview`) and the full `m365agents.yml` action catalog |
| [manifest-and-yaml.md](manifest-and-yaml.md) | `appPackage/manifest.json` + `appPackage/declarativeAgent.json`, YAML action field reference, common-mistake table, `signInAudience` configuration |
| [environments.md](environments.md) | `env/.env.{name}` + `.user` files, `${{VAR}}` resolution, `SECRET_` prefix, `.localConfigs` flow, multi-environment isolation |
| [playground.md](playground.md) | Agents Playground (`agentsplayground` CLI, `.m365agentsplayground.yml`, channel emulation, activity simulation) |
| [publish.md](publish.md) | Publishing workflow: sideload → org catalog (`atk publish`) → Teams Store / Partner Center; version bumping, validation requirements |

## Capability matrix

| Capability | Templates | What applies from this folder |
|---|---|---|
| **Declarative agents** | `declarative-agent` for legacy recognition; `declarative-agent-action*` and `declarative-agent-with-graph-connector` for project-owned backend scaffolds | `declarative-agent-lifecycle.md` for structural routing; `manifest-and-yaml.md` for read-only manifest questions. DAs run in M365 Copilot, not Playground. |
| **API plugins** | `declarative-agent-action*`, `declarative-agent-action-from-existing-api`, `add action` | Use ATK for a project-owned backend scaffold and the project's full lifecycle. Use WIQD when the DA attaches an existing OpenAPI API or remote MCP server. |
| **Standalone Copilot connector backend** | `copilot-connector` | ATK `templates.md`, `commands.md`, `lifecycle-cli.md`, and `environments.md`. A DA that only references an existing Connector connection follows `declarative-agent-lifecycle.md`. |
| **Custom engine agents** | `basic-custom-engine-agent`, `weather-agent`, `foundry-agent-to-m365`, `coffee-agent`, `data-analyst-agent-v2` | All. Compute deploy via `lifecycle-cli.md` (`arm/deploy` + `azureAppService/zipDeploy`). |
| **Teams bots / tabs / message extensions** | `bot`, `tab`, `message-extension`, `teams-agent*`, `teams-collaborator-agent`, `bot-sso` | All. Pair with [../experts/teams/](../experts/teams/) for SDK code patterns. |
| **Office add-ins** | `office-addin-outlook-taskpane`, `office-addin-wxpo-taskpane`, `office-addin-excel-cfshortcut`, `office-addin-config` | `templates.md`, `commands.md`, `lifecycle-cli.md`. Add-in-specific runtime is out of scope here. |

## Cross-references

- Workflow how-tos that consume this knowledge live one level up: [../create-project/](../create-project/), [../test-playground/](../test-playground/), [../test-teams/](../test-teams/), [../provision-deploy/](../provision-deploy/), [../troubleshoot/](../troubleshoot/).
- For Teams-bot SDK code (DevtoolsPlugin, ConsoleLogger, runtime handlers, Adaptive Cards): see [../experts/teams/](../experts/teams/).
- For deploying without ATK (manual `az` CLI walkthrough): see [../experts/deploy/azure-bot-deploy-ts.md](../experts/deploy/azure-bot-deploy-ts.md).
