# Declarative Agent Lifecycle Routing

Delegate every DA workflow to the `wiqd` plugin except TypeSpec DA projects. This includes pure DAs, DAs that attach existing OpenAPI, remote MCP, or Copilot Connector backends, and DAs that need a new Azure-hosted API backend. Do not run WIQD CLI commands from this skill or fall back to ATK when the delegated workflow fails.

Use ATK only when the request explicitly identifies the `declarative-agent-typespec` capability or the project contains a TypeSpec DA marker.

## DA Routing

Before selecting a lifecycle owner:

1. Determine DA intent using the [detection criteria](#detect-a-declarative-agent). An explicit DA creation request does not require an existing project or manifest.
2. Before invoking any skill for an existing project, inspect it and select ATK when the request explicitly identifies the `declarative-agent-typespec` capability or the project contains any of these TypeSpec DA markers:
   - A `typeSpec/compile` action in `m365agents.yml` or `m365agents.local.yml`.
   - A `tspconfig.yaml` entry for `@microsoft/typespec-m365-copilot`.
   - A `.tsp` source that imports `@microsoft/typespec-m365-copilot`.
3. Otherwise, invoke the WIQD plugin.

Backend source, build configuration, infrastructure, or lifecycle deployment actions do not create another ATK exception. For a new Azure-hosted OpenAPI backend, let the WIQD plugin coordinate the official Azure skill and, when required, the Entra skill.

## Detect a Declarative Agent

Treat the request or project as a DA when any of these conditions is true:

- The user explicitly says "Declarative Agent" or "DA" in a Declarative Agent context.
- `appPackage/declarativeAgent.json` exists.
- `appPackage/manifest.json` contains `copilotAgents.declarativeAgents`.

The presence of `m365agents.yml` does not make a project non-DA. Check the DA markers before selecting a lifecycle owner.

## Read-Only Reference Requests

Answer questions about DA schemas, manifest fields, capabilities, examples, or project structure from the local references without invoking the WIQD plugin. Delegate only when the user requests a DA lifecycle operation owned by WIQD.

## WIQD Plugin Delegation

After DA routing selects WIQD:

1. Invoke the `wiqd` skill.
2. Preserve the user's complete intent and relevant context, including the project path, requested lifecycle operations, selected OpenAPI operations, MCP endpoint, and authentication requirements already supplied by the user.
3. Let the WIQD plugin perform its own discovery and choose the appropriate workflow.
4. Let the plugin own prerequisites, authentication, EULA handling, command selection, file-based fallback, validation, and error recovery.
5. Keep the workflow delegated until it completes or the plugin asks the user for required input.

Do not:

- Run `wiqd` CLI commands directly.
- Copy WIQD command syntax, setup instructions, authentication steps, or EULA text into this skill.
- Preempt the plugin's workflow by editing DA project files before delegation.
- Invoke ATK as a fallback when the plugin reports a failure.
- Invoke the WIQD plugin for a read-only reference question that this skill can answer locally.

## Lifecycle Ownership

Apply one owner for the complete lifecycle:

| Scenario                                                      | Lifecycle owner |
| ------------------------------------------------------------- | --------------- |
| Pure DA                                                       | WIQD plugin     |
| DA that references an existing Copilot Connector connection   | WIQD plugin     |
| DA that attaches an existing OpenAPI API or remote MCP server | WIQD plugin     |
| DA that needs a new Azure-hosted API backend                  | WIQD plugin, coordinating Azure and Entra skills as required |
| DA with existing backend source or deployment actions         | WIQD plugin     |
| TypeSpec DA                                                   | ATK             |

Do not switch owners after a lifecycle failure. Apply this routing rule to every existing project.
