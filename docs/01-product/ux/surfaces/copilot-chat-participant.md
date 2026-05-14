# Copilot Chat participants

## What this extension registers

One participant, contributed by `packages/vscode-extension/package.json` `chatParticipants`:

| Participant ID | Name | Subcommands |
|----------------|------|-------------|
| `ms-teams-vscode-extension.office` | `office` | `create`, `generatecode`, `nextstep` |

The `office` participant assists with Office Add-in scaffolding and code generation. Source: `packages/vscode-extension/src/chat/`.

## What this extension does NOT register: `@m365agents`

`@m365agents` is a chat participant the user invokes for Microsoft 365 Agents Toolkit guidance. **It is registered by a separate companion extension**, not by `packages/vscode-extension`. This extension contains command handlers (in [`packages/vscode-extension/src/handlers/copilotChatHandlers.ts`](../../../packages/vscode-extension/src/handlers/copilotChatHandlers.ts)) that **construct queries and open Copilot Chat with `@m365agents` pre-filled**, but the participant binding lives elsewhere.

When the companion extension is not installed, the link buttons (e.g. "Get Help from GitHub Copilot") direct the user to install it via `aka.ms/install-m365agents`. The walkthrough `teamsAgentGetStarted` includes steps to install the M365 Agents companion plus its dependencies (GitHub Copilot, GitHub Copilot Chat).

From the user's perspective `@m365agents` is the M365 Agents Toolkit chat experience. From this repo's perspective the companion extension is an **external dependency** that the toolkit links to but does not own.

### Pre-filled queries this extension constructs

| Trigger | Pre-filled query (verbatim from `copilotChatHandlers.ts`) |
|---------|-----------------------------------------------------------|
| `fx-extension.invokeChat` (default intro) | `@m365agents Use this GitHub Copilot extension to ask questions about the development of apps and agents you build for Copilot and Microsoft 365 apps.` |
| Walkthrough sample 1 | `@m365agents What's the difference between declarative and custom agents?` |
| Walkthrough sample 2 | `@m365agents I want to create a ToDo app.` |
| Troubleshoot prompt | `@m365agents My app doesn't sideload when debugging with Microsoft 365 Agents Toolkit.` |
| `fx-extension.teamsAgentTroubleshootError` | `@m365agents` followed by the surfaced `FxError` |
| `fx-extension.teamsAgentTroubleshootSelectedText` | `@m365agents Resolve: <selected text>` |
| Capability discovery prompt | `@m365agents Write your own query message to find relevant templates or samples to build your app and agent...` |

## Why a chat participant rather than a generic Copilot prompt

- The companion `@m365agents` participant has the **toolkit's tools registered** (template registry, driver list, error catalogue) — generic Copilot would have to guess.
- It can call **engine operations directly** rather than emit code that the user copies and runs.
- Telemetry from this extension's commands is attributed: `triggerFrom = TelemetryTriggerFrom.CopilotChat`.

## Telemetry attribution

Every command invoked via the extension's `invokeChat`/`teamsAgentTroubleshoot*` paths passes `triggerFrom = TelemetryTriggerFrom.CopilotChat` so we can measure chat as an entry point separately from the tree view and command palette. The companion extension owns its own telemetry independently.

## v4 design implication

v4 may either:

1. Continue the current split — v4 owns the in-VS-Code link/handler surface, the companion extension owns `@m365agents`.
2. Move the participant in-house — v4 registers `@m365agents` directly in `chatParticipants`.

This is an open design decision. The current split was driven by extension-marketplace constraints (separate company-of-record for the participant) rather than architectural choice.
