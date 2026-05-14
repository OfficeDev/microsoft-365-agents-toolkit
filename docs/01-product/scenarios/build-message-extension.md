# Scenario — Build a Message Extension

**Persona:** P1 / P3
**Outcome:** A search-based message extension surfaced in Teams compose.
**Surface:** VS Code · CLI

## 1 — Scaffold

| Template ID | Languages |
|-------------|-----------|
| `message-extension-v2` | TS · Python |

## Lifecycle

Mostly identical to a [bot](build-bot.md): bot AAD app, bot channel, App Service / Functions, Teams app, deploy, publish.

The differentiator is the **manifest composeExtensions** block (registered as part of `teamsApp/configure`).

## Where this is tested

- ADO suite: 34869329
