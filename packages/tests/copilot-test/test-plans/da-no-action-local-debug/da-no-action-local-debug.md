# Test Plan: DA No Action – Local Debug in M365 Copilot

## Metadata

- **feature-slug**: `da-no-action-local-debug`
- **owner**: v-helzha
- **created**: 2026-06-01
- **updated**: 2026-06-01
- **triggers**: issue-label `atk-copilot-test`, manual
- **source-plan-id**: `plan_2bd0acc6`

## Scope

**Covers:**
- Creating a new Declarative Agent project with **No Action** from scratch in VS Code
- Selecting "Default folder" as the workspace and naming the app
- Signing in to M365 via the ATK Accounts panel
- Launching local debug in M365 Copilot via **"Preview Local in Copilot (Chrome)"**
- Completing the M365 web sign-in in the launched Chrome browser
- Verifying the DA agent appears in M365 Copilot chat
- Sending a message and asserting the agent returns a response

**Does NOT cover:**
- Declarative Agent with actions (covered by `da-add-action-*` test plans)
- CLI-based project creation or debug
- Remote debug / provision to Azure
- Teams debug surface (covered by separate plans)

---

## Test Cases

### TC-001 – Create DA (No Action), sign in to M365, local debug, and verify agent responds in Copilot

**Preconditions:**
- VS Code is open with no project loaded (fresh state; ATK sign-in welcome dialog may be visible)
- ATK extension v6.8.0+ is installed and activated
- M365 account credentials are available (`M365_ACCOUNT_NAME` / `M365_ACCOUNT_PASSWORD`)
- No prior M365 account is signed in to ATK

**App name:** Generate at test start: `` const appName = `da-no-action-${Date.now()}` ``. Use as-is; no spaces. Copilot agent entry will be `appName + "local"`.

**Wizard flow:**

| Step | VS Code command / QuickPick / InputBox                    | Value to select / type                     | Wait-for-text before click |
|------|-----------------------------------------------------------|--------------------------------------------|----------------------------|
| 1    | `executeCommand("notifications.clearAll")`                | — (dismiss welcome dialog)                 | — |
| 2    | `executeCommand("fx-extension.create")` (no await)        | — (opens wizard)                           | — |
| 3    | New Project QuickPick — App type                          | `Declarative Agent`                        | wait `"Declarative Agent"` up to 60 s |
| 4    | DA — Add action choice                                    | `No Action`                                | wait `"No Action"` up to 20 s |
| 5    | Workspace Folder QuickPick                                | `Default folder`                           | wait `"Default folder"` up to 15 s |
| 6    | Application Name InputBox                                 | `${appName}` (type + Enter)                | wait `"Application Name"` up to 15 s |

**Dynamic data:**
- Generate the app name at test-start: `const appName = \`da-no-action-${Date.now()}\`` (alphanumeric, no spaces)
- The expected agent entry in Copilot is `appName + "local"` — e.g. `da-no-action-1717000000000local`

**Surface tags:** `[VSC]` = VS Code workbench/QuickPick (via `vscode.commands` or `sendSignal`). `[Chrome]` = external Chrome browser tab driven by Playwright CDP.

**Steps:**

**Phase 1 – Create project** `[VSC]`

1. `[VSC]` Dismiss the ATK sign-in welcome notification if visible: `executeCommand("notifications.clearAll")`; `takeScreenshot("01-atk-sidebar")`
2. `[VSC]` Fire the create-project wizard **without await**: `executeCommand("fx-extension.create")`; `wait(500)` to yield the event loop
3. `[VSC]` `sendSignal("waitForTextThenScreenshot:Declarative Agent:60000:02-da-option", 68000)`; `sendSignal("clickText:Declarative Agent", 10000)`; `wait(1000)`
4. `[VSC]` `sendSignal("waitForTextThenScreenshot:No Action:20000:03-no-action", 28000)`; `sendSignal("clickText:No Action", 10000)`; `wait(1000)`
5. `[VSC]` `sendSignal("waitForTextThenScreenshot:Default folder:15000:04-workspace-folder", 23000)`; `sendSignal("clickText:Default folder", 10000)`; `wait(1000)`
6. `[VSC]` `sendSignal("waitForTextThenScreenshot:Application Name:15000:05-app-name-input", 23000)`; `sendSignal("type:${appName}", 8000)`; `wait(500)`; `sendSignal("pressKey:Enter", 5000)`

**Phase 2 – Wait for scaffold and verify files** `[VSC]`

7. `[VSC]` Poll for `appPackage/manifest.json` inside the newly created project directory (ATK default: `~/AgentsToolkitProjects/${appName}`) for up to 90 s; `takeScreenshot("06-scaffold-complete")` once found; fail if still missing after 90 s
8. `[VSC]` Assert `appPackage/declarativeAgent.json` exists in the project directory; record pass/fail step
9. `[VSC]` Assert `appPackage/ai-plugin.json` does **NOT** exist (No Action DA has no plugin file); record pass/fail step
10. `[VSC]` Close the auto-opened README tab with `executeCommand("workbench.action.closeActiveEditor")`; `takeScreenshot("07-readme-closed")`; `wait(500)`
11. `[VSC]` Clear all notifications: `executeCommand("notifications.clearAll")`; `wait(500)`; `takeScreenshot("08-notifications-cleared")`

**Phase 3 – Sign in to M365** `[VSC → Chrome]`

12. `[VSC]` Open ATK Accounts QuickPick **without await**: `executeCommand("fx-extension.cmpAccounts")`; `wait(500)`
13. `[VSC]` `sendSignal("waitForTextThenScreenshot:Sign in to Microsoft 365:20000:09-signin-option", 28000)`; `sendSignal("clickText:Sign in to Microsoft 365", 10000)`; `wait(500)`
14. `[VSC]` `sendSignal("waitForTextThenScreenshot:Sign in:15000:10-modal-signin", 23000)`; `sendSignal("clickText:Sign in", 10000)`; `wait(1000)`
15. `[Chrome]` `sendSignal("waitForTextThenScreenshot:Email or phone:60000:11-m365-login-page", 68000)`; `sendSignal("clickText:Email or phone", 10000)`; `sendSignal("type:${process.env.M365_ACCOUNT_NAME}", 5000)`; `sendSignal("clickText:Next", 10000)`; `wait(3000)`
16. `[Chrome]` `sendSignal("type:${process.env.M365_ACCOUNT_PASSWORD}", 5000)`; `sendSignal("pressKey:Enter", 5000)`; `wait(3000)`; `takeScreenshot("12-m365-login-complete")`
17. `[Chrome]` `sendSignal("waitForTextThenScreenshot:Close:10000:13-close-browser-tab", 15000)`; `sendSignal("clickText:Close", 5000)`; `wait(3000)`; `takeScreenshot("14-m365-signed-in")`

**Phase 4 – Launch local debug in M365 Copilot** `[VSC → Chrome]`

18. `[VSC]` Fire debug-selection command **without await**: `executeCommand("workbench.action.debug.selectandstart")`; `wait(500)`
19. `[VSC]` `sendSignal("waitForTextThenScreenshot:Preview Local in Copilot:20000:15-debug-config", 28000)`; `sendSignal("clickText:Preview Local in Copilot (chrome)", 10000)`; `wait(1000)`; `takeScreenshot("16-debug-started")`
20. `[Chrome]` `sendSignal("waitForTextThenScreenshot:Email or phone:120000:17-copilot-login-page", 130000)`; `sendSignal("clickText:Email or phone", 10000)`; `sendSignal("type:${process.env.M365_ACCOUNT_NAME}", 5000)`; `sendSignal("clickText:Next", 10000)`; `wait(3000)`
21. `[Chrome]` `sendSignal("type:${process.env.M365_ACCOUNT_PASSWORD}", 5000)`; `sendSignal("clickText:Sign in", 10000)`; `wait(3000)`; `takeScreenshot("18-copilot-login-complete")`
22. `[Chrome]` `sendSignal("waitForTextThenScreenshot:Stay signed in:10000:19-stay-signed-in", 15000)`; `sendSignal("pressKey:Enter", 5000)`; `wait(3000)`

**Phase 5 – Validate agent in M365 Copilot** `[Chrome]`

23. `[Chrome]` `wait(30000)` for agent registration to propagate; `sendSignal("pressKey:Ctrl+-", 5000)` to zoom out; `takeScreenshot("20-copilot-zoomed")`
24. `[Chrome]` `sendSignal("pressKey:F5", 5000)` to reload the Copilot page; `wait(3000)`; `takeScreenshot("21-copilot-reloaded")`
25. `[Chrome]` `sendSignal("waitForTextThenScreenshot:${appName}local:30000:22-agent-visible", 38000)`; assert the agent entry `${appName}local` is visible — this is a **pass criterion**; fail the TC if timeout
26. `[Chrome]` `sendSignal("clickText:Message Copilot", 10000)`; `sendSignal("type:how can you assistant me?", 5000)`; `sendSignal("pressKey:Enter", 5000)`; `takeScreenshot("23-query-sent")`
27. `[Chrome]` `sendSignal("waitForTextThenScreenshot:response message:60000:24-agent-response", 68000)`; assert at least one bot response is visible — this is a **pass criterion**; `takeScreenshot("25-agent-response")`

**Expected result:**
- `appPackage/manifest.json` exists in the created project directory
- `appPackage/declarativeAgent.json` exists in the created project directory
- `appPackage/ai-plugin.json` does **NOT** exist (No Action DA has no plugin file)
- The agent `${appName}local` is listed in M365 Copilot after local debug launch
- The agent returns a non-empty response to `how can you assistant me?`

**Pass criteria:**
- `appPackage/manifest.json` and `appPackage/declarativeAgent.json` found on disk within 90 s after Enter on app name (step 7)
- `appPackage/ai-plugin.json` absent (step 9 assertion)
- `waitForTextThenScreenshot` resolves for `${appName}local` within 30 s in Copilot (step 25)
- `waitForTextThenScreenshot` resolves for a bot response within 60 s after sending query (step 27)

**Test script:**
`packages/tests/copilot-test/src/da-no-action-local-debug.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                                        | Pass condition                                               | Why                                                                            |
|-----|---------------------------------------|------------------------------------------------------------------------|--------------------------------------------------------------|--------------------------------------------------------------------------------|
| 01  | `01-atk-sidebar.png`                  | ATK sidebar after dismissing welcome notification                      | ATK panel visible                                            | Confirms clean starting state for project creation                             |
| 02  | `02-da-option.png`                    | New Project QuickPick with "Declarative Agent" option highlighted      | `"Declarative Agent"` text present                           | Confirms wizard opened and DA path is reachable                                |
| 03  | `03-no-action.png`                    | DA action-choice QuickPick with "No Action" option                     | `"No Action"` text present                                   | Confirms No Action variant is offered (vs. add-action variants)                |
| 04  | `04-workspace-folder.png`             | Workspace Folder QuickPick with "Default folder"                       | `"Default folder"` text present                              | Confirms project root selection step reached                                   |
| 05  | `05-app-name-input.png`               | Application Name InputBox                                              | `"Application Name"` text visible                            | Records app name step; name `${appName}` typed after this screenshot           |
| 06  | `06-scaffold-complete.png`            | VS Code explorer showing created project files                         | `appPackage/manifest.json` visible in file tree              | Confirms scaffold completed within 90 s                                        |
| 07  | `07-readme-closed.png`                | VS Code editor after closing README tab                                | README.md tab no longer present                              | Confirms cleanup before sign-in phase                                          |
| 08  | `08-notifications-cleared.png`        | VS Code with no notification banners                                   | No notification badge                                        | Confirms clean state before M365 sign-in                                       |
| 09  | `09-signin-option.png`                | ATK Accounts QuickPick with "Sign in to Microsoft 365" option          | `"Sign in to Microsoft 365"` text present                    | Confirms `fx-extension.cmpAccounts` command opened Accounts QuickPick          |
| 10  | `10-modal-signin.png`                 | ATK sign-in modal with "Sign in" button                                | `"Sign in"` button visible                                   | Confirms ATK sign-in modal appeared after selecting the sign-in option         |
| 11  | `11-m365-login-page.png`              | Microsoft web sign-in page in Chrome (Email/phone field)               | `"Email or phone"` text present                              | Confirms M365 OAuth flow launched from ATK `[Chrome]`                          |
| 12  | `12-m365-login-complete.png`          | M365 login page after credentials entered                              | Password field interaction captured                          | Records credential entry for audit `[Chrome]`                                  |
| 13  | `13-close-browser-tab.png`            | Chrome tab with "Close" option visible                                 | Tab can be closed                                            | Confirms sign-in completed and close is available `[Chrome]`                   |
| 14  | `14-m365-signed-in.png`               | VS Code (refocused) after sign-in browser tab closed                   | ATK shows no sign-in prompt                                  | Confirms M365 account is now active in ATK                                     |
| 15  | `15-debug-config.png`                 | Debug configuration dropdown with "Preview Local in Copilot" option    | `"Preview Local in Copilot"` text present                    | Confirms correct debug target selected (Copilot, not Teams)                    |
| 16  | `16-debug-started.png`                | VS Code debug session starting                                         | Debug toolbar visible or status bar shows debug indicator    | Confirms debug session launched                                                |
| 17  | `17-copilot-login-page.png`           | Microsoft sign-in page in Chrome (second login for Copilot) `[Chrome]` | `"Email or phone"` text present; up to 120 s wait            | Confirms Copilot browser sign-in step reached                                  |
| 18  | `18-copilot-login-complete.png`       | M365 Copilot loading in Chrome after credentials entered `[Chrome]`    | Page loading or Copilot UI visible                           | Confirms second sign-in completed                                              |
| 19  | `19-stay-signed-in.png`               | "Stay signed in?" prompt `[Chrome]`                                    | Prompt text visible                                          | Confirms post-auth prompt handled before proceeding                            |
| 20  | `20-copilot-zoomed.png`               | M365 Copilot at reduced zoom after 30 s wait `[Chrome]`                | Copilot chat UI visible without clipping                     | Confirms Copilot is stable after `wait(30000)` and zoom-out                    |
| 21  | `21-copilot-reloaded.png`             | M365 Copilot after F5 reload `[Chrome]`                                | Copilot chat UI reloaded                                     | Confirms page refreshed to pick up newly registered agent                      |
| 22  | `22-agent-visible.png`                | M365 Copilot agent list with `${appName}local` entry `[Chrome]`        | Agent name visible (30 s retry)                              | **Primary pass criterion**: DA registered and visible in Copilot               |
| 23  | `23-query-sent.png`                   | Copilot chat with query submitted `[Chrome]`                           | Query text visible in chat                                   | Records the query that triggers agent response                                 |
| 24  | `24-agent-response.png`               | Copilot chat with bot response `[Chrome]`                              | Non-empty response visible (60 s retry)                      | **Primary E2E assertion**: DA No Action agent responds in Copilot              |

---

## Notes

- This is a **local debug** test (no Azure provisioning). The debug surface is **M365 Copilot** (`Preview Local in Copilot (Chrome)`), not Teams.
- The test creates a brand-new project from scratch (VS Code must have no project open at start).
- **App name**: generate at test start as `` `da-no-action-${Date.now()}` ``; expected agent entry is `appName + "local"`. Do NOT hardcode the name.
- **VS Code command IDs** used:
  - `notifications.clearAll` — dismiss welcome dialog and clear notifications
  - `fx-extension.create` — open the New Project wizard (fire **without await**)
  - `fx-extension.cmpAccounts` — open ATK Accounts QuickPick (fire **without await**)
  - `workbench.action.closeActiveEditor` — close README tab (safe replacement for right-click → Close)
  - `workbench.action.debug.selectandstart` — open debug config picker (fire **without await**)
- **Right-click is not supported** by the signal mechanism; the README tab is closed with `executeCommand("workbench.action.closeActiveEditor")` instead.
- **Two separate M365 sign-in flows** occur: Phase 3 (ATK Accounts QuickPick → browser OAuth) and Phase 4 (Copilot browser sign-in). Both use `M365_ACCOUNT_NAME` / `M365_ACCOUNT_PASSWORD`.
- **Phase 3 browser sign-in timeout**: `waitForTextThenScreenshot` with 60 s (68 s total `sendSignal` timeout) for the `"Email or phone"` field.
- **Phase 4 browser sign-in timeout**: `waitForTextThenScreenshot` with **120 s** (130 s total) — ATK launches a new Chrome process for Copilot; startup can take up to 2 minutes.
- **Copilot agent propagation**: after Copilot launches, `wait(30000)` before zoom-out + F5 reload to let the agent registration propagate to Copilot's service endpoint.
- **Agent entry assertion (step 25)**: `waitForTextThenScreenshot` with 30 s timeout for `${appName}local` text in Copilot.
- **Bot response assertion (step 27)**: `waitForTextThenScreenshot` with 60 s timeout. The text to wait for should be a stable substring of any Copilot response, e.g. the last part of the query echoed back or any non-empty text in the response bubble area.
- Steps tagged `[VSC]` are driven from the VS Code extension host via `vscode.commands` or `sendSignal` targeting the VS Code window. Steps tagged `[Chrome]` require Playwright to be connected to the Chrome browser page via CDP; the same `sendSignal` mechanism is used but targets the externally launched browser.
- A "No Action" DA has no `ai-plugin.json` — the scaffold contains only `manifest.json` and `declarativeAgent.json`. This absence is an explicit pass criterion (step 9).
- The duplicate notification-clear pattern from the source JSON (`create_app_name` group + `clean_notification` group) is collapsed to a **single** `executeCommand("notifications.clearAll")` call (step 11) to avoid redundancy.
