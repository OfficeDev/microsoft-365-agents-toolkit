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

**Wizard flow:**

| Step | QuickPick / InputBox                                      | Value to select / type |
|------|-----------------------------------------------------------|------------------------|
| 1    | ATK sidebar — "Create a New Agent/App"                    | click button |
| 2    | New Project — App type                                    | `Declarative Agent` |
| 3    | DA — Add action choice                                    | `No Action` |
| 4    | Workspace Folder                                          | `Default folder` |
| 5    | Application Name InputBox                                 | `vscuse_app_#####` (auto-generated; type + Enter) |

**Steps:**

**Phase 1 – Create project**

1. If the ATK sign-in welcome dialog is visible, close it (red close button); take screenshot confirming ATK sidebar is visible
2. Click **"Create a New Agent/App"** in the ATK sidebar; take screenshot of the New Project QuickPick
3. Select **"Declarative Agent"**; take screenshot
4. Select **"No Action"** in the "Add an action to your declarative agent" QuickPick; take screenshot
5. Click **"Default folder"** in the Workspace Folder dialog; take screenshot
6. Type the application name and press Enter; take screenshot

**Phase 2 – Post-creation cleanup**

7. Wait up to 60 s for the **"Preview README.md"** tab to open automatically; take screenshot
8. Right-click the **"Preview README.md"** tab header and click **"Close"** to dismiss it; take screenshot
9. Open Command Palette (Ctrl+Shift+P), type `notifications: clear`, press Enter to clear all notifications; take screenshot
10. Open Command Palette (F1), type `Notifications: Clear All Notifications`, press Enter; take screenshot confirming notification count is zero

**Phase 3 – Sign in to M365**

11. Open Command Palette (F1), type `View: Show Microsoft 365 Agents Toolkit`, press Enter; take screenshot of the ATK panel
12. Open Command Palette (F1), type `Microsoft 365 Agents: Accounts`, press Enter; click the **"Microsoft 365 Agents: Account"** option; take screenshot
13. Click **"Sign in to Microsoft 365"** in the ATK dropdown; click the **"Sign in"** button in the modal dialog; take screenshot
14. In the Chrome browser sign-in page: click the Email/phone field (wait up to 60 s for the page to appear); take screenshot
15. Type `M365_ACCOUNT_NAME` → click **"Next"**; type `M365_ACCOUNT_PASSWORD` → press **Enter**; take screenshot
16. Close the sign-in browser tab (`delay: 3 s`); take screenshot confirming M365 account name visible in ATK accounts section

**Phase 4 – Launch local debug in M365 Copilot**

17. Press **F1** to open Command Palette; take screenshot
18. Type `debug: Select and Start Debugging` → press Enter to select the command; take screenshot
19. Type `Preview Local in Copilot (chrome)` in the debug configuration dropdown → press Enter; take screenshot
20. In the newly launched Chrome browser, wait up to 120 s for the Microsoft sign-in page to appear; take screenshot showing the Email/phone field
21. Type `M365_ACCOUNT_NAME` → click **"Next"**; type `M365_ACCOUNT_PASSWORD` → click **"Sign in"** (`delay: 3 s`); take screenshot
22. On the "Stay signed in?" prompt, press **Enter** to confirm Yes (`delay: 3 s`); take screenshot

**Phase 5 – Validate agent in M365 Copilot**

23. Press **Ctrl+-** to zoom out the Copilot page (`delay: 30 s`, applied automatically by test runner); take screenshot of the Copilot interface
24. Click the **Reload** button in the Chrome address bar to refresh M365 Copilot; take screenshot
25. Assert that the DA agent named `<app_name>local` is visible in the M365 Copilot interface (retry up to 30 s); take screenshot
26. Click the **"Message Copilot"** input box; type `how can you assistant me?` → press Enter; take screenshot
27. Assert that a response from the agent appears in the chat (retry up to 60 s); take screenshot

**Expected result:**
- Project is created with `appPackage/manifest.json` and `appPackage/declarativeAgent.json` (no `ai-plugin.json` because No Action)
- The agent `<app_name>local` is listed and accessible in M365 Copilot after local debug launch
- The agent returns a non-empty response to the `how can you assistant me?` message

**Pass criteria:**
- Agent entry `<app_name>local` is visible in the M365 Copilot agent list (assertion with 30 s retry)
- At least one response message from the bot is visible in the Copilot chat after sending the query (assertion with 60 s retry)

**Test script:**
`packages/tests/src/da-no-action-local-debug.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                                       | Pass condition                                              | Why                                                                            |
|-----|---------------------------------------|-----------------------------------------------------------------------|-------------------------------------------------------------|--------------------------------------------------------------------------------|
| 01  | `01-atk-sidebar.png`                  | ATK sidebar after closing the welcome dialog                          | ATK panel visible; no welcome dialog                        | Confirms clean starting state for project creation                             |
| 02  | `02-new-project-quickpick.png`        | New Project QuickPick open with app-type options                      | QuickPick visible with options                              | Confirms "Create a New Agent/App" was triggered                                |
| 03  | `03-da-selected.png`                  | "Declarative Agent" highlighted in QuickPick                          | Option visible and focused                                  | Confirms DA path entered                                                       |
| 04  | `04-no-action-selected.png`           | "No Action" highlighted in the add-action QuickPick                   | Option visible and focused                                  | Confirms No Action variant chosen (vs. add-action variants)                    |
| 05  | `05-workspace-folder.png`             | Workspace Folder QuickPick with "Default folder"                      | Default folder option present                               | Confirms project root selection step reached                                   |
| 06  | `06-app-name-entered.png`             | Application Name InputBox with typed name                             | Name text visible in input                                  | Records app name used throughout the test                                      |
| 07  | `07-readme-tab.png`                   | VS Code editor with "Preview README.md" tab open                      | README.md preview tab visible                               | Confirms project was created and README auto-opened                            |
| 08  | `08-readme-closed.png`                | VS Code editor after closing README tab                               | README.md tab no longer present                             | Confirms cleanup of auto-opened tab before sign-in phase                       |
| 09  | `09-notifications-cleared.png`        | VS Code with all notifications cleared                                | No notification banners or bell badge                       | Confirms clean state before M365 sign-in to avoid interference                 |
| 10  | `10-m365-signin-browser.png`          | Microsoft web sign-in page in Chrome                                  | Email/phone input field visible                             | Confirms M365 sign-in flow launched from ATK                                   |
| 11  | `11-m365-signin-complete.png`         | ATK panel showing signed-in M365 account                              | Account name visible in ATK Accounts section                | Confirms M365 login succeeded before debug launch                              |
| 12  | `12-debug-command-palette.png`        | Command Palette with "debug: Select and Start Debugging" typed        | Command option visible in dropdown                          | Confirms debug launch path taken via Command Palette                           |
| 13  | `13-debug-config-selected.png`        | Debug configuration dropdown with "Preview Local in Copilot" focused  | Option visible and selected                                 | Confirms correct debug target (Copilot, not Teams)                             |
| 14  | `14-copilot-login-page.png`           | Microsoft sign-in page in Chrome (second login for Copilot session)   | Email/phone field visible                                   | Confirms Copilot-side sign-in step reached; precondition wait 120 s            |
| 15  | `15-copilot-login-complete.png`       | M365 Copilot interface loading in Chrome after sign-in                | Copilot chat UI visible                                     | Confirms sign-in and "Stay signed in?" completed successfully                  |
| 16  | `16-copilot-zoomed-out.png`           | M365 Copilot at reduced zoom level (delay 30 s applied)               | Full Copilot chat UI visible without clipping               | Confirms page is stable after zoom and reload; agent list accessible           |
| 17  | `17-agent-visible.png`                | M365 Copilot showing the DA agent entry in the app list               | `<app_name>local` agent visible (retry 30 s)                | Primary pass criterion: local DA is provisioned and visible in Copilot         |
| 18  | `18-query-sent.png`                   | Copilot chat with "how can you assistant me?" submitted               | Message text visible in chat history                        | Records the test query that verifies the agent is addressable                  |
| 19  | `19-agent-response.png`               | Copilot chat showing agent response                                   | Non-empty response bubble from agent visible (retry 60 s)   | Primary end-to-end assertion: DA No Action agent responds correctly in Copilot |

---

## Notes

- This is a **local debug** test (no Azure provisioning). The debug surface is **M365 Copilot** (`Preview Local in Copilot (Chrome)`), not Teams.
- The test creates a brand-new project from scratch (VS Code must have no project open at start).
- Two separate M365 sign-in flows occur: one in Phase 3 (ATK account sign-in via Command Palette) and one in Phase 4 (browser sign-in when Copilot launches). Both use the same credentials.
- `step_63aec6cc` has `precondition_wait_timeout: 120` — the browser sign-in page can take up to 2 minutes to appear after the debug session starts; the test runner waits accordingly.
- `step_904a54b9` has `delay: 30` and `force_run: true` — after Copilot loads, the test unconditionally waits 30 s before zooming out and reloading to let the agent registration propagate.
- `step_c1889d91` has `step_retry_timeout: 30` — the agent entry assertion retries for up to 30 s.
- `plan_r_1023_054001` (`validation_DA_no_action`) final assertion (`step_fc503e63`) has `step_retry_timeout: 60` — the bot response can take up to 60 s to appear in Copilot chat.
- A "No Action" DA has no `ai-plugin.json`; the scaffold contains only `manifest.json` and `declarativeAgent.json`. The validation query is generic (`how can you assistant me?`) because there are no plugin-defined intents to exercise.
