# Test Plan: DA No Action – Scaffold and Local Debug (VS Code UX)

## Metadata

- **feature-slug**: `da-no-action-local-debug`
- **owner**: v-helzha
- **created**: 2026-06-01
- **updated**: 2026-06-03
- **triggers**: issue-label `atk-copilot-test`, manual

## Scope

**Covers:**
- Creating a new Declarative Agent (No Action) project via the ATK wizard in VS Code
- Verifying all expected scaffold files exist: `m365agents.yml`, `m365agents.local.yml`, `appPackage/manifest.json`, `appPackage/declarativeAgent.json`, `.vscode/tasks.json`, `.vscode/launch.json`
- Opening and viewing key scaffold files in the VS Code editor to confirm correct content
- Launching local debug via **"Preview Local in Copilot (Chrome)"** and observing each VS Code-internal debug task:
  - "Start Agent Locally" compound task starts
  - "Validate prerequisites" subtask runs (checks M365 Copilot access)
  - "Create resources" subtask runs (ATK provision: zip → validate → update appPackage)
- Observing the Chrome browser window open with the M365 Copilot URL

**Does NOT cover:**
- M365 browser sign-in interaction (no credential entry steps in this plan)
- Sending chat messages or verifying agent response in Copilot
- Declarative Agent with actions (`da-add-action-*` plans)
- CLI-based project creation or remote provision to Azure

---

## Test Cases

### TC-001 – Create DA (No Action), verify scaffold files, and launch local debug

**Preconditions:**
- VS Code is open with no project loaded (fresh state; ATK sign-in welcome dialog may be visible)
- ATK extension v6.8.0+ is installed and activated
- M365 account credentials available in environment (`M365_ACCOUNT_NAME` / `M365_ACCOUNT_PASSWORD`) — needed only for the debug task to reach "Create resources" step

**App name:** Generate at test start: `` const appName = `da-no-action-${Date.now()}` ``. Use as-is; no spaces.

**Wizard flow:**

| Step | VS Code command / QuickPick / InputBox             | Value to select / type      | Wait-for-text before click     |
|------|----------------------------------------------------|-----------------------------|--------------------------------|
| 1    | `executeCommand("notifications.clearAll")`         | — (dismiss welcome dialog)  | —                              |
| 2    | `executeCommand("fx-extension.create")` (no await) | — (opens wizard)            | —                              |
| 3    | New Project QuickPick — App type                   | `Declarative Agent`         | wait `"Declarative Agent"` 60 s |
| 4    | DA — Add action choice                             | `No Action`                 | wait `"No Action"` 20 s        |
| 5    | Workspace Folder QuickPick                         | `Default folder`            | wait `"Default folder"` 15 s   |
| 6    | Application Name InputBox                          | `${appName}` + Enter        | wait `"Application Name"` 15 s |

**Dynamic data:**
- `const appName = \`da-no-action-${Date.now()}\`` — alphanumeric, no spaces

**Surface tags:** `[VSC]` = VS Code workbench/QuickPick. `[Chrome]` = external Chrome browser.

**Steps:**

**Phase 1 – Create project** `[VSC]`

1. `[VSC]` Dismiss the ATK welcome notification: `executeCommand("notifications.clearAll")`; `takeScreenshot("01-atk-fresh")`
2. `[VSC]` Fire create-project wizard **without await**: `executeCommand("fx-extension.create")`; `wait(500)`
3. `[VSC]` `sendSignal("waitForTextThenScreenshot:Declarative Agent:60000:02-wizard-da", 68000)`; `sendSignal("clickText:Declarative Agent", 10000)`; `wait(1000)`
4. `[VSC]` `sendSignal("waitForTextThenScreenshot:No Action:20000:03-wizard-no-action", 28000)`; `sendSignal("clickText:No Action", 10000)`; `wait(1000)`
5. `[VSC]` `sendSignal("waitForTextThenScreenshot:Default folder:15000:04-wizard-folder", 23000)`; `sendSignal("clickText:Default folder", 10000)`; `wait(1000)`
6. `[VSC]` `sendSignal("waitForTextThenScreenshot:Application Name:15000:05-wizard-appname", 23000)`; `sendSignal("type:${appName}", 8000)`; `wait(500)`; `sendSignal("pressKey:Enter", 5000)`

**Phase 2 – Verify scaffold files** `[VSC]`

7. `[VSC]` Poll `~/AgentsToolkitProjects/${appName}/m365agents.yml` for up to 90 s; `takeScreenshot("06-scaffold-explorer")` once found; fail TC if still missing after 90 s
8. `[VSC]` Assert all required files exist — fail TC if any are absent:
   - `m365agents.yml` ✓
   - `m365agents.local.yml` ✓
   - `appPackage/manifest.json` ✓
   - `appPackage/declarativeAgent.json` ✓
   - `.vscode/tasks.json` ✓
   - `.vscode/launch.json` ✓
9. `[VSC]` Assert `appPackage/ai-plugin.json` does **NOT** exist (No Action DA has no plugin); record pass/fail
10. `[VSC]` Open `m365agents.yml` in editor: `executeCommand("vscode.open", uri("m365agents.yml"))`; `wait(2000)`; `takeScreenshot("07-m365agents-yml")`
11. `[VSC]` Open `.vscode/tasks.json` in editor: `executeCommand("vscode.open", uri(".vscode/tasks.json"))`; `wait(2000)`; `takeScreenshot("08-tasks-json")`
12. `[VSC]` Open `.vscode/launch.json` in editor: `executeCommand("vscode.open", uri(".vscode/launch.json"))`; `wait(2000)`; `takeScreenshot("09-launch-json")`
13. `[VSC]` Close all open editors: `executeCommand("workbench.action.closeAllEditors")`; `wait(500)`; `executeCommand("notifications.clearAll")`; `wait(500)`; `takeScreenshot("10-editors-closed")`

**Phase 3 – Launch local debug and observe VS Code tasks** `[VSC → Chrome]`

14. `[VSC]` Fire debug-selection command **without await**: `executeCommand("workbench.action.debug.selectandstart")`; `wait(500)`
15. `[VSC]` `sendSignal("waitForTextThenScreenshot:Preview Local in Copilot:20000:11-debug-picker", 28000)`; `sendSignal("clickText:Preview Local in Copilot (Chrome)", 10000)`; `wait(1000)`
16. `[VSC]` `sendSignal("waitForTextThenScreenshot:Start Agent Locally:30000:12-task-started", 38000)` — terminal panel shows compound task beginning; `takeScreenshot("12-task-started")`
17. `[VSC]` `sendSignal("waitForTextThenScreenshot:Validate prerequisites:30000:13-validate-prereqs", 38000)` — terminal shows "Validate prerequisites" subtask output; `takeScreenshot("13-validate-prereqs")`
18. `[VSC]` `sendSignal("waitForTextThenScreenshot:Create resources:60000:14-create-resources", 68000)` — terminal shows "Create resources" (ATK provision) subtask output; `takeScreenshot("14-create-resources")`
19. `[VSC]` `wait(10000)` for provision to progress; `takeScreenshot("15-provision-progress")`
20. `[Chrome]` `sendSignal("waitForTextThenScreenshot:m365.cloud.microsoft:120000:16-copilot-browser", 130000)` — Chrome window opens with M365 Copilot URL; `takeScreenshot("16-copilot-browser")`

**Expected result:**
- All 6 scaffold files exist in `~/AgentsToolkitProjects/${appName}/` within 90 s
- `appPackage/ai-plugin.json` is absent (No Action variant)
- `m365agents.yml` opened in editor shows ATK lifecycle YAML with `provision:` block containing `teamsApp/create`, `teamsApp/zipAppPackage`, `copilotAgent/publish` actions
- `.vscode/tasks.json` opened in editor shows `"Start Agent Locally"` compound task with `"Validate prerequisites"` and `"Create resources"` subtasks
- `.vscode/launch.json` shows `"Preview Local in Copilot (Chrome)"` compound configuration with `preLaunchTask: "Start Agent Locally"`
- VS Code terminal panel shows each task running in sequence
- Chrome opens to `m365.cloud.microsoft` after tasks complete

**Pass criteria:**
- `m365agents.yml` found on disk within 90 s after Enter on app name (step 7)
- All 5 other required files present (step 8)
- `appPackage/ai-plugin.json` absent (step 9)
- `"Start Agent Locally"` text appears in VS Code terminal within 30 s of debug start (step 16)
- `"Validate prerequisites"` text appears in terminal within 30 s (step 17)
- `"Create resources"` text appears in terminal within 60 s (step 18)
- Chrome window opens to `m365.cloud.microsoft` within 120 s (step 20)

**Test script:**
`packages/tests/copilot-test/src/da-no-action-local-debug.test.ts`

**Screenshots produced by test:**

| ID  | Filename                      | What is visible                                                         | Pass condition                                    | Why                                                             |
|-----|-------------------------------|-------------------------------------------------------------------------|---------------------------------------------------|-----------------------------------------------------------------|
| 01  | `01-atk-fresh.png`            | VS Code fresh state, ATK sidebar visible, no project                    | ATK panel visible, no project in explorer         | Confirms clean test starting state                              |
| 02  | `02-wizard-da.png`            | New Project QuickPick with "Declarative Agent" highlighted              | `"Declarative Agent"` text present                | Confirms wizard opened to DA selection step                     |
| 03  | `03-wizard-no-action.png`     | DA action-choice QuickPick with "No Action" highlighted                 | `"No Action"` text present                        | Confirms No Action variant is available                         |
| 04  | `04-wizard-folder.png`        | Workspace Folder QuickPick with "Default folder"                        | `"Default folder"` text present                   | Confirms workspace selection step reached                       |
| 05  | `05-wizard-appname.png`       | Application Name InputBox                                               | `"Application Name"` text visible                 | Records app name entry step                                     |
| 06  | `06-scaffold-explorer.png`    | VS Code Explorer showing created project with `m365agents.yml` visible  | `m365agents.yml` in file tree                     | **Primary scaffold assertion**: file tree confirms creation     |
| 07  | `07-m365agents-yml.png`       | `m365agents.yml` open in editor showing provision lifecycle YAML        | `teamsApp/create` and `copilotAgent/publish` text | Confirms ATK lifecycle file content is correct                  |
| 08  | `08-tasks-json.png`           | `.vscode/tasks.json` open in editor showing `"Start Agent Locally"`     | `"Start Agent Locally"` task text present         | Confirms debug task chain wiring is correct                     |
| 09  | `09-launch-json.png`          | `.vscode/launch.json` showing `"Preview Local in Copilot (Chrome)"`     | compound config with `preLaunchTask` visible      | Confirms debug compound config is correct                       |
| 10  | `10-editors-closed.png`       | VS Code with all editors closed, clean explorer view                    | No open editor tabs                               | Confirms clean state before debug launch                        |
| 11  | `11-debug-picker.png`         | Debug configuration dropdown with `"Preview Local in Copilot (Chrome)"` | `"Preview Local in Copilot"` text present         | Confirms correct debug target is selectable via the picker      |
| 12  | `12-task-started.png`         | VS Code terminal panel showing `"Start Agent Locally"` task output      | Terminal visible with task name                   | Confirms VS Code executed the `preLaunchTask` compound task     |
| 13  | `13-validate-prereqs.png`     | Terminal showing `"Validate prerequisites"` subtask output              | `"Validate prerequisites"` text in terminal       | Confirms M365 Copilot access check step ran                     |
| 14  | `14-create-resources.png`     | Terminal showing `"Create resources"` subtask output (ATK provision)    | `"Create resources"` text in terminal             | Confirms ATK provision (zip/validate/update/publish) started    |
| 15  | `15-provision-progress.png`   | Terminal showing ATK provision in progress (teamsApp steps)             | ATK provision output visible                      | Documents provision step output for debugging failures          |
| 16  | `16-copilot-browser.png`      | Chrome window opened to `m365.cloud.microsoft` URL `[Chrome]`           | `m365.cloud.microsoft` in address bar             | Confirms entire debug task chain completed and browser launched |

---

## Notes

- This is a **VS Code UX** test — it verifies the scaffold output and VS Code internal debug task execution, not the Copilot chat end-to-end.
- The test creates a brand-new project from scratch (VS Code must have no project open at start).
- **App name**: generate at test start as `` `da-no-action-${Date.now()}` ``. Do NOT hardcode.
- **VS Code command IDs** used:
  - `notifications.clearAll` — dismiss welcome dialog and clear notifications
  - `fx-extension.create` — open the New Project wizard (fire **without await**)
  - `vscode.open` — open a file URI in the editor (for viewing scaffold files)
  - `workbench.action.closeAllEditors` — close all open editor tabs
  - `workbench.action.debug.selectandstart` — open debug config picker (fire **without await**)
- **Right-click is not supported** by the signal mechanism.
- **Scaffold file location**: ATK creates projects in `~/AgentsToolkitProjects/${appName}` by default when "Default folder" is selected.
- **DA no-action local debug task chain** (`tasks.json`):
  1. `"Start Agent Locally"` (compound, `dependsOrder: "sequence"`)
  2. → `"Validate prerequisites"` — `debug-check-prerequisites` with `prerequisites: ["copilotAccess"]`
  3. → `"Create resources"` — `provision` with `env: "local"` (runs `m365agents.local.yml`)
- **`m365agents.local.yml`** lifecycle: `teamsApp/create` → `teamsApp/zipAppPackage` → `teamsApp/validateAppPackage` → `teamsApp/update` → `copilotAgent/publish`
- **No `npm install` step** — DA no-action has no bot runtime code, so there is no dependency installation task. The local debug is purely ATK provision + browser launch.
- **Chrome launch**: after provision completes, VS Code launches Chrome with URL `https://m365.cloud.microsoft/chat/entity1-...?auth=2&developerMode=Basic`
- **`appPackage/ai-plugin.json` absence** is an explicit pass criterion — No Action DA contains only `manifest.json` and `declarativeAgent.json`.
