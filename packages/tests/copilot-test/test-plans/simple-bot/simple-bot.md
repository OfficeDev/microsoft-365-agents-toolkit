# Test Plan: Simple Bot Creation

## Metadata
- **feature-slug**: `simple-bot`
- **owner**: atk-qa
- **created**: 2026-05-09
- **updated**: 2026-05-13
- **triggers**: issue-label `atk-copilot-test`, manual

## Scope

**Covers:**
- Opening VSCode with the ATK extension installed
- Running `fx-extension.create` command (Teams: Create New App)
- Selecting "Teams Agents and Apps" > "Bot" > "Simple Bot" > "TypeScript"
- Choosing default workspace folder, then entering a project name
- Verifying the scaffold is created (key files present)

**Does NOT cover:**
- Local debug / F5 run
- Azure provisioning
- Multi-language (JS/Python) – separate TCs
- Project name validation (spaces/special chars) – TC-002 (not yet implemented)

---

## Test Cases

### TC-001 – Create Simple Bot (TypeScript)

**Preconditions:**
- VSCode is open with no project loaded
- ATK extension v6.8.0+ is installed and activated
- Extension sidebar shows "Microsoft 365 Agents Toolkit"

**Wizard flow (verified on ATK v6.8.0):**

| Step | QuickPick / InputBox | Value to select/type |
|------|----------------------|----------------------|
| 1    | App category         | Teams Agents and Apps |
| 2    | App type             | Bot                  |
| 3    | Bot variant          | Simple Bot           |
| 4    | Language             | TypeScript           |
| 5    | Workspace folder     | Default folder       |
| 6    | Application Name     | `test-teams-bot-001` (type + Enter) |

> **Note:** Workspace folder is selected **before** Application Name (step 5 before step 6).

**Steps:**
1. ATK extension activates
2. Fire `fx-extension.create` command directly (not via Command Palette)
3. Wait for "Teams Agents and Apps" QuickPick to appear; take screenshot
4. Screenshot → click "Teams Agents and Apps"
5. Screenshot → click "Bot"
6. Screenshot → click "Simple Bot"
7. Screenshot → click "TypeScript"
8. Screenshot → click "Default folder"
9. Screenshot → type app name `test-teams-bot-001`, press Enter
10. Wait up to 90 s for scaffold + new window to open
11. Take final state screenshot
12. Assert files exist in `~/AgentsToolkitProjects/test-teams-bot-001/`

**Expected result:**
- Wizard completes without error
- All 4 asserted files are present at project root (not inside `src/`):
  - `m365agents.yml`
  - `package.json`
  - `index.ts`
  - `appPackage/manifest.json`
- VSCode opens the new project folder automatically

**Test script:**
`packages/tests/src/simple-bot-create.test.ts`

**Screenshots produced by test:**

| ID  | Filename                  | What is visible                                 |
|-----|---------------------------|-------------------------------------------------|
| 01  | `01-extension-active.png` | VSCode at launch, ATK sidebar active            |
| 02  | `02-wizard-open.png`      | First QuickPick open after command fires        |
| 03  | `03-teams-agents-apps.png`| "Teams Agents and Apps" option highlighted      |
| 04  | `04-bot-selected.png`     | "Bot" option highlighted                        |
| 05  | `05-simple-bot.png`       | "Simple Bot" option highlighted                 |
| 06  | `06-typescript.png`       | "TypeScript" option highlighted                 |
| 07  | `07-workspace-folder.png` | "Default folder" + Browse option visible        |
| 08  | `08-app-name-input.png`   | Application Name InputBox (empty, before typing)|
| 09  | `09-project-created.png`  | State immediately after scaffold completes      |
| 10  | `10-final-state.png`      | Final file-tree verification state             |
