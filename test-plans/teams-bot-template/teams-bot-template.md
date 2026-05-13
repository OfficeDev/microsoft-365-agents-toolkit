# Test Plan: Teams Bot Template Creation

## Metadata
- **feature-slug**: `teams-bot-template`
- **owner**: atk-qa
- **created**: 2026-05-09
- **updated**: 2026-05-13 (aligned with verified ATK v6.8.0 wizard flow)
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
- Multi-language (JS/Python) - separate TCs

---

## Test Cases

### TC-001 - Create Teams Bot template (TypeScript)

**Preconditions:**
- VSCode is open with no project loaded
- ATK extension v6.8.0+ is installed and activated
- Extension sidebar shows "Microsoft 365 Agents Toolkit"

**Actual wizard flow (ATK v6.8.0, verified):**
The `fx-extension.create` command opens a multi-step QuickPick wizard in this order:
1. **App category** - select "Teams Agents and Apps"
2. **App type** - select "Bot"
3. **Bot variant** - select "Simple Bot"
4. **Language** - select "TypeScript"
5. **Workspace folder** - select "Default folder" (~~/AgentsToolkitProjects)
6. **Application Name** - InputBox, type the project name, press Enter

> Note: Folder is selected BEFORE app name (not after).
> "Simple Bot" is the correct option (not "Basic Bot").

**Steps:**
1. ATK extension activates
2. Fire `fx-extension.create` command (no Command Palette - called directly)
3. Wait for "Teams Agents and Apps" QuickPick to appear, screenshot
4. Click "Teams Agents and Apps", screenshot before click
5. Click "Bot", screenshot before click
6. Click "Simple Bot", screenshot before click
7. Click "TypeScript", screenshot before click
8. Click "Default folder", screenshot before click
9. Type app name `test-teams-bot-001`, press Enter
10. Wait 90s for scaffold + new window to open
11. Screenshot final state
12. Assert these files exist in `~/AgentsToolkitProjects/test-teams-bot-001/`:
    - `m365agents.yml`
    - `package.json`
    - `index.ts`
    - `appPackage/manifest.json`

**Expected result:**
- Wizard completes without error
- All 4 asserted files are present at project root (NOT in src/)
- VSCode opens the new project folder

**Test script:**
`packages/tests/src/ui-test/copilot-driven/teams-bot-create-template.test.ts`

**Screenshots produced by test script:**
| ID | File | Description |
|---|---|---|
| 01 | `01-extension-active.png` | VSCode at launch with ATK active |
| 02 | `02-wizard-open.png` | First QuickPick visible after command fires |
| 03 | `03-teams-agents-apps.png` | QuickPick showing "Teams Agents and Apps" option |
| 04 | `04-bot-selected.png` | QuickPick showing Bot option |
| 05 | `05-simple-bot.png` | QuickPick showing Simple Bot option |
| 06 | `06-typescript.png` | QuickPick showing TypeScript option |
| 07 | `07-workspace-folder.png` | QuickPick showing Default folder + Browse |
| 08 | `08-app-name-input.png` | InputBox for Application Name (empty, before typing) |
| 09 | `09-project-created.png` | State after scaffold completes |
| 10 | `10-final-state.png` | Final file verification state |

---

### TC-002 - Project name validation (spaces not allowed)

**Preconditions:** Same as TC-001.

**Steps:**
1. Open "Teams: Create New App"
2. Navigate to the app-name step
3. Enter `my bot app` (contains spaces)
4. Screenshot the validation state

**Expected result:**
- An inline validation error appears: "Project name must not contain spaces"
- OR the name is auto-sanitised to `my-bot-app`
- Wizard does NOT create a broken scaffold

**Test script:**
`packages/tests/src/ui-test/copilot-driven/teams-bot-name-validation.test.ts`