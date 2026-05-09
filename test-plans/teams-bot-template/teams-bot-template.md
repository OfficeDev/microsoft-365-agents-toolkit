# Test Plan: Teams Bot Template Creation

## Metadata
- **feature-slug**: `teams-bot-template`
- **owner**: atk-qa
- **created**: 2026-05-09
- **triggers**: issue-label `atk-copilot-test`, manual

## Scope

**Covers:**
- Opening VSCode with the ATK extension installed
- Running "Teams: Create New App" via the Command Palette
- Selecting "Bot" as the app type
- Choosing TypeScript as the language
- Entering a project name
- Verifying the scaffold is created (key files present)

**Does NOT cover:**
- Local debug / F5 run
- Azure provisioning
- Multi-language (JS/Python) – separate TCs

---

## Test Cases

### TC-001 – Create Teams Bot template (TypeScript)

**Preconditions:**
- VSCode is open with no project loaded
- ATK extension is installed and activated
- Extension sidebar shows "Microsoft 365 Agents Toolkit"

**Steps:**
1. Screenshot startup state
2. Open Command Palette (`Ctrl+Shift+P`)
3. Type and run "Teams: Create New App"
4. In the wizard, select **"Teams Agent & Apps"**
5. Select **"Bot"**
6. Select **"Basic Bot"**
7. Select **"TypeScript"**
8. Enter app name: `test-teams-bot-001`
9. Choose output folder `/tmp/atk-test-output/projects`
10. Wait for scaffold to complete
11. Screenshot the file explorer showing the generated project tree
12. Assert these files exist in the scaffold:
    - `src/index.ts`
    - `teamsapp.yml`
    - `package.json`
    - `appPackage/manifest.json`

**Expected result:**
- Wizard completes without error
- All 4 asserted files are present
- VSCode opens the new project folder
- No error notifications are shown

**Test script:**
`packages/tests/src/ui-test/copilot-driven/teams-bot-create-template.test.ts`

**Screenshots expected:**
| ID | Description |
|---|---|
| `01-startup` | VSCode at launch |
| `02-command-palette` | Command Palette open |
| `03-wizard-app-type` | Wizard step: select app type |
| `04-wizard-bot` | Wizard step: select Bot |
| `05-wizard-language` | Wizard step: select TypeScript |
| `06-wizard-name` | Wizard step: enter project name |
| `07-scaffold-complete` | File explorer showing generated project |

---

### TC-002 – Project name validation (spaces not allowed)

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