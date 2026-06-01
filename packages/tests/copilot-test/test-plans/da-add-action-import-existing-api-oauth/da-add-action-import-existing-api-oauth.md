# Test Plan: DA Add Action – Import Existing API (Basic OAuth)

## Metadata

- **feature-slug**: `da-add-action-import-existing-api-oauth`
- **owner**: v-helzha
- **created**: 2026-05-11
- **updated**: 2026-06-01
- **triggers**: issue-label `atk-copilot-test`, manual
- **source-plan-id**: `plan_7ca218e3`
- **workitem**: `31434597`

## Scope

**Covers:**
- Creating a new Declarative Agent project with an OpenAPI action (OAuth-authenticated spec) from scratch
- Selecting operations from the loaded OAuth spec (OK → Clear All flow)
- M365 account sign-in during project creation
- Provisioning the DA to Azure (via Command Palette), selecting the `dev` environment
- Entering OAuth registration credentials (client ID and client secret) that appear mid-provision
- Confirming the OAuth registration dialog
- Launching the provisioned DA in M365 Copilot (remote debug / Preview in Copilot)
- Sending a repair query and handling the Allow-action permission prompt
- Verifying the agent presents the OAuth **"Sign in to Repair Service"** button in Copilot chat

**Does NOT cover:**
- No-auth OpenAPI import (covered by `DA_Add_Action_Import_Existing_API_Basic_No_Auth`)
- Local debug flow
- CLI-based project creation or add-action flow
- Azure resource clean-up after the test

---

## Test Cases

### TC-001 – Create DA with OAuth API action, provision, and verify sign-in button in Copilot

**Preconditions:**
- VS Code is open with no project loaded (fresh state, sign-in welcome dialog may be visible)
- ATK extension v6.8.0+ is installed and activated
- M365 account credentials are available (`M365_ACCOUNT_NAME` / `M365_ACCOUNT_PASSWORD`)
- Azure subscription is available for provisioning

**Wizard flow:**

| Step | QuickPick / InputBox                              | Value to select / type |
|------|---------------------------------------------------|------------------------|
| 1    | ATK sidebar — "Create a New Agent/App"            | click button |
| 2    | New Project — App type                            | Declarative Agent |
| 3    | DA — Add action choice                            | Add an Action (or equivalent option leading to action source) |
| 4    | Create an Action — Action source                  | Start with an OpenAPI Description Document |
| 5    | OpenAPI Spec Document source                      | Enter OpenAPI Document URL |
| 6    | Enter OpenAPI Description Document URL InputBox   | `https://raw.githubusercontent.com/SLdragon/example-openapi-spec/refs/heads/main/real-oauth.yaml` (type + Enter) |
| 7    | Select Operation(s) Copilot Can Interact with     | click **OK** |
| 8    | Select Operation(s) — clear selection             | click **Clear All** |
| 9    | Workspace Folder                                  | Default folder |
| 10   | Application Name InputBox                         | `vscuse_app_#####` (auto-generated; type + Enter) |
| 11   | Sign in to Microsoft 365 (mid-creation)           | full web login (`M365_ACCOUNT_NAME` / `M365_ACCOUNT_PASSWORD`) |
| 12   | Command Palette — Provision                       | F1 → type `provision` → Enter |
| 13   | Select an environment                             | `dev` |
| 14   | OAuth registration client ID InputBox [delay 30s] | `fake_client_id` (type + Enter) |
| 15   | OAuth registration client secret InputBox         | `fake_clientsecret` (type + Enter) |
| 16   | OAuth registration confirmation dialog            | click **Confirm** |
| 17   | Command Palette — Debug: Select and Start Debugging | F1 → type `debug: Select and Start Debugging` → Enter |
| 18   | Debug configuration dropdown                      | Preview in Copilot (Chrome) |
| 19   | M365 Copilot chat — send query                    | type `show repair records assigned to karin blair` → Enter |
| 20   | Allow-action permission prompt                    | click first **Allow** button → click blue **Allow** |

**Steps:**

**Phase 1 – Create project**
1. VS Code opens; close the ATK sign-in welcome dialog (red close button); take screenshot confirming ATK sidebar is visible
2. Click **"Create a New Agent/App"** in the ATK sidebar; take screenshot of the New Project QuickPick
3. Select **"Declarative Agent"**; take screenshot
4. Select the action choice that leads to adding an OpenAPI action (e.g., "Add an Action"); take screenshot
5. Select **"Start with an OpenAPI Description Document"**; take screenshot
6. Select **"Enter OpenAPI Document URL"**; take screenshot of the URL InputBox
7. Type `https://raw.githubusercontent.com/SLdragon/example-openapi-spec/refs/heads/main/real-oauth.yaml` and press Enter; take screenshot

**Phase 2 – Select operations and set project location**

8. Wait up to 60 s for the **"Select Operation(s) Copilot Can Interact with"** dialog to load the spec; take screenshot showing loaded operations
9. Click **OK** to accept the initial selection; take screenshot
10. Click **Clear All** to deselect all operations; take screenshot confirming cleared state
11. Click **Default folder** in the Workspace Folder dialog; take screenshot
12. Type the application name and press Enter; take screenshot
13. Close the auto-opened **Preview README.md** tab; open Command Palette and clear all notifications

**Phase 3 – Sign in to M365**

14. Open Command Palette (F1), type `View: Show Microsoft 365 Agents Toolkit`, press Enter
15. Open Command Palette (F1), type `Microsoft 365 Agents: Accounts`, press Enter, select the Accounts command
16. Click **"Sign in to Microsoft 365"**; click the sign-in button in the modal
17. Complete the Microsoft web sign-in: enter `M365_ACCOUNT_NAME` → Next → enter `M365_ACCOUNT_PASSWORD` → Enter; take screenshot of successful sign-in
18. Close the sign-in browser tab; take screenshot confirming M365 account shown in ATK

**Phase 4 – Provision**

19. Open Command Palette (F1), type `provision`, press Enter; take screenshot of provision starting
20. When "Select an environment" appears, click **dev**; take screenshot
21. Wait up to 30 s for the **OAuth registration client ID** InputBox to appear; take screenshot
22. Type `fake_client_id` and press Enter; take screenshot
23. When the **OAuth registration client secret** InputBox appears, type `fake_clientsecret` and press Enter; take screenshot
24. When the **OAuth registration confirmation dialog** appears, click **Confirm**; take screenshot
25. Open the notifications panel (F1 → `Notifications: Show Notifications`); take screenshot
26. Assert the notification **"provision stage executed successfully"** is present; take screenshot

**Phase 5 – Remote debug in M365 Copilot**

27. Open Command Palette (F1), type `debug: Select and Start Debugging`, press Enter
28. Type `Preview in Copilot (Chrome)` and click that option; take screenshot
29. Complete the second M365 sign-in in the browser (same credentials); click **Yes** on "Stay signed in?"
30. Press Ctrl+- to reduce zoom; press F5 to refresh Copilot; take screenshot of the Copilot interface

**Phase 6 – Validate OAuth action in Copilot**

31. Type `show repair records assigned to karin blair` in the Copilot chat input and press Enter; take screenshot
32. Assert that an **Allow** button is visible in the Copilot response; take screenshot
33. Click the first **Allow** button; then click the blue **Allow** confirmation button; take screenshot
34. Wait up to 30 s (with up to 60 s retry) for the final state; take screenshot
35. Assert that a button labeled **"Sign in to Repair Service"** is visible in the chat

**Expected result:**
- Project is created with `appPackage/manifest.json`, `appPackage/declarativeAgent.json`, and `appPackage/ai-plugin.json` referencing the OAuth spec
- Provision completes with "provision stage executed successfully" notification
- M365 Copilot chat shows a **"Sign in to Repair Service"** button after the repair query and Allow confirmation

**Pass criteria:**
- "provision stage executed successfully" notification appears in VS Code after provision
- A button labeled **"Sign in to Repair Service"** is visible in M365 Copilot chat (assertion at delay 30 s, retry up to 60 s)

**Test script:**
`packages/tests/src/da-add-action-import-existing-api-oauth.test.ts`

**Screenshots produced by test:**

| ID  | Filename                               | What is visible                                                   | Pass condition                                              | Why                                                                     |
|-----|----------------------------------------|-------------------------------------------------------------------|-------------------------------------------------------------|-------------------------------------------------------------------------|
| 01  | `01-atk-sidebar.png`                   | ATK sidebar after closing welcome dialog                          | ATK sidebar panel visible                                   | Confirms clean starting state                                           |
| 02  | `02-new-project-quickpick.png`         | New Project QuickPick open                                        | QuickPick with app type options visible                     | Confirms "Create a New Agent/App" fired                                 |
| 03  | `03-da-selected.png`                   | "Declarative Agent" highlighted in QuickPick                      | Option visible and focused                                  | Confirms DA path entered                                                |
| 04  | `04-action-choice.png`                 | DA action-choice QuickPick                                        | Option leading to OpenAPI action visible                    | Confirms add-action branch taken                                        |
| 05  | `05-openapi-source.png`                | "Start with an OpenAPI Description Document" highlighted          | Option visible and focused                                  | Confirms OpenAPI import branch selected                                 |
| 06  | `06-enter-url-option.png`              | "Enter OpenAPI Document URL" option highlighted                   | Option visible                                              | Confirms URL input path chosen                                          |
| 07  | `07-url-input-filled.png`              | URL InputBox with OAuth spec URL typed                            | Full URL visible in input field                             | Records exact spec URL used                                             |
| 08  | `08-operations-loaded.png`             | "Select Operation(s)" dialog with operations loaded               | Operations list visible; OK button present                  | Proves spec was fetched and parsed from the remote URL                  |
| 09  | `09-ok-clicked.png`                    | Operations dialog after clicking OK                               | Dialog advancing                                            | Confirms OK interaction recorded                                        |
| 10  | `10-clear-all.png`                     | Operations dialog with all items deselected                       | No checked operations visible                               | Confirms Clear All interaction                                          |
| 11  | `11-workspace-folder.png`              | Workspace Folder QuickPick with "Default folder"                  | Default folder option present                               | Confirms folder selection step reached                                  |
| 12  | `12-m365-signin-complete.png`          | ATK sidebar showing signed-in M365 account                        | Account name visible in ATK accounts section                | Confirms M365 login succeeded before provision                          |
| 13  | `13-provision-starting.png`            | VS Code with provision command executing                          | No error banner; progress indicator visible                 | Confirms provision was triggered via Command Palette                    |
| 14  | `14-dev-env-selected.png`              | "Select an environment" QuickPick with dev highlighted            | `dev` option focused                                        | Confirms correct environment selected                                   |
| 15  | `15-oauth-client-id-input.png`         | OAuth registration client ID InputBox                             | Prompt "OAuth registration client ID" visible               | Confirms OAuth credential collection appears mid-provision              |
| 16  | `16-oauth-client-secret-input.png`     | OAuth registration client secret InputBox                         | Prompt "OAuth registration client secret" visible           | Confirms secret step follows client ID step                             |
| 17  | `17-oauth-confirm-dialog.png`          | OAuth registration confirmation dialog with Confirm button        | Confirm button present                                      | Proves review step is shown before finalising OAuth registration        |
| 18  | `18-provision-succeeded.png`           | VS Code notifications panel                                       | "provision stage executed successfully" notification visible | Primary provision pass criterion                                        |
| 19  | `19-copilot-opened.png`                | M365 Copilot interface in Chrome                                  | Copilot chat UI visible and responsive                      | Confirms remote debug launch succeeded                                  |
| 20  | `20-repair-query-sent.png`             | Copilot chat with repair query submitted                          | Query text visible in chat                                  | Records the query that triggers the OAuth action                        |
| 21  | `21-allow-prompt.png`                  | Copilot response with Allow button                                | Allow button visible in chat                                | Confirms agent triggered permission prompt for the OAuth action         |
| 22  | `22-sign-in-button.png`                | Copilot chat showing "Sign in to Repair Service" button           | Button labeled "Sign in to Repair Service" present          | Primary end-to-end assertion: OAuth action wired correctly in Copilot   |

---

## Notes

- This is a full end-to-end test: it creates a new DA project, provisions to Azure, and validates in M365 Copilot remote. It is **not** a local-debug test.
- The test targets **M365 Copilot** (not Microsoft Teams). The debug launch option is "Preview in Copilot (Chrome)".
- The "Select Operation(s)" dialog has a 60-second precondition wait because the ATK extension fetches the spec from the remote URL at this step.
- The OK → Clear All sequence is the recorded interaction order from the original plan: OK first confirms the pre-populated selection, then Clear All resets it. The test accepts any state after Clear All and does not re-select individual operations before proceeding.
- OAuth credentials appear as VS Code InputBoxes during the provision run, not during wizard creation. Allow up to 30 s after the `dev` environment selection for the client ID prompt to appear (this is the `delay: 30` tag on the step).
- The "Sign in to Repair Service" assertion uses `delay: 30` and `step_retry_timeout: 60` — the test runner waits 30 s before checking and retries for up to 60 s. The button appearing is the sole end-to-end pass criterion for the OAuth action wiring.
- `fake_client_id` / `fake_clientsecret` are placeholders sufficient for Azure OAuth registration creation. They do not need to be real credentials for the sign-in button to appear.
