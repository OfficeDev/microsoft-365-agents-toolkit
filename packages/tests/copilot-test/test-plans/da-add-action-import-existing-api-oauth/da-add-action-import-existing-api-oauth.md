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
- Adding an action to an existing Declarative Agent project via "Start with an OpenAPI Description Document"
- Loading an OpenAPI spec from a remote URL that requires OAuth authentication
- Selecting operations from the loaded spec (clear-all then confirm flow)
- Entering OAuth registration credentials (client ID and client secret) during provision
- Confirming the OAuth registration dialog
- Verifying the provisioned agent presents the OAuth sign-in button in the chat

**Does NOT cover:**
- Creating the initial DA project (assumed pre-existing)
- No-auth OpenAPI import (covered by `DA_Add_Action_Import_Existing_API_Basic_No_Auth`)
- CLI-based add-action flow
- Azure deployment and remote debug

---

## Test Cases

### TC-001 – Add OAuth API action to DA, verify sign-in button appears

**Preconditions:**
- VSCode is open with an existing Declarative Agent project loaded
- ATK extension v6.8.0+ is installed and activated
- User is signed in to their M365 account in ATK
- Dev Tunnel is configured (Public access)
- No action has been added to the DA project yet

**Wizard flow (verified on ATK v6.8.0):**

| Step | QuickPick / InputBox                                    | Value to select / type |
|------|---------------------------------------------------------|------------------------|
| 1    | Trigger "Add action" on the DA project (CodeLens or ATK tree view command) | — |
| 2    | Action source QuickPick                                 | Start with an OpenAPI Description Document |
| 3    | Enter OpenAPI Description Document URL                  | `https://raw.githubusercontent.com/SLdragon/example-openapi-spec/refs/heads/main/real-oauth.yaml` (type + Enter) |
| 4    | Select Operation(s) Copilot Can Interact with – click **OK** | OK (accept loaded operations) |
| 5    | Select Operation(s) – click **Clear All**               | Clear All (deselect all, then let wizard proceed) |
| 6    | Workspace Folder                                        | Default folder |
| 7    | Application Name                                        | `test-da-oauth-action-001` (type + Enter) |
| 8    | OAuth registration client ID input                      | `fake_client_id` (type + Enter) |
| 9    | OAuth registration client secret input                  | `fake_clientsecret` (type + Enter) |
| 10   | OAuth registration confirmation dialog                  | Confirm |

**Steps:**
1. Open an existing DA project in VS Code; wait for ATK extension to activate
2. Trigger the "Add action" command (CodeLens on `declarativeAgent.json` or ATK sidebar > "Add action"); take screenshot
3. When action source QuickPick appears, select "Start with an OpenAPI Description Document"; take screenshot
4. When "Enter OpenAPI Description Document URL" input appears, type the OAuth spec URL and press Enter; take screenshot of the filled input
5. Wait up to 60 s for the "Select Operation(s) Copilot Can Interact with" dialog to load the spec; take screenshot
6. Click **OK** in the operations dialog; take screenshot
7. Click **Clear All** to deselect all operations; take screenshot showing cleared state
8. Click **Default folder** in the Workspace Folder dialog; take screenshot
9. Type app name `test-da-oauth-action-001` and press Enter
10. Wait for project/action generation to begin; then when the OAuth client ID input appears (may take up to 30 s during provision), take screenshot of the input box
11. Type `fake_client_id` and press Enter; take screenshot
12. When the OAuth client secret input appears, type `fake_clientsecret` and press Enter; take screenshot
13. When the OAuth registration confirmation dialog appears, click **Confirm**; take screenshot
14. Wait up to 30 s for provisioning and agent registration to complete
15. Open the agent in Teams (via ATK "Preview in" > Teams)
16. Send a message to the agent and observe the response; take final screenshot

**Expected result:**
- Action files are generated in `appPackage/` (e.g., `ai-plugin.json` referencing the OAuth spec)
- Provision completes without error
- Teams chat with the agent shows a **"Sign in to Repair Service"** button, confirming the OAuth auth flow is wired correctly

**Pass criteria:**
- `appPackage/ai-plugin.json` exists and references the `real-oauth.yaml` spec URL
- Provision exit code is 0 (no error notification in VS Code)
- A button labeled "Sign in to Repair Service" is visible in the Teams agent chat

**Test script:**
`packages/tests/src/da-add-action-import-existing-api-oauth.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                              | Pass condition                                                  | Why                                                                  |
|-----|---------------------------------------|--------------------------------------------------------------|-----------------------------------------------------------------|----------------------------------------------------------------------|
| 01  | `01-add-action-triggered.png`         | Action source QuickPick open after "Add action" command      | QuickPick present with "Start with an OpenAPI Description Document" option | Confirms add-action wizard launched successfully                      |
| 02  | `02-openapi-source-selected.png`      | OpenAPI Description Document option highlighted              | Option visible and focused                                      | Confirms correct action source branch taken                          |
| 03  | `03-url-input-filled.png`             | URL InputBox with OAuth spec URL typed                       | Full URL visible in input field                                 | Records exact spec URL used for traceability                         |
| 04  | `04-operations-dialog.png`            | "Select Operation(s)" dialog with operations loaded          | Operations list visible; OK button present                      | Proves spec was fetched and parsed successfully                       |
| 05  | `05-ok-clicked.png`                   | Operations dialog after clicking OK                          | Dialog still visible or transitioning                           | Confirms OK interaction recorded                                      |
| 06  | `06-clear-all-clicked.png`            | Operations dialog with all items deselected                  | No checked operations visible                                   | Confirms Clear All interaction worked                                 |
| 07  | `07-workspace-folder.png`             | Workspace Folder dialog with "Default folder" option         | Default folder option present                                   | Confirms flow progressed to folder selection                          |
| 08  | `08-oauth-client-id-input.png`        | OAuth registration client ID InputBox (before typing)        | InputBox prompt "OAuth registration client ID" visible          | Confirms OAuth credential collection step triggered during provision  |
| 09  | `09-oauth-client-id-entered.png`      | OAuth client ID InputBox with `fake_client_id` typed         | Text visible in input                                           | Records client ID entry step                                         |
| 10  | `10-oauth-client-secret-input.png`    | OAuth registration client secret InputBox (before typing)    | InputBox prompt "OAuth registration client secret" visible      | Confirms secret collection follows client ID step                     |
| 11  | `11-oauth-client-secret-entered.png`  | OAuth client secret InputBox with `fake_clientsecret` typed  | Text visible in input                                           | Records client secret entry step                                     |
| 12  | `12-oauth-confirm-dialog.png`         | OAuth registration confirmation dialog with Confirm button   | Confirm button present with client ID/secret summary            | Proves OAuth registration review step is shown before finalising      |
| 13  | `13-sign-in-button.png`               | Teams chat with agent showing "Sign in to Repair Service"    | Button labeled "Sign in to Repair Service" visible in chat      | Primary assertion: proves OAuth action was provisioned end-to-end     |

---

## Notes

- The "Select Operation(s)" dialog appears with a 60-second wait timeout because the ATK extension fetches the spec from the remote URL at this step; network latency can vary.
- Steps 4–7 (OK then Clear All) reflect the current recorded interaction sequence: the extension auto-selects all operations by default, OK confirms that selection, and Clear All is used to reset the state before a specific subset is confirmed. Tests that need a specific operation selected should insert selection steps between Clear All and proceeding.
- The OAuth credential inputs appear mid-provision (not during wizard creation). Expect a delay of up to 30 s between entering the app name and seeing the client ID prompt.
- `fake_client_id` / `fake_clientsecret` are placeholder values sufficient for schema validation in test environments; a real OAuth app registration is needed for end-to-end sign-in to succeed.
- The final "Sign in to Repair Service" button assertion only requires the button to be visible — it does not require the sign-in to complete. The referenced OpenAPI spec models a "Repair Service" API, and this button is the ATK-generated OAuth entry point for that service.
