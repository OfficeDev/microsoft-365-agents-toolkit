# Test Plan: Create Declarative Agent With MCP Server

## Metadata

- **feature-slug**: `create-da-with-mcp-server`
- **owner**: atk-qa
- **created**: 2026-05-28
- **updated**: 2026-05-29
- **triggers**: issue-label `atk-copilot-test`, manual
- **scenario-id**: SCN-DA-CREATE-WITH-MCP-SERVER
- **related-scenario**: docs/01-product/scenarios/da/create-da-with-mcp-server.md

## Scope

**Covers:**
- TC-001: VS Code happy path — create a DA project with a remote MCP server action (no `odr.exe`)
- TC-002: VS Code happy path — create a DA project with a remote MCP server action (`odr.exe` present, user chooses Remote)
- TC-003: CLI non-interactive happy path — create DA project with `--api-plugin-type mcp` and `--mcp-da-server-url`
- TC-004: CLI non-interactive error path — missing `--mcp-da-server-url` when `--api-plugin-type mcp`
- TC-005: CLI non-interactive error path — missing `--mcp-da-auth-type` when tools file provided and auth is required
- TC-006: VS Code cancellation path — cancel before project generation leaves no partial project

**Does NOT cover:**
- Local MCP server path via `odr.exe` (covered separately when odr integration is stable)
- `SCN-DA-ADD-MCP-ACTION-TO-DA` (post-create CodeLens flow)
- `SCN-DA-FETCH-MCP-TOOLS` (tool discovery follow-up)
- Visual Studio and chat surfaces
- CLI interactive create flow (separate plan)
- OAuth / Entra SSO authentication type selection in detail (covered by auth-specific plans)

---

## Wizard flow source

The VS Code wizard question tree is loaded from `packages/fx-core/templates/ui/wizardNode.json`
and built by `scaffoldQuestionForVSCode()` in `createRootNode.ts`. The relevant steps for the
DA + MCP remote path are:

1. **`project-type`** QuickPick (title: "New Project") — user selects **"Declarative Agent"**
   (id `copilot-agent-type`, group "Agents for Microsoft 365 Copilot").
2. **`capabilities`** QuickPick (title: "App Features Using Agents") — **auto-skipped** because
   `skipSingleOption: true` and the only option is "Declarative Agent".
3. **`with-plugin`** QuickPick (title: "Create Declarative Agent") — user selects **"Add an Action"**
   (id `yes`).
4. **`action-type`** QuickPick (title: "Create an Action") — user selects **"Start with a MCP server"**
   (id `mcp`).
5. **`mcp-server-type`** QuickPick (title: "MCP Server Type") — `skipSingleOption: true`; dynamic
   options from `ODRProvider.listServers()`. When `odr.exe` is absent only "Remote MCP Server"
   is returned and the step is **auto-skipped**. When `odr.exe` is present both "Local MCP Server"
   and "Remote MCP Server" appear and the user must choose.
6. **`mcp-da-server-url`** InputBox (title: "MCP Server URL") — user types the remote MCP
   server URL.
7. **Workspace folder** picker — user picks project location.
8. **`app-name`** InputBox — user types the application name.

Evidence: `packages/fx-core/src/question/scaffold/vsc/daProjectTypeNode.ts`,
`packages/fx-core/src/question/scaffold/vsc/teamsProjectTypeNode.ts` (MCPServerTypeNode),
`packages/fx-core/templates/ui/wizardNode.json`,
`packages/fx-core/resource/package.nls.json`.

---

## Test Cases

### TC-001 – VS Code: Create DA with remote MCP server (no odr.exe, happy path)

**Preconditions:**
- VS Code is open with no project loaded
- ATK extension v6.8.0+ is installed and activated
- `odr.exe` is **not** present on the machine (MCP Server Type QuickPick is auto-skipped)
- Network is not required for project scaffolding (URL value is stored in `.vscode/mcp.json` as-is)

**Wizard flow (from `wizardNode.json` + `daProjectTypeNode.ts`):**

| Step | QuickPick / InputBox                        | Value to select / type                         |
|------|---------------------------------------------|------------------------------------------------|
| 1    | New Project                                 | Declarative Agent                              |
| 2    | (App Features Using Agents — auto-skipped)  | —                                              |
| 3    | Create Declarative Agent                    | Add an Action                                  |
| 4    | Create an Action                            | Start with a MCP server                        |
| 5    | (MCP Server Type — auto-skipped, Remote)    | —                                              |
| 6    | MCP Server URL (InputBox)                   | `https://mcptest.example.com/sse`              |
| 7    | Workspace folder                            | Default folder                                 |
| 8    | Application Name (InputBox)                 | `test-da-mcp-001` (type + Enter)               |

> **Note:** Step 2 is auto-skipped because `skipSingleOption: true` and only "Declarative Agent"
> is in the capabilities list. Step 5 is auto-skipped when `odr.exe` is absent because
> `ODRProvider.listServers()` returns only the "Remote MCP Server" option and the QuickPick
> has `skipSingleOption: true`.

**Steps:**
1. Fire `fx-extension.create` command via the Command Palette or programmatically. Observe the "New Project" QuickPick appears.
2. Take screenshot of the "New Project" QuickPick panel (step 1 of wizard).
3. Click "Declarative Agent" in the "Agents for Microsoft 365 Copilot" group. Observe the "Create Declarative Agent" QuickPick appears (the "App Features Using Agents" step is auto-skipped).
4. Take screenshot showing the "Create Declarative Agent" QuickPick with "Add an Action" option visible.
5. Click "Add an Action". Observe the "Create an Action" QuickPick appears.
6. Take screenshot showing the "Create an Action" QuickPick with "Start with a MCP server" option visible.
7. Click "Start with a MCP server". Observe the "MCP Server URL" InputBox appears directly (MCP Server Type step is auto-skipped because `odr.exe` is absent).
8. Take screenshot showing the "MCP Server URL" InputBox (empty, before typing).
9. Type `https://mcptest.example.com/sse` and press Enter. Observe the workspace folder picker appears.
10. Take screenshot showing the workspace folder picker.
11. Click "Default folder". Observe the "Application Name" InputBox appears.
12. Take screenshot showing the "Application Name" InputBox (empty, before typing).
13. Type `test-da-mcp-001` and press Enter. Observe the scaffold process begins.
14. Wait up to 120 s for project generation and VS Code to open the new project folder.
15. Take final state screenshot showing the VS Code file explorer tree.
16. Assert required project files exist under `~/AgentsToolkitProjects/test-da-mcp-001/`.

**Expected result:**
- Wizard completes without error.
- The following files exist at the project root:
  - `m365agents.yml`
  - `appPackage/manifest.json`
  - `appPackage/declarativeAgent.json`
  - `.vscode/mcp.json`
- `.vscode/mcp.json` contains the MCP server URL `https://mcptest.example.com/sse`.
- VS Code opens the new project folder automatically.

**Pass criteria:**
- All 4 asserted files are present.
- `.vscode/mcp.json` file exists and contains `mcptest.example.com/sse`.
- No error notification is shown in VS Code.

**Test script:**
`packages/tests/src/create-da-with-mcp-server-vscode-happy.test.ts`

**Screenshots produced by test:**

| ID  | Filename                             | What is visible                                               | Pass condition                                                | Why                                                                       |
|-----|--------------------------------------|---------------------------------------------------------------|---------------------------------------------------------------|---------------------------------------------------------------------------|
| 01  | `01-wizard-open.png`                 | "New Project" QuickPick open after command fires              | "Declarative Agent" option visible in group                   | Proves create wizard launched and DA option is present at top level        |
| 02  | `02-create-da-quickpick.png`         | "Create Declarative Agent" QuickPick                          | "Add an Action" option is visible                             | Confirms DA path reached; capabilities step was auto-skipped              |
| 03  | `03-create-an-action.png`            | "Create an Action" QuickPick                                  | "Start with a MCP server" option is visible                   | Confirms action source choices appear after "Add an Action"               |
| 04  | `04-mcp-url-input.png`               | "MCP Server URL" InputBox (empty, before typing)              | InputBox title contains "MCP Server URL"                      | Confirms URL input appears directly (MCP Server Type auto-skipped)        |
| 05  | `05-workspace-folder.png`            | Workspace folder picker with "Default folder" option          | "Default folder" option visible                               | Confirms folder step reached after URL entry                              |
| 06  | `06-app-name-input.png`              | "Application Name" InputBox (empty, before typing)            | InputBox is open and empty                                    | Captures pre-entry state before typing app name                           |
| 07  | `07-project-created.png`             | VS Code immediately after scaffold completes                  | No error notification visible                                 | Proves project generation succeeded without error                         |
| 08  | `08-final-state.png`                 | File explorer tree showing generated project files            | `.vscode/mcp.json` and `appPackage/` folder visible           | Proves DA project with MCP config was generated correctly                 |

---

### TC-002 – VS Code: Create DA with remote MCP server (odr.exe present, user chooses Remote)

**Preconditions:**
- VS Code is open with no project loaded
- ATK extension v6.8.0+ is installed and activated
- `odr.exe` **is present** on the machine (MCP Server Type QuickPick is shown with both Local and Remote options)

**Wizard flow:**

| Step | QuickPick / InputBox                        | Value to select / type                         |
|------|---------------------------------------------|------------------------------------------------|
| 1    | New Project                                 | Declarative Agent                              |
| 2    | (App Features Using Agents — auto-skipped)  | —                                              |
| 3    | Create Declarative Agent                    | Add an Action                                  |
| 4    | Create an Action                            | Start with a MCP server                        |
| 5    | MCP Server Type                             | Remote MCP Server                              |
| 6    | MCP Server URL (InputBox)                   | `https://mcptest.example.com/sse`              |
| 7    | Workspace folder                            | Default folder                                 |
| 8    | Application Name (InputBox)                 | `test-da-mcp-odr-001` (type + Enter)           |

**Steps:**
1. Fire `fx-extension.create` command. Observe the "New Project" QuickPick appears.
2. Take screenshot of the "New Project" QuickPick panel.
3. Click "Declarative Agent". Observe "Create Declarative Agent" QuickPick appears.
4. Click "Add an Action". Observe "Create an Action" QuickPick appears.
5. Click "Start with a MCP server". Observe the "MCP Server Type" QuickPick appears (because `odr.exe` is present, both Local and Remote options are shown).
6. Take screenshot showing the "MCP Server Type" QuickPick with "Remote MCP Server" and "Local MCP Server" options.
7. Click "Remote MCP Server". Observe the "MCP Server URL" InputBox appears.
8. Take screenshot showing the "MCP Server URL" InputBox (empty).
9. Type `https://mcptest.example.com/sse` and press Enter. Observe workspace folder picker appears.
10. Click "Default folder". Observe the "Application Name" InputBox.
11. Type `test-da-mcp-odr-001` and press Enter. Observe scaffold begins.
12. Wait up to 120 s for project generation.
13. Take final state screenshot showing the VS Code file explorer tree.
14. Assert required project files exist under `~/AgentsToolkitProjects/test-da-mcp-odr-001/`.

**Expected result:**
- Wizard completes without error.
- Files present: `m365agents.yml`, `appPackage/manifest.json`, `appPackage/declarativeAgent.json`, `.vscode/mcp.json`.
- `.vscode/mcp.json` contains `https://mcptest.example.com/sse`.

**Pass criteria:**
- All 4 asserted files are present.
- `.vscode/mcp.json` exists and contains the MCP server URL.
- No error notification shown.

**Test script:**
`packages/tests/src/create-da-with-mcp-server-vscode-odr.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                               | Pass condition                                               | Why                                                                       |
|-----|---------------------------------------|---------------------------------------------------------------|--------------------------------------------------------------|---------------------------------------------------------------------------|
| 01  | `01-wizard-open.png`                  | "New Project" QuickPick open                                  | "Declarative Agent" option visible in group                  | Proves create wizard launched and DA option is available at top level     |
| 02  | `02-mcp-server-type.png`              | "MCP Server Type" QuickPick with Local and Remote options     | Both "Remote MCP Server" and "Local MCP Server" visible      | Proves odr.exe presence triggers the server-type choice step              |
| 03  | `03-mcp-url-input.png`                | "MCP Server URL" InputBox after choosing Remote               | InputBox open with title "MCP Server URL"                    | Confirms Remote choice leads to URL input                                 |
| 04  | `04-final-state.png`                  | File explorer tree with `.vscode/mcp.json` visible            | `.vscode/mcp.json` and `appPackage/` visible                 | Proves DA + MCP project generated with Remote server type                 |

---

### TC-003 – CLI non-interactive: Create DA with remote MCP server (happy path)

**Preconditions:**
- ATK CLI is installed and available on `PATH` as `atk`
- A valid target folder exists or can be created
- No project exists at `~/atk-test-out/test-da-mcp-cli-001`

**Steps:**
1. Open an integrated terminal in VS Code (or use a system terminal).
2. Run the following command and observe the terminal output:
   ```
   atk new -c declarative-agent --with-plugin yes --api-plugin-type mcp --mcp-server-type remote --mcp-da-server-url https://mcptest.example.com/sse -n test-da-mcp-cli-001 -f ~/atk-test-out --interactive false
   ```
3. Observe the command exits with status `0`.
4. Observe terminal output shows project creation succeeded (no error lines).
5. Take screenshot of the terminal output.
6. Verify the following files exist under `~/atk-test-out/test-da-mcp-cli-001/`:
   - `m365agents.yml`
   - `appPackage/manifest.json`
   - `appPackage/declarativeAgent.json`
   - `.vscode/mcp.json`
7. Take screenshot of the file tree showing the generated project.

**Expected result:**
- CLI exits with status `0`.
- All 4 files are present.
- `.vscode/mcp.json` contains the URL `https://mcptest.example.com/sse`.
- Terminal output contains no error or missing-option messages.
- If MCP tools were auto-fetched, `ai-plugin.json` is also present; if not, a hint warning containing `atk add action` is printed to stdout.

**Pass criteria:**
- CLI exit code is `0`.
- `m365agents.yml`, `appPackage/manifest.json`, `appPackage/declarativeAgent.json`, `.vscode/mcp.json` are present.
- `.vscode/mcp.json` contains `mcptest.example.com/sse`.

**Test script:**
`packages/tests/src/create-da-with-mcp-server-cli-noninteractive.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                             | Pass condition                                              | Why                                                                    |
|-----|---------------------------------------|-------------------------------------------------------------|-------------------------------------------------------------|------------------------------------------------------------------------|
| 01  | `01-cli-output.png`                   | Terminal showing CLI command and its output                 | No error lines; exit code 0                                 | Proves CLI completed without error                                     |
| 02  | `02-project-files.png`                | File tree with generated project files                      | `.vscode/mcp.json` and `appPackage/` folder visible         | Proves project files were scaffolded correctly                         |

---

### TC-004 – CLI non-interactive error: missing `--mcp-da-server-url`

**Preconditions:**
- ATK CLI is installed and available as `atk`

**Steps:**
1. Open a terminal.
2. Run the following command (intentionally omitting `--mcp-da-server-url`):
   ```
   atk new -c declarative-agent --with-plugin yes --api-plugin-type mcp --mcp-server-type remote -n test-da-mcp-err-001 -f ~/atk-test-out --interactive false
   ```
3. Observe the command exits with a non-zero status code.
4. Take screenshot of the terminal showing the error message.
5. Verify that the error output references the missing `--mcp-da-server-url` option.
   The error message text is defined in `core.MCPForDA.missingServerUrl`:
   "The --mcp-da-server-url option is required when --api-plugin-type is mcp."
6. Verify that no project folder was created at `~/atk-test-out/test-da-mcp-err-001`.

**Expected result:**
- CLI exits with a non-zero status code.
- Error output references `mcp-da-server-url` (from `core.MCPForDA.missingServerUrl`).
- No partial project folder is created.

**Pass criteria:**
- CLI exit code is non-zero (e.g., `1`).
- stderr or stdout contains `mcp-da-server-url`.
- `~/atk-test-out/test-da-mcp-err-001` does not exist.

**Test script:**
`packages/tests/src/create-da-with-mcp-server-cli-missing-url.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                               | Pass condition                                                 | Why                                                                         |
|-----|---------------------------------------|---------------------------------------------------------------|----------------------------------------------------------------|-----------------------------------------------------------------------------|
| 01  | `01-cli-error-output.png`             | Terminal showing error output for missing MCP URL             | Error message mentioning `mcp-da-server-url` is visible        | Proves CLI validates required MCP URL flag and surfaces a useful error      |

---

### TC-005 – CLI non-interactive error: missing `--mcp-da-auth-type` when tools file is provided

**Preconditions:**
- ATK CLI is installed and available as `atk`
- A valid MCP tools JSON file exists at `~/atk-test-out/tools.json`

**Steps:**
1. Open a terminal.
2. Create a minimal tools JSON file:
   ```
   echo '{"tools":[{"name":"search","description":"Search the web"}]}' > ~/atk-test-out/tools.json
   ```
3. Run the following command (tools file provided but `--mcp-da-auth-type` omitted, targeting a server
   that is expected to require auth):
   ```
   atk new -c declarative-agent --with-plugin yes --api-plugin-type mcp --mcp-server-type remote --mcp-da-server-url https://mcpauth.example.com/sse --mcp-tools-file-path ~/atk-test-out/tools.json -n test-da-mcp-auth-001 -f ~/atk-test-out --interactive false
   ```
4. Observe the terminal output.
5. Take screenshot of the terminal output.
6. If the CLI exits non-zero, verify the error output references `mcp-da-auth-type`.
   The error message text is defined in `core.MCPForDA.missingAuthType`:
   "The --mcp-da-auth-type option is required when the MCP server requires authentication and tools are provided."

**Expected result:**
- When the CLI detects the server requires authentication and `--mcp-da-auth-type` is not supplied,
  it exits with a non-zero status and an error message referencing `mcp-da-auth-type`
  (from `core.MCPForDA.missingAuthType` in `package.nls.json`).
- Alternatively, if auth detection is skipped because the server is unreachable, the CLI generates
  the project with a warning hint containing `atk add action`.

**Pass criteria:**
- If auth is probed and required: CLI exit code is non-zero; stderr/stdout contains `mcp-da-auth-type`.
- If auth detection is skipped: CLI exit code is `0` and a warning is printed containing `atk add action`.

**Test script:**
`packages/tests/src/create-da-with-mcp-server-cli-missing-auth.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                               | Pass condition                                                  | Why                                                                           |
|-----|---------------------------------------|---------------------------------------------------------------|-----------------------------------------------------------------|-------------------------------------------------------------------------------|
| 01  | `01-cli-auth-error-output.png`        | Terminal showing CLI output when auth type is missing         | Error or warning referencing `mcp-da-auth-type` or `atk add action` | Proves CLI handles missing auth type gracefully                          |

---

### TC-006 – VS Code: Cancel before project generation leaves no partial project

**Preconditions:**
- VS Code is open with no project loaded
- ATK extension v6.8.0+ is installed and activated
- No folder named `test-da-mcp-cancel-001` exists under `~/AgentsToolkitProjects/`

**Steps:**
1. Fire `fx-extension.create` command. Observe the "New Project" QuickPick appears.
2. Click "Declarative Agent". Observe "Create Declarative Agent" QuickPick appears.
3. Click "Add an Action". Observe "Create an Action" QuickPick appears.
4. Click "Start with a MCP server". Observe the "MCP Server URL" InputBox appears.
5. Take screenshot showing the "MCP Server URL" InputBox open, awaiting input.
6. Press `Escape` to cancel the wizard. Observe the InputBox closes and no project generation starts.
7. Take screenshot of VS Code after cancellation (no new folder or window opened).
8. Verify that `~/AgentsToolkitProjects/test-da-mcp-cancel-001` does not exist.

**Expected result:**
- Pressing Escape during the URL input step cancels the wizard.
- No partial project folder is created.
- VS Code returns to its pre-creation state (no new window or folder opened).

**Pass criteria:**
- `~/AgentsToolkitProjects/test-da-mcp-cancel-001` does not exist after cancellation.
- No error notification is shown in VS Code after cancellation.

**Test script:**
`packages/tests/src/create-da-with-mcp-server-vscode-cancel.test.ts`

**Screenshots produced by test:**

| ID  | Filename                              | What is visible                                               | Pass condition                                               | Why                                                                    |
|-----|---------------------------------------|---------------------------------------------------------------|--------------------------------------------------------------|------------------------------------------------------------------------|
| 01  | `01-mcp-url-input-before-cancel.png`  | "MCP Server URL" InputBox open, awaiting input                | InputBox is open with title "MCP Server URL"                 | Captures state just before cancellation                                |
| 02  | `02-after-cancel.png`                 | VS Code after Escape is pressed, no QuickPick or InputBox     | No wizard panel open; VS Code in idle state                  | Proves cancellation cleanly dismisses the wizard                       |
