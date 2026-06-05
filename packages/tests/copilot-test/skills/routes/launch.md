# Route: Launch VS Code

This step runs **before every other route**. It produces a running VS Code instance
with the ATK extension loaded and the Playwright CDP connection ready.

Two modes are supported: **local dev** (for iterative development) and
**VSIX/remote** (for CI or stable-build verification).

---

## Mode 1: Local dev (`--extensionDevelopmentPath`)

### Prerequisites

**1. Build the extension first**

```bash
cd packages/vscode-extension
node esbuild.mjs --skip-type-check
```

Re-run after any source change (`m365Login.ts`, `codeFlowLogin.ts`, etc.).
The compiled bundle lands in `packages/vscode-extension/dist/extension.js`.

**2. Prepare a workspace with `m365agents.yml`**

ATK's activation event is `workspaceContains:m365agents*.yml`. If the opened
workspace folder does not contain that file at startup, the extension never
activates and all test steps will fail silently.

```typescript
import * as fs   from "fs";
import * as path from "path";
import * as os   from "os";

function prepareWorkspace(dir: string): void {
  fs.mkdirSync(dir, { recursive: true });
  const yml = path.join(dir, "m365agents.yml");
  if (!fs.existsSync(yml)) {
    // Minimal stub — enough to satisfy workspaceContains
    fs.writeFileSync(yml, "version: 1.0\n", "utf8");
  }
}

const workspaceDir = path.join(os.tmpdir(), "atk-test-workspace");
prepareWorkspace(workspaceDir);
```

### Launch code

```typescript
import { downloadAndUnzipVSCode } from "@vscode/test-electron";
import { _electron as electron }  from "playwright";

const VS_CODE_VERSION = "1.104.2"; // must satisfy engines.vscode in the extension's package.json
const EXT_DEV_PATH    = path.resolve(__dirname, "../../../../vscode-extension");

async function launchLocalDev(workspaceDir: string) {
  const vscodePath = await downloadAndUnzipVSCode(VS_CODE_VERSION);

  // Isolated dirs — no state leaks between runs
  const userDataDir  = path.join(os.tmpdir(), `atk-ud-${Date.now()}`);
  const extensionDir = path.join(os.tmpdir(), `atk-ed-${Date.now()}`);
  fs.mkdirSync(userDataDir,  { recursive: true });
  fs.mkdirSync(extensionDir, { recursive: true });

  // Suppress first-run prompts
  const settingsDir = path.join(userDataDir, "User");
  fs.mkdirSync(settingsDir, { recursive: true });
  fs.writeFileSync(path.join(settingsDir, "settings.json"), JSON.stringify({
    "workbench.colorTheme":    "Default Dark Modern",
    "workbench.startupEditor": "none",
    "telemetry.telemetryLevel": "off",
    "extensions.autoUpdate":   false,
  }, null, 2), "utf8");

  const app = await electron.launch({
    executablePath: vscodePath,
    args: [
      `--extensionDevelopmentPath=${EXT_DEV_PATH}`,
      `--user-data-dir=${userDataDir}`,
      `--extensions-dir=${extensionDir}`,
      "--disable-workspace-trust",
      "--skip-release-notes",
      workspaceDir,           // MUST be last — VS Code opens this as the workspace
    ],
    env: {
      ...process.env,
      TEAMSFX_AUTO_CONFIRM_LOGIN: "true",  // bypasses confirmation modal (see routes/auth.md)
      TEAMSFX_BROKER_AUTH:        "false", // forces MSAL browser flow, not WAM
    },
  });

  const page = await app.firstWindow();
  await page.waitForLoadState("domcontentloaded");
  return { app, page };
}
```

---

## Mode 2: VSIX / remote CI (`--install-extension` + `--extensions-dir`)

Used in CI pipelines or when testing a packaged release build.
The VSIX is installed into an isolated extensions directory before VS Code starts.

### Build the VSIX (if not already produced by CI)

```bash
cd packages/vscode-extension
npx vsce package --out dist/ms-teams-vscode-extension.vsix
```

### Install and launch

```typescript
import { downloadAndUnzipVSCode, resolveCliPathFromVSCodeExecutablePath } from "@vscode/test-electron";
import { _electron as electron } from "playwright";
import * as cp from "child_process";

async function launchVsix(vsixPath: string, workspaceDir: string) {
  const vscodePath  = await downloadAndUnzipVSCode(VS_CODE_VERSION);
  const cliPath     = resolveCliPathFromVSCodeExecutablePath(vscodePath);

  const userDataDir  = path.join(os.tmpdir(), `atk-ud-${Date.now()}`);
  const extensionDir = path.join(os.tmpdir(), `atk-ed-${Date.now()}`);
  fs.mkdirSync(userDataDir,  { recursive: true });
  fs.mkdirSync(extensionDir, { recursive: true });

  // Install VSIX into the isolated extensions dir
  cp.execFileSync(cliPath, [
    "--install-extension", vsixPath,
    "--extensions-dir",    extensionDir,
    "--force",
  ], { stdio: "inherit" });

  // Suppress first-run prompts (same as local dev)
  const settingsDir = path.join(userDataDir, "User");
  fs.mkdirSync(settingsDir, { recursive: true });
  fs.writeFileSync(path.join(settingsDir, "settings.json"), JSON.stringify({
    "workbench.colorTheme":    "Default Dark Modern",
    "workbench.startupEditor": "none",
    "telemetry.telemetryLevel": "off",
    "extensions.autoUpdate":   false,
  }, null, 2), "utf8");

  // Launch — NO --extensionDevelopmentPath; extension comes from --extensions-dir
  const app = await electron.launch({
    executablePath: vscodePath,
    args: [
      `--extensions-dir=${extensionDir}`,
      `--user-data-dir=${userDataDir}`,
      "--disable-workspace-trust",
      "--skip-release-notes",
      workspaceDir,
    ],
    env: {
      ...process.env,
      TEAMSFX_AUTO_CONFIRM_LOGIN: "true",
      TEAMSFX_BROKER_AUTH:        "false",
    },
  });

  const page = await app.firstWindow();
  await page.waitForLoadState("domcontentloaded");
  return { app, page };
}
```

| | Local dev | VSIX / remote |
|---|---|---|
| Extension loaded via | `--extensionDevelopmentPath` | `--extensions-dir` (VSIX pre-installed) |
| Build step | `node esbuild.mjs` | `npx vsce package` |
| Reflects source changes | Immediately | Requires re-package |
| Typical use | Local iteration | CI, release validation |

---

## Windows: avoiding `STATUS_CONTROL_C_EXIT`

On Windows, if the Node/PowerShell process that owns the VS Code process group
exits (e.g., a CI step finishes), VS Code receives `STATUS_CONTROL_C_EXIT` and
crashes — corrupting any in-progress test.

When VS Code must outlive its parent process, use a WScript batch wrapper to
create a fully independent process:

```typescript
import * as cp from "child_process";

function launchDetached(vscodePath: string, args: string[]): void {
  const bat = path.join(os.tmpdir(), `vsc-launch-${Date.now()}.bat`);
  const quoted = args.map(a => `"${a.replace(/"/g, '\\"')}"`).join(" ");
  fs.writeFileSync(bat, `@echo off\nstart "" "${vscodePath}" ${quoted}\n`, "utf8");
  cp.spawn("cmd.exe", ["/c", bat], { detached: true, stdio: "ignore" }).unref();
}
```

> **When to use**: only when VS Code must stay alive after the launcher script exits.
> For the `electron.launch()` paths above, Playwright owns the lifecycle — this
> wrapper is not needed.

---

## Post-launch: verify ATK is active (Step 0)

Always run this immediately after launch. See `test-code-helpers.md` for the
full `waitForATKActivation()` implementation.

```typescript
const active = await waitForATKActivation(60_000);
await takeElementScreenshot("00-atk-active",
  "[id='workbench.view.extension.teamsfx'] .pane-body");
assert.ok(active, "ATK extension must be active before any test step runs");
```

**If ATK is not active after 60 s, check:**
1. `m365agents.yml` exists in the workspace folder opened at launch.
2. Local dev: `--extensionDevelopmentPath` points to the extension root (not `dist/`).
3. Local dev: `packages/vscode-extension/dist/extension.js` exists — rebuild if missing.
4. VSIX: `code --install-extension` completed without error before `electron.launch()`.
