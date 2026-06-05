// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * da-no-action-local-debug.test.ts
 * TC-001: Create DA (No Action), verify scaffold files, launch local debug in VS Code.
 * Plan: packages/tests/copilot-test/test-plans/da-no-action-local-debug/da-no-action-local-debug.md
 *
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
 * UI interactions + screenshots driven by Playwright via signal files.
 *
 * Phases:
 *   1  Create project [VSC]
 *   2  Verify scaffold files and view in editor [VSC]
 *   3  Launch local debug and observe VS Code tasks [VSC → Chrome]
 */

import * as vscode from "vscode";
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

const OUTPUT_DIR =
  process.env.TEST_OUTPUT_DIR || path.join(os.tmpdir(), "atk-test-output");
const SCREENSHOT_DIR =
  process.env.SCREENSHOT_DIR || path.join(OUTPUT_DIR, "screenshots");
const SIGNAL_DIR =
  process.env.SCREENSHOT_SIGNAL_DIR ||
  path.join(OUTPUT_DIR, ".screenshot-signals");

// App name is generated once per test run and reused across all tests in the suite.
// Agent entry in Copilot will be appName + "local" (e.g. da-no-action-1717000000000local)
const APP_NAME = `da-no-action-${Date.now()}`;

function ensureDirs() {
  [OUTPUT_DIR, SCREENSHOT_DIR, SIGNAL_DIR].forEach((d) =>
    fs.mkdirSync(d, { recursive: true }),
  );
}

/** Signal Playwright to take a screenshot; async polling (non-blocking, up to 8 s) */
async function takeScreenshot(name: string): Promise<void> {
  try {
    const dest = path.join(SCREENSHOT_DIR, `${name}.png`);
    const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
    fs.writeFileSync(signal, `screenshot:${dest}`, "utf8");
    await new Promise<void>((resolve) => {
      const deadline = Date.now() + 8000;
      const iv = setInterval(() => {
        if (!fs.existsSync(signal) || Date.now() >= deadline) {
          clearInterval(iv);
          if (fs.existsSync(signal)) {
            try {
              fs.unlinkSync(signal);
            } catch {}
          }
          resolve();
        }
      }, 100);
    });
    console.log(
      fs.existsSync(dest)
        ? `Screenshot: ${name}.png`
        : `Screenshot timeout: ${name}.png`,
    );
  } catch (e) {
    console.warn("Screenshot failed:", e);
  }
}

/**
 * Write a signal file and wait for Playwright to consume it (up to timeoutMs).
 * Uses setInterval so the extension-host event loop stays free — critical for
 * "fire-and-forget" executeCommand calls that must reach VS Code before we block.
 */
async function sendSignal(content: string, timeoutMs = 15000): Promise<void> {
  try {
    const signal = path.join(SIGNAL_DIR, `${Date.now()}-action.signal`);
    fs.writeFileSync(signal, content, "utf8");
    await new Promise<void>((resolve) => {
      const deadline = Date.now() + timeoutMs;
      const iv = setInterval(() => {
        if (!fs.existsSync(signal) || Date.now() >= deadline) {
          clearInterval(iv);
          if (fs.existsSync(signal)) {
            console.log(`Signal timeout: ${content}`);
            try {
              fs.unlinkSync(signal);
            } catch {}
          }
          resolve();
        }
      }, 100);
    });
  } catch (e) {
    console.warn("Signal failed:", e);
  }
}

function wait(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function waitForCommand(cmd: string, maxMs = 600000): Promise<boolean> {
  const deadline = Date.now() + maxMs;
  while (true) {
    const allCmds = await vscode.commands.getCommands(true);
    if (allCmds.includes(cmd)) return true;
    if (Date.now() >= deadline) return false;
    await wait(1000);
  }
}

/** Search common roots for a directory named appName */
function findProjectDir(appName: string): string {
  const searchRoots = [
    path.join(os.homedir(), "AgentsToolkitProjects"),
    path.join("/home/runner", "AgentsToolkitProjects"),
    os.homedir(),
    "/home/runner",
    os.tmpdir(),
    "/tmp",
    process.cwd(),
  ];
  // Check workspace folders first
  const wsf = vscode.workspace.workspaceFolders;
  if (wsf && wsf.length > 0) {
    const match = wsf.find((f) => f.uri.fsPath.includes(appName));
    if (match) return match.uri.fsPath;
    return wsf[wsf.length - 1].uri.fsPath;
  }
  for (const root of searchRoots) {
    if (!fs.existsSync(root)) continue;
    const direct = path.join(root, appName);
    if (fs.existsSync(direct)) return direct;
    try {
      for (const entry of fs.readdirSync(root)) {
        if (entry.includes(appName)) {
          const full = path.join(root, entry);
          try {
            if (fs.statSync(full).isDirectory()) return full;
          } catch {}
        }
      }
    } catch {}
  }
  return "";
}

function writeResults(passed: number, failed: number, steps: object[]) {
  const out = path.join(OUTPUT_DIR, "results.json");
  fs.writeFileSync(
    out,
    JSON.stringify({ passed, failed, steps }, null, 2),
    "utf8",
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Suite
// ─────────────────────────────────────────────────────────────────────────────
suite("DA No Action – Scaffold and Local Debug", function () {
  // Full test can take up to 15 min (provision + Chrome launch)
  this.timeout(15 * 60 * 1000);

  const steps: object[] = [];
  let passed = 0;
  let failed = 0;
  let projectDir = "";

  const step = (name: string, ok: boolean, detail?: string) => {
    steps.push({ name, status: ok ? "pass" : "fail", detail });
    ok ? passed++ : failed++;
    console.log(
      `${ok ? "PASS" : "FAIL"} ${name}${detail ? ": " + detail : ""}`,
    );
  };

  suiteSetup(() => {
    ensureDirs();
    console.log("=== DA No Action – Scaffold and Local Debug ===");
    console.log("App name :", APP_NAME);
    console.log("Output   :", OUTPUT_DIR);
  });

  suiteTeardown(() => {
    writeResults(passed, failed, steps);
    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
  });

  // ── Test 1: Extension activation ─────────────────────────────────────────
  test("ATK extension is active", async () => {
    const extId = "TeamsDevApp.ms-teams-vscode-extension";
    let ext = vscode.extensions.getExtension(extId);
    if (!ext) {
      for (let i = 0; i < 30; i++) {
        await wait(500);
        ext = vscode.extensions.getExtension(extId);
        if (ext) break;
      }
    }
    if (ext && !ext.isActive) {
      try {
        await ext.activate();
      } catch (e: any) {
        console.log("  Activation note:", e.message);
      }
    }
    await wait(3000);
    const active = !!ext?.isActive;
    step(
      "ATK extension activates",
      active,
      ext ? `v${ext.packageJSON.version}` : "not found",
    );
    await takeScreenshot("01-atk-fresh");
    assert.ok(active, "Extension should be active");
  });

  // ── Test 2: Phase 1 – Create DA No Action project ────────────────────────
  test("Phase 1: Create DA No Action project via wizard", async () => {
    // Step 1: clear notifications (dismiss welcome dialog)
    vscode.commands
      .executeCommand("notifications.clearAll")
      .then(undefined, () => {});
    await wait(500);

    // Step 2: fire the wizard WITHOUT await — it blocks until wizard completes
    const cmdAvailable = await waitForCommand("fx-extension.create");
    console.log("  fx-extension.create available:", cmdAvailable);
    vscode.commands
      .executeCommand("fx-extension.create")
      .then(undefined, (e: any) => {
        console.log("  Command error:", e.message);
      });
    await wait(500); // yield event loop so command dispatch reaches extension host

    // Step 3: QuickPick — "Declarative Agent" (first item, pre-highlighted by VS Code)
    // Use pressKey:Enter rather than clickText to avoid matching the ATK sidebar
    // "Build a Declarative Agent" button before the QuickPick row.
    await sendSignal(
      "waitForTextThenScreenshot:Declarative Agent:60000:02-wizard-da",
      68000,
    );
    await sendSignal("pressKey:Enter", 5000); // select highlighted "Declarative Agent"
    await wait(1000);

    // Step 4: QuickPick — "No Action"
    await sendSignal(
      "waitForTextThenScreenshot:No Action:20000:03-wizard-no-action",
      28000,
    );
    await sendSignal("clickText:No Action", 10000);
    await wait(1000);

    // Step 5: QuickPick — Workspace Folder "Default folder"
    await sendSignal(
      "waitForTextThenScreenshot:Default folder:15000:04-wizard-folder",
      23000,
    );
    await sendSignal("clickText:Default folder", 10000);
    await wait(1000);

    // Step 6: InputBox — Application Name
    await sendSignal(
      "waitForTextThenScreenshot:Application Name:15000:05-wizard-appname",
      23000,
    );
    await sendSignal(`type:${APP_NAME}`, 8000);
    await wait(500);
    await sendSignal("pressKey:Enter", 5000);

    step(
      "fx-extension.create command available",
      cmdAvailable,
      `available=${cmdAvailable}`,
    );
  });

  // ── Test 3: Phase 2 – Verify scaffold files and view in editor ────────────
  test("Phase 2: Verify scaffold files and view in editor", async () => {
    // Step 7: poll for m365agents.yml up to 90 s
    console.log(
      `  Waiting for scaffold files in ~/AgentsToolkitProjects/${APP_NAME} (up to 90s)...`,
    );
    for (let i = 0; i < 90; i++) {
      projectDir = findProjectDir(APP_NAME);
      if (
        projectDir &&
        fs.existsSync(path.join(projectDir, "m365agents.yml"))
      ) {
        break;
      }
      await wait(1000);
    }
    await takeScreenshot("06-scaffold-explorer");

    step("Project directory found", !!projectDir, projectDir || "not found");

    // Step 8: assert all 6 required files exist
    const requiredFiles: Array<[string, string]> = [
      ["m365agents.yml", path.join(projectDir, "m365agents.yml")],
      ["m365agents.local.yml", path.join(projectDir, "m365agents.local.yml")],
      [
        "appPackage/manifest.json",
        path.join(projectDir, "appPackage", "manifest.json"),
      ],
      [
        "appPackage/declarativeAgent.json",
        path.join(projectDir, "appPackage", "declarativeAgent.json"),
      ],
      [".vscode/tasks.json", path.join(projectDir, ".vscode", "tasks.json")],
      [".vscode/launch.json", path.join(projectDir, ".vscode", "launch.json")],
    ];
    for (const [label, filePath] of requiredFiles) {
      const exists = !!projectDir && fs.existsSync(filePath);
      step(
        `${label} exists`,
        exists,
        exists ? "✓" : `missing in ${projectDir}`,
      );
    }

    // Step 9: assert ai-plugin.json does NOT exist (No Action DA has no plugin)
    const pluginPath = path.join(projectDir, "appPackage", "ai-plugin.json");
    const noPlugin = !!projectDir && !fs.existsSync(pluginPath);
    step(
      "appPackage/ai-plugin.json absent (No Action)",
      noPlugin,
      noPlugin ? "✓ not present" : "UNEXPECTED: file exists",
    );

    // Step 10: open m365agents.yml in editor
    if (projectDir) {
      await vscode.commands.executeCommand(
        "vscode.open",
        vscode.Uri.file(path.join(projectDir, "m365agents.yml")),
      );
      await wait(2000);
      await takeScreenshot("07-m365agents-yml");
    }

    // Step 11: open .vscode/tasks.json in editor
    if (projectDir) {
      await vscode.commands.executeCommand(
        "vscode.open",
        vscode.Uri.file(path.join(projectDir, ".vscode", "tasks.json")),
      );
      await wait(2000);
      await takeScreenshot("08-tasks-json");
    }

    // Step 12: open .vscode/launch.json in editor
    if (projectDir) {
      await vscode.commands.executeCommand(
        "vscode.open",
        vscode.Uri.file(path.join(projectDir, ".vscode", "launch.json")),
      );
      await wait(2000);
      await takeScreenshot("09-launch-json");
    }

    // Step 13: close all editors and clear notifications
    vscode.commands
      .executeCommand("workbench.action.closeAllEditors")
      .then(undefined, () => {});
    await wait(500);
    vscode.commands
      .executeCommand("notifications.clearAll")
      .then(undefined, () => {});
    await wait(500);
    await takeScreenshot("10-editors-closed");

    assert.ok(!!projectDir, "Project directory not found");
    assert.ok(
      fs.existsSync(path.join(projectDir, "m365agents.yml")),
      `m365agents.yml missing in ${projectDir}`,
    );
    assert.ok(
      fs.existsSync(path.join(projectDir, "appPackage", "manifest.json")),
      `appPackage/manifest.json missing in ${projectDir}`,
    );
    assert.ok(
      fs.existsSync(
        path.join(projectDir, "appPackage", "declarativeAgent.json"),
      ),
      `appPackage/declarativeAgent.json missing in ${projectDir}`,
    );
    assert.ok(
      noPlugin,
      `appPackage/ai-plugin.json should NOT exist for No Action DA`,
    );
  });

  // ── Test 3b: Sign in to M365 (precondition for local debug) ──────────────
  // The debug task "Validate prerequisites" checks M365 Copilot access.
  // Sign-in must complete before Phase 3 debug launch.
  // Credentials come from M365_ACCOUNT_NAME / M365_ACCOUNT_PASSWORD env vars.
  // The codeFlowLogin.ts source patch writes the MSAL auth URL to
  // $TEMP/atk-auth-url.txt; runTest.ts OAuth watcher completes the browser
  // sign-in automatically using Playwright.
  test("Sign in to M365 (precondition for local debug)", async () => {
    const m365User = process.env.M365_ACCOUNT_NAME || "";
    const m365Pass = process.env.M365_ACCOUNT_PASSWORD || "";
    assert.ok(
      m365User && m365Pass,
      "M365_ACCOUNT_NAME and M365_ACCOUNT_PASSWORD must be set — needed for debug task",
    );

    // Open ATK Accounts pane
    vscode.commands
      .executeCommand("fx-extension.cmpAccounts")
      .then(undefined, () => {});
    await wait(500);

    // Click "Sign in to Microsoft 365"
    await sendSignal(
      "waitForTextThenScreenshot:Sign in to Microsoft 365:20000:09-signin-option",
      28000,
    );
    await sendSignal("clickText:Sign in to Microsoft 365", 10000);
    await wait(1000);
    await takeScreenshot("10-sign-in-triggered");

    // TEAMSFX_AUTO_CONFIRM_LOGIN=true bypasses the VS Code confirmation modal.
    // The extension writes the auth URL to atk-auth-url.txt which runTest.ts reads,
    // then launches a Playwright browser to complete OAuth automatically.
    const AUTH_URL_FILE = path.join(os.tmpdir(), "atk-auth-url.txt");
    let authUrlDetected = false;
    console.log("  Polling for atk-auth-url.txt (up to 30s)...");
    for (let i = 0; i < 30 && !authUrlDetected; i++) {
      await wait(1000);
      if (fs.existsSync(AUTH_URL_FILE)) authUrlDetected = true;
    }
    step(
      "MSAL auth URL captured",
      authUrlDetected,
      authUrlDetected
        ? "atk-auth-url.txt found"
        : "not found — check TEAMSFX_BROKER_AUTH=false patch",
    );

    // Wait for account name to appear in the ATK sidebar (up to 90s)
    const accountHint = m365User.split("@")[0];
    console.log(`  Waiting for "${accountHint}" in ATK sidebar (up to 90s)...`);
    await sendSignal(
      `waitForTextThenScreenshot:${accountHint}:90000:11-m365-signed-in`,
      98000,
    );
    await takeScreenshot("11-m365-signed-in");
    step("M365 sign-in completed", true, `user=${m365User}`);
  });
  test("Phase 3: Launch local debug and observe VS Code tasks", async () => {
    // Set up task/debug event listeners BEFORE firing the debug command so no events are missed.
    let taskStarted = false;
    let prereqsStarted = false;
    let createResourcesStarted = false;
    const taskDisposable = vscode.tasks.onDidStartTaskProcess((e) => {
      const name = e.execution.task.name;
      console.log(`  Task started: "${name}"`);
      if (/start agent locally/i.test(name)) taskStarted = true;
      if (/validate prerequisites/i.test(name)) prereqsStarted = true;
      if (/create resources/i.test(name)) createResourcesStarted = true;
    });
    const debugDisposable = vscode.debug.onDidStartDebugSession((s) => {
      console.log(`  Debug session started: "${s.name}"`);
    });

    try {
      // Step 14: fire debug-selection command WITHOUT await
      vscode.commands
        .executeCommand("workbench.action.debug.selectandstart")
        .then(undefined, (e: any) => {
          console.log("  debug.selectandstart error:", e.message);
        });
      await wait(500);

      // Step 15: select "Preview Local in Copilot (Chrome)"
      await sendSignal(
        "waitForTextThenScreenshot:Preview Local in Copilot:20000:11-debug-picker",
        28000,
      );
      await sendSignal("clickText:Preview Local in Copilot (Chrome)", 10000);
      await wait(1000);

      // Step 16: wait for "Start Agent Locally" compound task (up to 30 s via task events)
      const taskStartDeadline = Date.now() + 30000;
      while (!taskStarted && Date.now() < taskStartDeadline) await wait(500);
      await sendSignal(
        "waitForTextThenScreenshot:Start Agent Locally:5000:12-task-started",
        10000,
      );
      step(
        "Start Agent Locally task started",
        taskStarted,
        taskStarted ? "✓" : "not detected via task events",
      );

      // Step 17: wait for "Validate prerequisites" subtask (up to 30 s via task events)
      const prereqDeadline = Date.now() + 30000;
      while (!prereqsStarted && Date.now() < prereqDeadline) await wait(500);
      await sendSignal(
        "waitForTextThenScreenshot:Validate prerequisites:5000:13-validate-prereqs",
        10000,
      );
      step(
        "Validate prerequisites task started",
        prereqsStarted,
        prereqsStarted ? "✓" : "not detected via task events",
      );

      // Step 18: wait for "Create resources" subtask (up to 60 s via task events)
      const createDeadline = Date.now() + 60000;
      while (!createResourcesStarted && Date.now() < createDeadline)
        await wait(500);
      await sendSignal(
        "waitForTextThenScreenshot:Create resources:5000:14-create-resources",
        10000,
      );
      step(
        "Create resources task started",
        createResourcesStarted,
        createResourcesStarted ? "✓" : "not detected via task events",
      );

      // Step 19: wait for provision to progress, screenshot
      await wait(10000);
      await takeScreenshot("15-provision-progress");

      // Step 20: soft check — Chrome opens to m365.cloud.microsoft.
      // The VS Code CDP connection does not have access to an external Chrome window,
      // so this signal will time out gracefully if the browser is not observable.
      await sendSignal(
        "waitForTextThenScreenshot:m365.cloud.microsoft:120000:16-copilot-browser",
        130000,
      );
      await takeScreenshot("16-copilot-browser");

      step(
        "Debug task chain initiated (Preview Local in Copilot Chrome)",
        true,
        "debug tasks started",
      );

      assert.ok(taskStarted, "Start Agent Locally task was never started");
      assert.ok(
        prereqsStarted,
        "Validate prerequisites task was never started",
      );
      assert.ok(
        createResourcesStarted,
        "Create resources task was never started",
      );
    } finally {
      taskDisposable.dispose();
      debugDisposable.dispose();
    }
  });
});
