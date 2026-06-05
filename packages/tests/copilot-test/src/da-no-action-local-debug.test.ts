// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * da-no-action-local-debug.test.ts
 * TC-001: Create DA (No Action), verify scaffold files, sign in to M365, launch local debug.
 * Plan: packages/tests/copilot-test/test-plans/da-no-action-local-debug/da-no-action-local-debug.md
 *
 * Two-window design:
 *   Window 1 (no project workspace): runs Phase 1 (wizard) → writes scaffold-done.json → exits
 *   Window 2 (scaffolded project):   reads scaffold-done.json → runs Phase 2+3+4 → exits
 *
 * ATK automatically opens a new VS Code window after scaffold with the project as
 * workspace root. That window is perfect for Phase 2 onward: ATK is already active,
 * launch.json is already in the workspace, and no updateWorkspaceFolders hacks needed.
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

// Marker file written by Window 1, read by Window 2
const SCAFFOLD_MARKER = path.join(OUTPUT_DIR, "scaffold-done.json");

// App name for Window 1 (fresh timestamp each run)
const APP_NAME = `da-no-action-${Date.now()}`;

function ensureDirs() {
  [OUTPUT_DIR, SCREENSHOT_DIR, SIGNAL_DIR].forEach((d) =>
    fs.mkdirSync(d, { recursive: true }),
  );
}

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
          try {
            if (fs.existsSync(signal)) fs.unlinkSync(signal);
          } catch {}
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
            console.log(`Signal timeout: ${content.slice(0, 60)}`);
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

async function waitForCommand(cmd: string, maxMs = 60000): Promise<boolean> {
  const deadline = Date.now() + maxMs;
  while (Date.now() < deadline) {
    const cmds = await vscode.commands.getCommands(true);
    if (cmds.includes(cmd)) return true;
    await wait(1000);
  }
  return false;
}

function writeResults(passed: number, failed: number, steps: object[]) {
  fs.writeFileSync(
    path.join(OUTPUT_DIR, "results.json"),
    JSON.stringify({ passed, failed, steps }, null, 2),
    "utf8",
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Determine which window we are in:
//   isWindow2 = scaffold-done.json exists → this is the new project window
// ─────────────────────────────────────────────────────────────────────────────
const isWindow2 = fs.existsSync(SCAFFOLD_MARKER);
let scaffoldedProjectDir = "";
if (isWindow2) {
  try {
    const marker = JSON.parse(fs.readFileSync(SCAFFOLD_MARKER, "utf8"));
    scaffoldedProjectDir = marker.projectDir ?? "";
    console.log(`[Window 2] project dir: ${scaffoldedProjectDir}`);
  } catch {}
}

// ─────────────────────────────────────────────────────────────────────────────
// Suite
// ─────────────────────────────────────────────────────────────────────────────
suite("DA No Action – Scaffold and Local Debug", function () {
  this.timeout(15 * 60 * 1000);

  const steps: object[] = [];
  let passed = 0;
  let failed = 0;

  const step = (name: string, ok: boolean, detail?: string) => {
    steps.push({ name, status: ok ? "pass" : "fail", detail });
    ok ? passed++ : failed++;
    console.log(
      `${ok ? "PASS" : "FAIL"} ${name}${detail ? ": " + detail : ""}`,
    );
  };

  suiteSetup(() => {
    ensureDirs();
    console.log(
      `=== DA No Action ${isWindow2 ? "[Window 2 – project]" : "[Window 1 – wizard]"} ===`,
    );
    console.log("Output:", OUTPUT_DIR);
  });

  suiteTeardown(() => {
    // Only write results from the final window (Window 2), or Window 1 if no Window 2
    if (isWindow2 || !fs.existsSync(SCAFFOLD_MARKER)) {
      writeResults(passed, failed, steps);
    }
    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
  });

  // ── Step 0: ATK extension active (both windows) ───────────────────────────
  test("ATK extension is active", async () => {
    const extId = "TeamsDevApp.ms-teams-vscode-extension";
    let ext = vscode.extensions.getExtension(extId);
    for (let i = 0; i < 30 && !ext?.isActive; i++) {
      await wait(1000);
      ext = vscode.extensions.getExtension(extId);
    }
    if (ext && !ext.isActive) {
      try {
        await ext.activate();
      } catch {}
    }
    await wait(2000);
    const active = !!ext?.isActive;
    step(
      "ATK extension activates",
      active,
      ext ? `v${ext.packageJSON.version}` : "not found",
    );
    await takeScreenshot("01-atk-fresh");
    assert.ok(active, "ATK extension must be active");
  });

  if (!isWindow2) {
    // ── WINDOW 1: Phase 1 – wizard ──────────────────────────────────────────
    test("Phase 1: Create DA No Action project via wizard", async () => {
      vscode.commands
        .executeCommand("notifications.clearAll")
        .then(undefined, () => {});
      await wait(500);

      const cmdAvailable = await waitForCommand("fx-extension.create");
      vscode.commands
        .executeCommand("fx-extension.create")
        .then(undefined, () => {});
      await wait(500);

      // Step 3: Declarative Agent
      await sendSignal(
        "waitForTextThenScreenshot:Declarative Agent:60000:02-wizard-da",
        68000,
      );
      await sendSignal("pressKey:Enter", 5000);
      await wait(800);

      // Step 4: No Action
      await sendSignal(
        "waitForTextThenScreenshot:No Action:20000:03-wizard-no-action",
        28000,
      );
      await wait(500);
      await sendSignal("clickText:No Action", 12000);
      await wait(800);

      // Step 5: Default folder
      await sendSignal(
        "waitForTextThenScreenshot:Default folder:20000:04-wizard-folder",
        28000,
      );
      await wait(500);
      await sendSignal("clickText:Default folder", 12000);
      await wait(800);

      // Step 6: Application Name
      await sendSignal(
        "waitForTextThenScreenshot:Application Name:15000:05-wizard-appname",
        23000,
      );
      await sendSignal(`type:${APP_NAME}`, 8000);
      await wait(300);
      await sendSignal("pressKey:Enter", 5000);

      // ATK now opens the project in a NEW VS Code window and this window closes.
      // Write marker so Window 2 can pick up where we left off.
      const projectDir = path.join(
        os.homedir(),
        "AgentsToolkitProjects",
        APP_NAME,
      );
      // Poll briefly for the folder to confirm scaffold started
      for (let i = 0; i < 30; i++) {
        if (fs.existsSync(projectDir)) break;
        await wait(1000);
      }
      fs.writeFileSync(
        SCAFFOLD_MARKER,
        JSON.stringify({
          appName: APP_NAME,
          projectDir,
          createdAt: Date.now(),
        }),
        "utf8",
      );
      console.log(`  Marker written: ${SCAFFOLD_MARKER}`);
      await takeScreenshot("06-scaffold-started");
      step("fx-extension.create available", cmdAvailable);
      step("Scaffold marker written", true, `projectDir=${projectDir}`);
    });
  } else {
    // ── WINDOW 2: Phases 2 + 3 + 4 ─────────────────────────────────────────
    // This window was opened by ATK with the project already as workspace root.

    test("Phase 2: Verify scaffold files and open in editor", async () => {
      // Wait for scaffold to complete if still in progress
      const projectDir = scaffoldedProjectDir;
      const sentinel = path.join(projectDir, "m365agents.yml");
      console.log(`  Waiting for scaffold at: ${projectDir}`);
      for (let i = 0; i < 90 && !fs.existsSync(sentinel); i++) await wait(1000);
      await takeScreenshot("06-scaffold-explorer");

      const required = [
        "m365agents.yml",
        "m365agents.local.yml",
        "appPackage/manifest.json",
        "appPackage/declarativeAgent.json",
        ".vscode/tasks.json",
        ".vscode/launch.json",
      ];
      for (const rel of required) {
        const exists = fs.existsSync(path.join(projectDir, rel));
        step(
          `${rel} exists`,
          exists,
          exists ? "✓" : `missing in ${projectDir}`,
        );
      }
      const noPlugin = !fs.existsSync(
        path.join(projectDir, "appPackage", "ai-plugin.json"),
      );
      step(
        "appPackage/ai-plugin.json absent (No Action)",
        noPlugin,
        noPlugin ? "✓" : "UNEXPECTED: present",
      );

      // Open key files in editor for screenshot evidence (steps 10-12 of plan)
      await vscode.commands.executeCommand(
        "vscode.open",
        vscode.Uri.file(path.join(projectDir, "m365agents.yml")),
      );
      await wait(2000);
      await takeScreenshot("07-m365agents-yml");
      await vscode.commands.executeCommand(
        "vscode.open",
        vscode.Uri.file(path.join(projectDir, ".vscode", "tasks.json")),
      );
      await wait(2000);
      await takeScreenshot("08-tasks-json");
      await vscode.commands.executeCommand(
        "vscode.open",
        vscode.Uri.file(path.join(projectDir, ".vscode", "launch.json")),
      );
      await wait(2000);
      await takeScreenshot("09-launch-json");
      vscode.commands
        .executeCommand("workbench.action.closeAllEditors")
        .then(undefined, () => {});
      vscode.commands
        .executeCommand("notifications.clearAll")
        .then(undefined, () => {});
      await wait(500);
      await takeScreenshot("10-editors-closed");

      assert.ok(
        fs.existsSync(sentinel),
        `m365agents.yml missing in ${projectDir}`,
      );
      assert.ok(noPlugin, "ai-plugin.json should NOT exist for No Action DA");
    });

    test("Sign in to M365 (precondition for local debug)", async () => {
      const m365User = process.env.M365_ACCOUNT_NAME || "";
      const m365Pass = process.env.M365_ACCOUNT_PASSWORD || "";
      assert.ok(
        m365User && m365Pass,
        "M365_ACCOUNT_NAME and M365_ACCOUNT_PASSWORD must be set",
      );

      // Open ATK Accounts pane
      vscode.commands
        .executeCommand("fx-extension.cmpAccounts")
        .then(undefined, () => {});
      await wait(500);

      await sendSignal(
        "waitForTextThenScreenshot:Sign in to Microsoft 365:20000:09-signin-option",
        28000,
      );
      await sendSignal("clickText:Sign in to Microsoft 365", 10000);
      await wait(1000);
      await takeScreenshot("10-sign-in-triggered");

      // Poll for auth URL written by codeFlowLogin.ts patch
      const AUTH_URL_FILE = path.join(os.tmpdir(), "atk-auth-url.txt");
      let authUrlDetected = false;
      for (let i = 0; i < 30 && !authUrlDetected; i++) {
        await wait(1000);
        if (fs.existsSync(AUTH_URL_FILE)) authUrlDetected = true;
      }
      step(
        "MSAL auth URL captured",
        authUrlDetected,
        authUrlDetected ? "atk-auth-url.txt found" : "not found",
      );

      // runTest.ts OAuth watcher completes browser sign-in automatically.
      // Wait for account name in ATK sidebar (up to 90s).
      const accountHint = m365User.split("@")[0];
      await sendSignal(
        `waitForTextThenScreenshot:${accountHint}:90000:11-m365-signed-in`,
        98000,
      );
      await takeScreenshot("11-m365-signed-in");
      step("M365 sign-in completed", true, `user=${m365User}`);
    });

    test("Phase 3: Launch local debug and observe VS Code tasks", async () => {
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
        console.log(`  Debug session: "${s.name}"`);
      });

      try {
        // The project is already the workspace root — startDebugging works directly
        const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
        console.log(
          `  workspaceFolder: ${workspaceFolder?.uri.fsPath ?? "none"}`,
        );
        await vscode.commands.executeCommand("workbench.view.debug");
        await wait(1000);
        await takeScreenshot("11-debug-panel");

        const debugStarted = await vscode.debug.startDebugging(
          workspaceFolder,
          "Preview Local in Copilot (Chrome)",
        );
        console.log(`  debug.startDebugging: ${debugStarted}`);
        step(
          "Debug session started",
          debugStarted,
          debugStarted ? "✓" : "failed",
        );
        await wait(1000);

        // Step 16: "Start Agent Locally" compound task
        // NOTE: compound tasks (dependsOn) do NOT fire onDidStartTaskProcess.
        // Detect them by waiting for any sub-task to fire instead.
        const t1 = Date.now() + 30000;
        while (!prereqsStarted && !createResourcesStarted && Date.now() < t1)
          await wait(500);
        const compoundStarted = prereqsStarted || createResourcesStarted;
        await sendSignal(
          "waitForTextThenScreenshot:Start Agent Locally:5000:12-task-started",
          10000,
        );
        step(
          "Start Agent Locally task started (via sub-task proxy)",
          compoundStarted,
          compoundStarted ? "✓ (sub-task fired)" : "no sub-task detected",
        );

        // Step 17: "Validate prerequisites"
        const t2 = Date.now() + 30000;
        while (!prereqsStarted && Date.now() < t2) await wait(500);
        await sendSignal(
          "waitForTextThenScreenshot:Validate prerequisites:5000:13-validate-prereqs",
          10000,
        );
        step(
          "Validate prerequisites task started",
          prereqsStarted,
          prereqsStarted ? "✓" : "not detected",
        );

        // Step 18: "Create resources"
        const t3 = Date.now() + 60000;
        while (!createResourcesStarted && Date.now() < t3) await wait(500);
        await sendSignal(
          "waitForTextThenScreenshot:Create resources:5000:14-create-resources",
          10000,
        );
        step(
          "Create resources task started",
          createResourcesStarted,
          createResourcesStarted ? "✓" : "not detected",
        );

        // Step 19: provision progress
        await wait(10000);
        await takeScreenshot("15-provision-progress");

        // Step 20: Chrome opens to m365.cloud.microsoft (soft check — external browser)
        await sendSignal(
          "waitForTextThenScreenshot:m365.cloud.microsoft:120000:16-copilot-browser",
          130000,
        );
        await takeScreenshot("16-copilot-browser");

        // Clean up marker now that this window has finished
        try {
          fs.unlinkSync(SCAFFOLD_MARKER);
        } catch {}

        assert.ok(debugStarted, "debug.startDebugging returned false");
        assert.ok(compoundStarted, "Start Agent Locally: no sub-task fired");
        assert.ok(prereqsStarted, "Validate prerequisites task never started");
        assert.ok(
          createResourcesStarted,
          "Create resources task never started",
        );
      } finally {
        taskDisposable.dispose();
        debugDisposable.dispose();
      }
    });
  }
});
