// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * simple-bot-create.test.ts
 * TC-001: Create Teams Bot template via the ATK VSCode wizard (UI-driven).
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
 * Screenshots and UI interactions are driven by Playwright via signal files.
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

function ensureDirs() {
  [OUTPUT_DIR, SCREENSHOT_DIR, SIGNAL_DIR].forEach((d) =>
    fs.mkdirSync(d, { recursive: true }),
  );
}

/** Signal Playwright to take a screenshot; blocks up to 8s */
function takeScreenshot(name: string): void {
  try {
    const dest = path.join(SCREENSHOT_DIR, `${name}.png`);
    const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
    fs.writeFileSync(signal, `screenshot:${dest}`, "utf8");
    const deadline = Date.now() + 8000;
    while (fs.existsSync(signal) && Date.now() < deadline) {
      const end = Date.now() + 100;
      while (Date.now() < end) {
        /* busy wait */
      }
    }
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
 * Send an action signal to Playwright and wait for it to be processed.
 * content: "clickText:Bot", "type:my-app", "pressKey:Enter", etc.
 */
function sendSignal(content: string, timeoutMs = 15000): void {
  try {
    const signal = path.join(SIGNAL_DIR, `${Date.now()}-action.signal`);
    fs.writeFileSync(signal, content, "utf8");
    const deadline = Date.now() + timeoutMs;
    while (fs.existsSync(signal) && Date.now() < deadline) {
      const end = Date.now() + 100;
      while (Date.now() < end) {
        /* busy wait */
      }
    }
    if (fs.existsSync(signal)) {
      console.log(`Signal timeout: ${content}`);
      try {
        fs.unlinkSync(signal);
      } catch {}
    }
  } catch (e) {
    console.warn("Signal failed:", e);
  }
}

function wait(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function waitForCommand(cmd: string, maxMs = 20000): Promise<boolean> {
  const deadline = Date.now() + maxMs;
  while (Date.now() < deadline) {
    const allCmds = await vscode.commands.getCommands(true);
    if (allCmds.includes(cmd)) return true;
    await wait(500);
  }
  return false;
}

function writeResults(passed: number, failed: number, steps: object[]) {
  const out = path.join(OUTPUT_DIR, "results.json");
  fs.writeFileSync(
    out,
    JSON.stringify({ passed, failed, steps }, null, 2),
    "utf8",
  );
}

suite("ATK Teams Bot Template Creation (UI Wizard)", function () {
  this.timeout(8 * 60 * 1000);

  const steps: object[] = [];
  let passed = 0;
  let failed = 0;
  let createdProjectDir = "";

  const step = (name: string, ok: boolean, detail?: string) => {
    steps.push({ name, status: ok ? "pass" : "fail", detail });
    ok ? passed++ : failed++;
    console.log(
      `${ok ? "PASS" : "FAIL"} ${name}${detail ? ": " + detail : ""}`,
    );
  };

  suiteSetup(() => {
    ensureDirs();
    console.log("=== ATK Teams Bot Template Test (UI Wizard) ===");
    console.log("Output:", OUTPUT_DIR);
  });

  suiteTeardown(() => {
    writeResults(passed, failed, steps);
    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
  });

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
    takeScreenshot("01-extension-active");
    assert.ok(active, "Extension should be active");
  });

  test("Navigate wizard to create Teams Bot template", async () => {
    const cmdAvailable = await waitForCommand("fx-extension.create", 15000);
    console.log("  fx-extension.create available:", cmdAvailable);

    // Fire command without awaiting — wizard blocks until user completes it
    vscode.commands.executeCommand("fx-extension.create").catch((e: any) => {
      console.log("  Command error:", e.message);
    });

    // Wait for wizard first step to appear, then screenshot (QuickPick visible)
    sendSignal("waitForText:Teams Agents and Apps", 12000);
    await wait(200);
    takeScreenshot("02-wizard-open"); // shows first wizard QuickPick

    // Step 1: Select "Teams Agents and Apps" - screenshot BEFORE click
    console.log("  Clicking: Teams Agents and Apps");
    takeScreenshot("03-teams-agents-apps"); // QuickPick visible with Teams Agents and Apps
    sendSignal("clickText:Teams Agents and Apps", 8000);
    await wait(300);

    // Step 2: Bot - screenshot BEFORE click
    console.log("  Clicking: Bot");
    sendSignal("waitForText:Bot", 8000);
    await wait(200);
    takeScreenshot("04-bot-selected"); // QuickPick showing Bot option
    sendSignal("clickText:Bot", 8000);
    await wait(300);

    // Step 3: Simple Bot - screenshot BEFORE click
    console.log("  Clicking: Simple Bot");
    sendSignal("waitForText:Simple Bot", 8000);
    await wait(200);
    takeScreenshot("05-simple-bot"); // QuickPick showing Simple Bot option
    sendSignal("clickText:Simple Bot", 8000);
    await wait(300);

    // Step 4: TypeScript - screenshot BEFORE click
    console.log("  Clicking: TypeScript");
    sendSignal("waitForText:TypeScript", 8000);
    await wait(200);
    takeScreenshot("06-typescript"); // QuickPick showing TypeScript option
    sendSignal("clickText:TypeScript", 8000);
    await wait(300);

    // Step 5: Workspace Folder QuickPick - screenshot BEFORE click
    console.log("  Selecting default folder");
    sendSignal("waitForText:Workspace Folder", 15000);
    await wait(200);
    takeScreenshot("07-workspace-folder"); // QuickPick showing Default folder + Browse...
    sendSignal("clickText:Default folder", 8000);
    await wait(300);

    // Step 6: Application Name InputBox - screenshot BEFORE typing
    console.log("  Typing app name");
    sendSignal("waitForText:Application Name", 8000);
    await wait(200);
    takeScreenshot("08-app-name-input"); // InputBox visible and empty
    sendSignal("type:test-teams-bot-001", 5000);
    await wait(300);
    sendSignal("pressKey:Enter", 3000);
    await wait(90000); // scaffold + new window; 90s for slow CI
    takeScreenshot("09-project-created");

    step(
      "Navigate wizard to create Teams Bot template",
      cmdAvailable,
      `command=${cmdAvailable}`,
    );
  });

  test("Verify scaffolded project files exist", async () => {
    const appName = "test-teams-bot-001";
    let projectDir = "";

    // 1. Check workspace folders (project opened in current window)
    const wsf = vscode.workspace.workspaceFolders;
    if (wsf && wsf.length > 0) {
      // find folder that contains our app name
      const match = wsf.find((f) => f.uri.fsPath.includes(appName));
      projectDir = match ? match.uri.fsPath : wsf[wsf.length - 1].uri.fsPath;
      createdProjectDir = projectDir;
      console.log("  Workspace folder:", projectDir);
    }

    // 2. Broad filesystem search for the project directory
    if (!projectDir || !fs.existsSync(projectDir)) {
      const agentsDir = path.join(os.homedir(), "AgentsToolkitProjects");
      const searchRoots = [
        agentsDir, // ATK default: ~/AgentsToolkitProjects
        path.join("/home/runner", "AgentsToolkitProjects"),
        os.homedir(),
        "/home/runner",
        os.tmpdir(),
        "/tmp",
        process.cwd(),
      ];
      outer: for (const root of searchRoots) {
        if (!fs.existsSync(root)) continue;
        // Direct child with matching name
        const direct = path.join(root, appName);
        if (fs.existsSync(direct)) {
          projectDir = direct;
          break;
        }
        // Any child directory containing our app name
        try {
          for (const entry of fs.readdirSync(root)) {
            if (entry.includes("test-teams-bot") || entry.includes(appName)) {
              const full = path.join(root, entry);
              try {
                if (fs.statSync(full).isDirectory()) {
                  projectDir = full;
                  break outer;
                }
              } catch {}
            }
          }
        } catch {}
      }
      console.log("  Filesystem search found:", projectDir || "none");
    }

    const expectedFiles = [
      "m365agents.yml",
      "package.json",
      "index.ts",
      "appPackage/manifest.json",
    ];

    let allFound = true;
    for (const f of expectedFiles) {
      const exists = projectDir
        ? fs.existsSync(path.join(projectDir, f))
        : false;
      step(`File: ${f}`, exists, exists ? "✓" : `not found in ${projectDir}`);
      if (!exists) allFound = false;
    }

    takeScreenshot("10-final-state");
    assert.ok(allFound, `Expected project files missing in ${projectDir}`);
  });
});
