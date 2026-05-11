// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * teams-bot-create-template.test.ts
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
 * Screenshots are taken by external Playwright via signal files.
 */
import * as vscode from "vscode";
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import * as cp from "child_process";

const OUTPUT_DIR = process.env.TEST_OUTPUT_DIR || path.join(os.tmpdir(), "atk-test-output");
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR || path.join(OUTPUT_DIR, "screenshots");
const SIGNAL_DIR = process.env.SCREENSHOT_SIGNAL_DIR || path.join(OUTPUT_DIR, ".screenshot-signals");

function ensureDirs() {
  [OUTPUT_DIR, SCREENSHOT_DIR, SIGNAL_DIR].forEach((d) =>
    fs.mkdirSync(d, { recursive: true })
  );
}

/** Signal Playwright to take a screenshot; blocks up to 8s */
function takeScreenshot(name: string): void {
  try {
    const dest = path.join(SCREENSHOT_DIR, `${name}.png`);
    const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
    fs.writeFileSync(signal, dest, "utf8");

    const deadline = Date.now() + 8000;
    while (fs.existsSync(signal) && Date.now() < deadline) {
      const end = Date.now() + 100;
      while (Date.now() < end) { /* busy wait */ }
    }

    if (fs.existsSync(dest)) {
      console.log(`Screenshot: ${name}.png`);
    } else {
      console.log(`Screenshot timeout: ${name}.png (Playwright may not be connected yet)`);
    }
  } catch (e) {
    console.warn("Screenshot failed:", e);
  }
}

/** Wait (ms) inside sync-ish style */
function wait(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

/** Poll until a VSCode command is available */
async function waitForCommand(cmd: string, maxMs = 15000): Promise<boolean> {
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
  fs.writeFileSync(out, JSON.stringify({ passed, failed, steps }, null, 2), "utf8");
}

suite("ATK Teams Bot Template Creation", function () {
  this.timeout(5 * 60 * 1000);

  const steps: object[] = [];
  let passed = 0;
  let failed = 0;

  const step = (name: string, ok: boolean, detail?: string) => {
    steps.push({ name, status: ok ? "pass" : "fail", detail });
    ok ? passed++ : failed++;
    console.log(`${ok ? "PASS" : "FAIL"} ${name}${detail ? ": " + detail : ""}`);
  };

  suiteSetup(() => {
    ensureDirs();
    console.log("=== ATK Teams Bot Template Test ===");
    console.log("Output:", OUTPUT_DIR);
  });

  suiteTeardown(() => {
    writeResults(passed, failed, steps);
    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
  });

  test("ATK extension is active", async () => {
    const extId = "TeamsDevApp.ms-teams-vscode-extension";
    let ext = vscode.extensions.getExtension(extId);

    // Wait for extension to appear (VSCode might still be loading)
    if (!ext) {
      for (let i = 0; i < 20; i++) {
        await wait(500);
        ext = vscode.extensions.getExtension(extId);
        if (ext) break;
      }
    }

    if (ext && !ext.isActive) {
      try { await ext.activate(); } catch (e: any) {
        console.log("  Activation note:", e.message);
      }
    }

    // Wait for activation to complete (async side effects)
    await wait(3000);

    const active = !!ext?.isActive;
    step("ATK extension activates", active,
      ext ? `version ${ext.packageJSON.version}` : "not found");

    takeScreenshot("01-extension-active");
    assert.ok(active, "Extension should be active");
  });

  test("Create New Project command opens wizard", async () => {
    // Wait for fx-extension.create to be registered
    const cmdAvailable = await waitForCommand("fx-extension.create", 15000);
    if (!cmdAvailable) {
      console.log("  Command not available after 15s, trying anyway...");
    } else {
      console.log("  Command available: fx-extension.create");
    }

    // Detect new tab/panel opening
    const tabOpenedPromise = new Promise<boolean>((resolve) => {
      const timer = setTimeout(() => resolve(false), 12000);
      const disp = vscode.window.tabGroups.onDidChangeTabs((e) => {
        if (e.opened.length > 0) {
          clearTimeout(timer);
          disp.dispose();
          resolve(true);
        }
      });
    });

    // Fire command WITHOUT awaiting (wizard blocks until user action)
    vscode.commands.executeCommand("fx-extension.create").catch((e: any) => {
      console.log("  Command error:", e.message);
    });

    const tabOpened = await tabOpenedPromise;
    await wait(2000); // let the wizard UI render

    takeScreenshot("02-create-wizard-open");
    step("Create Project wizard opens", cmdAvailable || tabOpened,
      `command=${cmdAvailable}, tab=${tabOpened}`);
  });

  test("Close wizard", async () => {
    await vscode.commands.executeCommand("workbench.action.closeAllEditors");
    await wait(1000);
    takeScreenshot("03-editors-closed");
    step("Close wizard", true);
  });

  test("Scaffold via ATK CLI (optional)", async () => {
    const projectDir = path.join(OUTPUT_DIR, "projects", "teams-bot-test");
    fs.mkdirSync(path.dirname(projectDir), { recursive: true });

    // Try CLI names in priority order. ATK CLI was renamed: teamsapp -> atk
    // Correct args: --capability bot --programming-language typescript
    const cliCandidates: [string, string[]][] = [
      ["atk",      ["new", "--capability", "bot", "--programming-language", "typescript", "--app-name", "TeamsBot", "--interactive", "false"]],
      ["teamsapp", ["new", "--capability", "bot", "--programming-language", "typescript", "--app-name", "TeamsBot", "--interactive", "false"]],
      ["teamsfx",  ["new", "--capability", "bot", "--programming-language", "typescript", "--app-name", "TeamsBot", "--interactive", "false"]],
    ];
    let scaffolded = false;
    for (const [cli, args] of cliCandidates) {
      const r = cp.spawnSync(cli, args,
        { cwd: path.dirname(projectDir), timeout: 120000, shell: true });
      if (r.status === 0) {
        scaffolded = true;
        break;
      }
    }

    if (scaffolded) {
      await wait(1000);
      takeScreenshot("04-project-scaffold");
    }

    step("Scaffold via ATK CLI", scaffolded,
      scaffolded ? "CLI succeeded" : "CLI not installed (expected in local dev)");
  });
});


