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

/** Signal Playwright to take a screenshot; polls async (non-blocking, up to 8s) */
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
 * Send an action signal to Playwright and wait for it to be processed.
 * Uses async polling (setInterval) so the extension host event loop stays free,
 * allowing VS Code commands dispatched before this call to actually execute.
 * content: "clickText:Bot", "type:my-app", "pressKey:Enter", etc.
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
  // NOTE: getCommands() may block for the entire unresponsive period when the
  // extension host is busy (e.g. ATK doing heavy JS init). The deadline check
  // BEFORE getCommands() can be stale. We therefore also check AFTER the call.
  const deadline = Date.now() + maxMs;
  while (true) {
    const allCmds = await vscode.commands.getCommands(true);
    if (allCmds.includes(cmd)) return true;
    if (Date.now() >= deadline) return false;
    await wait(1000);
  }
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
  this.timeout(12 * 60 * 1000);

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
    await takeScreenshot("01-extension-active");
    assert.ok(active, "Extension should be active");
  });

  test("Navigate wizard to create Teams Bot template", async () => {
    const cmdAvailable = await waitForCommand("fx-extension.create");
    console.log("  fx-extension.create available:", cmdAvailable);

    // Fire command without awaiting — wizard blocks until user completes it.
    // IMPORTANT: sendSignal() is async (setInterval-based), so the event loop
    // stays free and the executeCommand dispatch can reach the extension host.
    vscode.commands.executeCommand("fx-extension.create").catch((e: any) => {
      console.log("  Command error:", e.message);
    });
    // Yield the event loop so the command dispatch reaches the extension host.
    await wait(500);

    // ATK v6.8+ wizard flow for Simple Bot:
    //   Step 1 (project-type):   "Teams Agents and Apps"
    //   Step 2 (teams-app-type): "Other Teams Capabilities"
    //   Step 3 (teams-other):    "Simple Bot"
    //   Step 4 (folder):         "Default folder"
    //   Step 5 (name):            type "test-teams-bot-001" + Enter

    // Step 1: With baked extension (native-FS), QuickPick appears in <60s.
    // sendSignal is now async — event loop stays free so the command can execute.
    await sendSignal(
      "waitForTextThenScreenshot:Teams Agents and Apps:60000:02-step1-project-type",
      68000,
    );
    await sendSignal("clickText:Teams Agents and Apps", 10000);
    await wait(1000);

    // Step 2: "Other Teams Capabilities"
    await sendSignal(
      "waitForTextThenScreenshot:Other Teams Capabilities:20000:03-step2-teams-app-type",
      28000,
    );
    await sendSignal("clickText:Other Teams Capabilities", 10000);
    await wait(1000);

    // Step 3: "Simple Bot"
    await sendSignal(
      "waitForTextThenScreenshot:Simple Bot:15000:04-step3-simple-bot",
      23000,
    );
    await sendSignal("clickText:Simple Bot", 10000);
    await wait(1000);

    // Step 4: Programming Language — TypeScript / JavaScript / Python
    await sendSignal(
      "waitForTextThenScreenshot:TypeScript:15000:05-step4-language",
      23000,
    );
    await sendSignal("clickText:TypeScript", 10000);
    await wait(1000);

    // Step 5: Workspace Folder
    await sendSignal(
      "waitForTextThenScreenshot:Default folder:15000:06-step5-workspace-folder",
      23000,
    );
    await sendSignal("clickText:Default folder", 10000);
    await wait(1000);

    // Step 6: Application Name InputBox
    await sendSignal(
      "waitForTextThenScreenshot:Application Name:15000:07-step6-app-name",
      23000,
    );
    await sendSignal("type:test-teams-bot-001", 8000);
    await wait(500);
    await sendSignal("pressKey:Enter", 5000);
    // Scaffold takes up to 90s; capture intermediate screenshots to catch errors
    await wait(15000);
    await takeScreenshot("08a-scaffold-15s");
    await wait(15000);
    await takeScreenshot("08b-scaffold-30s");
    await wait(30000);
    await takeScreenshot("08c-scaffold-60s");
    await wait(30000); // total ~90s
    await takeScreenshot("10-project-created");

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

    // ATK creates the folder first then async-copies template files.
    // Poll up to 60s for sentinel file before checking all files.
    if (projectDir) {
      const sentinel = path.join(projectDir, "m365agents.yml");
      console.log("  Waiting for scaffold files (up to 60s)...");
      for (let i = 0; i < 60; i++) {
        if (fs.existsSync(sentinel)) break;
        await new Promise((r) => setTimeout(r, 1000));
      }
    }

    let allFound = true;
    for (const f of expectedFiles) {
      const exists = projectDir
        ? fs.existsSync(path.join(projectDir, f))
        : false;
      step(`File: ${f}`, exists, exists ? "✓" : `not found in ${projectDir}`);
      if (!exists) allFound = false;
    }

    await takeScreenshot("09-final-state");
    assert.ok(allFound, `Expected project files missing in ${projectDir}`);
  });
});
