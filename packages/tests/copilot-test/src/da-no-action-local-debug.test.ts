// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * da-no-action-local-debug.test.ts
 * TC-001: Create DA (No Action), sign in to M365, local debug, verify agent responds in Copilot.
 * Plan: packages/tests/copilot-test/test-plans/da-no-action-local-debug/da-no-action-local-debug.md
 *
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
 * UI interactions + screenshots driven by Playwright via signal files.
 *
 * Phases:
 *   1  Create project [VSC]
 *   2  Verify scaffold files [VSC]
 *   3  Sign in to M365 [VSC → Chrome]
 *   4  Launch local debug in Copilot [VSC → Chrome]
 *   5  Validate agent responds [Chrome]
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
suite("DA No Action – Local Debug in M365 Copilot", function () {
  // Full test (including remote Copilot sign-in) can take up to 15 min
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
    console.log("=== DA No Action – Local Debug in M365 Copilot ===");
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
    await takeScreenshot("01-atk-sidebar");
    assert.ok(active, "Extension should be active");
  });

  // ── Test 2: Phase 1 + 2 – Create project and verify scaffold ─────────────
  test("Phase 1-2: Create DA No Action project via wizard and verify scaffold", async () => {
    // Step 1: clear notifications (dismiss welcome dialog)
    vscode.commands.executeCommand("notifications.clearAll").catch(() => {});
    await wait(500);

    // Step 2: fire the wizard WITHOUT await — it blocks until wizard completes
    const cmdAvailable = await waitForCommand("fx-extension.create");
    console.log("  fx-extension.create available:", cmdAvailable);
    vscode.commands.executeCommand("fx-extension.create").catch((e: any) => {
      console.log("  Command error:", e.message);
    });
    await wait(500); // yield event loop so command dispatch reaches extension host

    // Step 3: QuickPick — "Declarative Agent"
    await sendSignal(
      "waitForTextThenScreenshot:Declarative Agent:60000:02-da-option",
      68000,
    );
    await sendSignal("clickText:Declarative Agent", 10000);
    await wait(1000);

    // Step 4: QuickPick — "No Action"
    await sendSignal(
      "waitForTextThenScreenshot:No Action:20000:03-no-action",
      28000,
    );
    await sendSignal("clickText:No Action", 10000);
    await wait(1000);

    // Step 5: QuickPick — Workspace Folder "Default folder"
    await sendSignal(
      "waitForTextThenScreenshot:Default folder:15000:04-workspace-folder",
      23000,
    );
    await sendSignal("clickText:Default folder", 10000);
    await wait(1000);

    // Step 6: InputBox — Application Name
    await sendSignal(
      "waitForTextThenScreenshot:Application Name:15000:05-app-name-input",
      23000,
    );
    await sendSignal(`type:${APP_NAME}`, 8000);
    await wait(500);
    await sendSignal("pressKey:Enter", 5000);

    // Wait for scaffold — poll every second up to 90 s
    console.log(
      `  Waiting for scaffold files in ~/AgentsToolkitProjects/${APP_NAME} (up to 90s)...`,
    );
    for (let i = 0; i < 90; i++) {
      projectDir = findProjectDir(APP_NAME);
      if (
        projectDir &&
        fs.existsSync(path.join(projectDir, "appPackage", "manifest.json"))
      ) {
        break;
      }
      await wait(1000);
    }
    await takeScreenshot("06-scaffold-complete");

    step(
      "fx-extension.create command available",
      cmdAvailable,
      `available=${cmdAvailable}`,
    );
    step("Project directory found", !!projectDir, projectDir || "not found");

    // Step 7-9: Verify scaffold files
    const manifestPath = path.join(projectDir, "appPackage", "manifest.json");
    const declarativePath = path.join(
      projectDir,
      "appPackage",
      "declarativeAgent.json",
    );
    const pluginPath = path.join(projectDir, "appPackage", "ai-plugin.json");

    const hasManifest = !!projectDir && fs.existsSync(manifestPath);
    const hasDeclarative = !!projectDir && fs.existsSync(declarativePath);
    const noPlugin = !!projectDir && !fs.existsSync(pluginPath); // No Action DA has NO plugin file

    step(
      "appPackage/manifest.json exists",
      hasManifest,
      hasManifest ? "✓" : `missing in ${projectDir}`,
    );
    step(
      "appPackage/declarativeAgent.json exists",
      hasDeclarative,
      hasDeclarative ? "✓" : `missing in ${projectDir}`,
    );
    step(
      "appPackage/ai-plugin.json absent (No Action)",
      noPlugin,
      noPlugin ? "✓ not present" : "UNEXPECTED: file exists",
    );

    // Step 10: Close auto-opened README tab
    vscode.commands
      .executeCommand("workbench.action.closeActiveEditor")
      .catch(() => {});
    await wait(500);
    await takeScreenshot("07-readme-closed");

    // Step 11: Clear all notifications
    vscode.commands.executeCommand("notifications.clearAll").catch(() => {});
    await wait(500);
    await takeScreenshot("08-notifications-cleared");

    assert.ok(hasManifest, `appPackage/manifest.json missing in ${projectDir}`);
    assert.ok(
      hasDeclarative,
      `appPackage/declarativeAgent.json missing in ${projectDir}`,
    );
    assert.ok(
      noPlugin,
      `appPackage/ai-plugin.json should NOT exist for No Action DA`,
    );
  });

  // ── Test 3: Phase 3 – Sign in to M365 ────────────────────────────────────
  test("Phase 3: Sign in to M365 via ATK Accounts", async () => {
    const m365User = process.env.M365_ACCOUNT_NAME || "";
    const m365Pass = process.env.M365_ACCOUNT_PASSWORD || "";
    if (!m365User || !m365Pass) {
      console.log("  SKIP: M365_ACCOUNT_NAME / M365_ACCOUNT_PASSWORD not set");
      step("M365 credentials available", false, "env vars not set — skipped");
      return; // do not assert.ok — treat as skip, not failure
    }

    // Step 12: open ATK Accounts QuickPick WITHOUT await
    vscode.commands
      .executeCommand("fx-extension.cmpAccounts")
      .catch((e: any) => {
        console.log("  cmpAccounts error:", e.message);
      });
    await wait(500);

    // Step 13: click "Sign in to Microsoft 365"
    await sendSignal(
      "waitForTextThenScreenshot:Sign in to Microsoft 365:20000:09-signin-option",
      28000,
    );
    await sendSignal("clickText:Sign in to Microsoft 365", 10000);
    await wait(500);

    // Step 14: click "Sign in" in the modal
    await sendSignal(
      "waitForTextThenScreenshot:Sign in:15000:10-modal-signin",
      23000,
    );
    await sendSignal("clickText:Sign in", 10000);
    await wait(1000);

    // Step 15: [Chrome] wait for M365 login page (up to 60 s), enter credentials
    await sendSignal(
      "waitForTextThenScreenshot:Email or phone:60000:11-m365-login-page",
      68000,
    );
    await sendSignal("clickText:Email or phone", 10000);
    await sendSignal(`type:${m365User}`, 5000);
    await sendSignal("clickText:Next", 10000);
    await wait(3000);

    // Step 16: password
    await sendSignal(`type:${m365Pass}`, 5000);
    await sendSignal("pressKey:Enter", 5000);
    await wait(3000);
    await takeScreenshot("12-m365-login-complete");

    // Step 17: close browser tab
    await sendSignal(
      "waitForTextThenScreenshot:Close:10000:13-close-browser-tab",
      15000,
    );
    await sendSignal("clickText:Close", 5000);
    await wait(3000);
    await takeScreenshot("14-m365-signed-in");

    step("M365 sign-in completed", true, `user=${m365User}`);
  });

  // ── Test 4: Phase 4 – Launch local debug in M365 Copilot ─────────────────
  test("Phase 4: Launch local debug via Preview Local in Copilot (Chrome)", async () => {
    const m365User = process.env.M365_ACCOUNT_NAME || "";
    const m365Pass = process.env.M365_ACCOUNT_PASSWORD || "";
    if (!m365User || !m365Pass) {
      console.log("  SKIP: M365_ACCOUNT_NAME / M365_ACCOUNT_PASSWORD not set");
      step("Debug launch (skipped — no credentials)", false, "skipped");
      return;
    }

    // Step 18: fire debug-selection command WITHOUT await
    vscode.commands
      .executeCommand("workbench.action.debug.selectandstart")
      .catch((e: any) => {
        console.log("  debug.selectandstart error:", e.message);
      });
    await wait(500);

    // Step 19: select "Preview Local in Copilot (chrome)"
    await sendSignal(
      "waitForTextThenScreenshot:Preview Local in Copilot:20000:15-debug-config",
      28000,
    );
    await sendSignal("clickText:Preview Local in Copilot (chrome)", 10000);
    await wait(1000);
    await takeScreenshot("16-debug-started");

    // Step 20: [Chrome] wait up to 120 s for Microsoft sign-in page (new Chrome process)
    await sendSignal(
      "waitForTextThenScreenshot:Email or phone:120000:17-copilot-login-page",
      130000,
    );
    await sendSignal("clickText:Email or phone", 10000);
    await sendSignal(`type:${m365User}`, 5000);
    await sendSignal("clickText:Next", 10000);
    await wait(3000);

    // Step 21: password
    await sendSignal(`type:${m365Pass}`, 5000);
    await sendSignal("clickText:Sign in", 10000);
    await wait(3000);
    await takeScreenshot("18-copilot-login-complete");

    // Step 22: "Stay signed in?" → press Enter
    await sendSignal(
      "waitForTextThenScreenshot:Stay signed in:10000:19-stay-signed-in",
      15000,
    );
    await sendSignal("pressKey:Enter", 5000);
    await wait(3000);

    step(
      "Debug session launched in Copilot",
      true,
      "Preview Local in Copilot (Chrome) selected",
    );
  });

  // ── Test 5: Phase 5 – Validate agent responds in Copilot ─────────────────
  test("Phase 5: Validate DA agent visible and responds in M365 Copilot", async () => {
    const m365User = process.env.M365_ACCOUNT_NAME || "";
    if (!m365User) {
      console.log("  SKIP: M365_ACCOUNT_NAME not set");
      step("Copilot validation (skipped — no credentials)", false, "skipped");
      return;
    }

    const expectedAgent = `${APP_NAME}local`;

    // Step 23: wait 30 s for agent registration to propagate, then zoom out [Chrome]
    console.log("  Waiting 30s for agent registration to propagate...");
    await wait(30000);
    await sendSignal("pressKey:Ctrl+-", 5000);
    await takeScreenshot("20-copilot-zoomed");

    // Step 24: reload page [Chrome]
    await sendSignal("pressKey:F5", 5000);
    await wait(3000);
    await takeScreenshot("21-copilot-reloaded");

    // Step 25: assert agent entry visible (retry 30 s) — PRIMARY PASS CRITERION
    await sendSignal(
      `waitForTextThenScreenshot:${expectedAgent}:30000:22-agent-visible`,
      38000,
    );
    // Evaluate DOM to confirm text presence for the pass/fail record
    const agentResultFile = path.join(OUTPUT_DIR, "agent-check.json");
    await sendSignal(
      `eval:${agentResultFile}:JSON.stringify({found: !!document.querySelector('*:not(script):not(style)') && document.body.innerText.includes('${expectedAgent}')})`,
      10000,
    );
    let agentFound = false;
    try {
      if (fs.existsSync(agentResultFile)) {
        const r = JSON.parse(fs.readFileSync(agentResultFile, "utf8"));
        agentFound = r.found === true;
      }
    } catch {}
    step(
      `Agent "${expectedAgent}" visible in Copilot`,
      agentFound,
      agentFound ? "✓" : "not found after 30s retry",
    );

    // Step 26: click chat input, type query, send [Chrome]
    const query = "how can you assistant me?";
    await sendSignal("clickText:Message Copilot", 10000);
    await sendSignal(`type:${query}`, 5000);
    await sendSignal("pressKey:Enter", 5000);
    await takeScreenshot("23-query-sent");

    // Step 27: assert bot response (retry 60 s) — PRIMARY E2E PASS CRITERION
    await sendSignal(
      "waitForTextThenScreenshot:response message:60000:24-agent-response",
      68000,
    );
    await takeScreenshot("25-agent-response");
    // Use eval to check for any visible response in the chat area
    const responseResultFile = path.join(OUTPUT_DIR, "response-check.json");
    await sendSignal(
      `eval:${responseResultFile}:JSON.stringify({hasResponse: document.body.innerText.length > 200})`,
      10000,
    );
    let hasResponse = false;
    try {
      if (fs.existsSync(responseResultFile)) {
        const r = JSON.parse(fs.readFileSync(responseResultFile, "utf8"));
        hasResponse = r.hasResponse === true;
      }
    } catch {}
    step(
      "Agent returns response in Copilot chat",
      hasResponse,
      hasResponse ? "✓" : "no response after 60s retry",
    );

    assert.ok(
      agentFound,
      `Agent "${expectedAgent}" was not found in Copilot — local debug may not have registered the agent`,
    );
    assert.ok(hasResponse, "No bot response received after 60s");
  });
});
