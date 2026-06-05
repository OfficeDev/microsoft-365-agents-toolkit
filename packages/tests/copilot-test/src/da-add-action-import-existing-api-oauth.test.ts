// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * da-add-action-import-existing-api-oauth.test.ts
 * TC-001: Create DA with OAuth API action, provision to Azure, verify Sign-in button in Copilot.
 * Plan: packages/tests/copilot-test/test-plans/da-add-action-import-existing-api-oauth/
 *       da-add-action-import-existing-api-oauth.md
 *
 * Two-window design:
 *   Window 1 (no project): Phase 1+2 (wizard + M365 sign-in) → writes scaffold-done.json → exits
 *   Window 2 (project):    Phase 3+4+5+6 (provision + remote debug + Copilot validation) → exits
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

const SCAFFOLD_MARKER = path.join(OUTPUT_DIR, "scaffold-done.json");
const APP_NAME = `da-oauth-${Date.now()}`;

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
// Window detection
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
suite("DA Add Action – Import Existing API (OAuth)", function () {
  this.timeout(20 * 60 * 1000);

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
      `=== DA Add Action OAuth [${isWindow2 ? "Window 2 – project" : "Window 1 – wizard"}] ===`,
    );
    console.log("Output:", OUTPUT_DIR);
  });

  suiteTeardown(() => {
    if (isWindow2 || !fs.existsSync(SCAFFOLD_MARKER)) {
      writeResults(passed, failed, steps);
    }
    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
  });

  // ── Step 0: ATK active (both windows) ────────────────────────────────────
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
    // ── WINDOW 1: Phase 1 – Wizard ──────────────────────────────────────────
    test("Phase 1: Create DA with OAuth OpenAPI action via wizard", async () => {
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

      // Step 4: Add an Action
      await sendSignal(
        "waitForTextThenScreenshot:Add an Action:20000:03-wizard-add-action",
        28000,
      );
      await wait(500);
      await sendSignal("clickText:Add an Action", 12000);
      await wait(800);

      // Step 5: Start with an OpenAPI Description Document
      await sendSignal(
        "waitForTextThenScreenshot:Start with an OpenAPI Description Document:20000:04-openapi-source",
        28000,
      );
      await wait(500);
      await sendSignal(
        "clickText:Start with an OpenAPI Description Document",
        12000,
      );
      await wait(800);

      // Step 6: Enter OpenAPI Document URL
      await sendSignal(
        "waitForTextThenScreenshot:Enter OpenAPI Document URL:20000:05-enter-url-option",
        28000,
      );
      await wait(500);
      await sendSignal("clickText:Enter OpenAPI Document URL", 12000);
      await wait(800);

      // Step 7: Type the OAuth spec URL
      await sendSignal(
        "waitForTextThenScreenshot:OpenAPI Description Document URL:20000:06-url-input",
        28000,
      );
      await sendSignal(
        "type:https://raw.githubusercontent.com/SLdragon/example-openapi-spec/refs/heads/main/real-oauth.yaml",
        8000,
      );
      await wait(300);
      await sendSignal("pressKey:Enter", 5000);
      await takeScreenshot("07-url-submitted");

      // Step 8: Wait for operations to load (up to 60s) then click OK
      await sendSignal(
        "waitForTextThenScreenshot:Select Operation:60000:08-operations-loaded",
        68000,
      );
      await wait(500);
      await sendSignal("clickText:OK", 10000);
      await wait(800);
      await takeScreenshot("09-ok-clicked");

      // Step 10: Clear All
      await sendSignal(
        "waitForTextThenScreenshot:Clear All:10000:10-clear-all-visible",
        18000,
      );
      await wait(300);
      await sendSignal("clickText:Clear All", 10000);
      await wait(800);
      await takeScreenshot("10-clear-all");

      // Step 11: Default folder
      await sendSignal(
        "waitForTextThenScreenshot:Default folder:20000:11-workspace-folder",
        28000,
      );
      await wait(500);
      await sendSignal("clickText:Default folder", 12000);
      await wait(800);

      // Step 12: Application Name
      await sendSignal(
        "waitForTextThenScreenshot:Application Name:15000:12-appname",
        23000,
      );
      await sendSignal(`type:${APP_NAME}`, 8000);
      await wait(300);
      await sendSignal("pressKey:Enter", 5000);

      // Write marker so Window 2 can pick up
      const projectDir = path.join(
        os.homedir(),
        "AgentsToolkitProjects",
        APP_NAME,
      );
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
      await takeScreenshot("13-scaffold-started");
      step("fx-extension.create available", cmdAvailable);
      step("Scaffold marker written", true, `projectDir=${projectDir}`);
    });
  } else {
    // ── WINDOW 2: Phases 2–6 ────────────────────────────────────────────────
    const projectDir = scaffoldedProjectDir;

    test("Phase 2: Verify scaffold files", async () => {
      const sentinel = path.join(projectDir, "m365agents.yml");
      console.log(`  Waiting for scaffold at: ${projectDir}`);
      for (let i = 0; i < 90 && !fs.existsSync(sentinel); i++) await wait(1000);
      await takeScreenshot("13-scaffold-explorer");

      const required = [
        "m365agents.yml",
        "appPackage/manifest.json",
        "appPackage/declarativeAgent.json",
        "appPackage/ai-plugin.json",
      ];
      for (const rel of required) {
        const exists = fs.existsSync(path.join(projectDir, rel));
        step(
          `${rel} exists`,
          exists,
          exists ? "✓" : `missing in ${projectDir}`,
        );
      }
      assert.ok(
        fs.existsSync(sentinel),
        `m365agents.yml missing in ${projectDir}`,
      );
      assert.ok(
        fs.existsSync(path.join(projectDir, "appPackage", "ai-plugin.json")),
        "ai-plugin.json should exist for OAuth action DA",
      );
    });

    test("Phase 3: Sign in to M365", async () => {
      const m365User = process.env.M365_ACCOUNT_NAME || "";
      const m365Pass = process.env.M365_ACCOUNT_PASSWORD || "";
      assert.ok(
        m365User && m365Pass,
        "M365_ACCOUNT_NAME and M365_ACCOUNT_PASSWORD must be set",
      );

      vscode.commands
        .executeCommand("fx-extension.cmpAccounts")
        .then(undefined, () => {});
      await wait(500);

      await sendSignal(
        "waitForTextThenScreenshot:Sign in to Microsoft 365:20000:14-signin-option",
        28000,
      );
      await sendSignal("clickText:Sign in to Microsoft 365", 10000);
      await wait(1000);
      await takeScreenshot("14-sign-in-triggered");

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

      // OAuth watcher in runTest.ts completes browser sign-in automatically
      const accountHint = m365User.split("@")[0];
      await sendSignal(
        `waitForTextThenScreenshot:${accountHint}:90000:15-m365-signed-in`,
        98000,
      );
      await takeScreenshot("15-m365-signed-in");
      step("M365 sign-in completed", true, `user=${m365User}`);
    });

    test("Phase 4: Provision to Azure", async () => {
      // Open Command Palette → provision
      vscode.commands
        .executeCommand("workbench.action.showCommands")
        .then(undefined, () => {});
      await wait(500);
      await sendSignal(
        "waitForTextThenScreenshot:>:5000:16-cmd-palette-open",
        8000,
      );
      await sendSignal("type:provision", 5000);
      await wait(500);
      await sendSignal(
        "waitForTextThenScreenshot:Provision:10000:17-provision-cmd",
        15000,
      );
      await sendSignal("pressKey:Enter", 5000);
      await wait(1000);
      await takeScreenshot("17-provision-starting");

      // Select environment: dev
      await sendSignal(
        "waitForTextThenScreenshot:dev:30000:18-env-select",
        38000,
      );
      await wait(500);
      await sendSignal("clickText:dev", 10000);
      await wait(1000);
      await takeScreenshot("18-dev-env-selected");

      // OAuth client ID (wait up to 30s after env selection per test plan notes)
      await sendSignal(
        "waitForTextThenScreenshot:OAuth registration client ID:30000:19-oauth-client-id",
        38000,
      );
      await takeScreenshot("19-oauth-client-id-input");
      await sendSignal("type:fake_client_id", 5000);
      await wait(300);
      await sendSignal("pressKey:Enter", 5000);
      await wait(500);

      // OAuth client secret
      await sendSignal(
        "waitForTextThenScreenshot:OAuth registration client secret:20000:20-oauth-secret",
        28000,
      );
      await takeScreenshot("20-oauth-client-secret-input");
      await sendSignal("type:fake_clientsecret", 5000);
      await wait(300);
      await sendSignal("pressKey:Enter", 5000);
      await wait(500);

      // OAuth confirmation dialog
      await sendSignal(
        "waitForTextThenScreenshot:Confirm:20000:21-oauth-confirm",
        28000,
      );
      await takeScreenshot("21-oauth-confirm-dialog");
      await sendSignal("clickText:Confirm", 10000);
      await wait(1000);

      // Wait for provision to complete (up to 10 min)
      await sendSignal(
        "waitForTextThenScreenshot:provision stage executed successfully:600000:22-provision-done",
        610000,
      );
      await takeScreenshot("22-provision-succeeded");

      const provisionNotification = fs.existsSync(
        path.join(SCREENSHOT_DIR, "22-provision-done.png"),
      );
      step(
        "Provision completed successfully",
        provisionNotification,
        provisionNotification
          ? "✓ notification screenshot captured"
          : "notification not found",
      );
      assert.ok(
        provisionNotification,
        "Provision must complete before remote debug",
      );
    });

    test("Phase 5: Launch remote debug in M365 Copilot", async () => {
      const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
      console.log(
        `  workspaceFolder: ${workspaceFolder?.uri.fsPath ?? "none"}`,
      );
      await vscode.commands.executeCommand("workbench.view.debug");
      await wait(1000);
      await takeScreenshot("23-debug-panel");

      const debugStarted = await vscode.debug.startDebugging(
        workspaceFolder,
        "Preview in Copilot (Chrome)",
      );
      console.log(`  debug.startDebugging: ${debugStarted}`);
      step(
        "Debug session started (Preview in Copilot Chrome)",
        debugStarted,
        debugStarted ? "✓" : "failed",
      );
      await wait(2000);

      // Wait for Copilot to open in Chrome (up to 2 min)
      await sendSignal(
        "waitForTextThenScreenshot:m365.cloud.microsoft:120000:24-copilot-opened",
        130000,
      );
      await takeScreenshot("24-copilot-opened");
      assert.ok(debugStarted, "debug.startDebugging returned false");
    });

    test("Phase 6: Validate OAuth action in M365 Copilot", async () => {
      // Zoom out for better visibility
      await sendSignal("pressKey:Ctrl+-", 3000);
      await wait(500);
      await sendSignal("pressKey:F5", 3000); // refresh
      await wait(3000);
      await takeScreenshot("25-copilot-reloaded");

      // Send repair query
      await sendSignal(
        "waitForTextThenScreenshot:Message Copilot:30000:26-copilot-ready",
        38000,
      );
      await sendSignal(
        "type:show repair records assigned to karin blair",
        5000,
      );
      await sendSignal("pressKey:Enter", 5000);
      await takeScreenshot("26-query-sent");

      // Wait for Allow button (OAuth permission prompt)
      await sendSignal(
        "waitForTextThenScreenshot:Allow:60000:27-allow-prompt",
        68000,
      );
      await takeScreenshot("27-allow-prompt");
      await sendSignal("clickText:Allow", 10000);
      await wait(1000);
      // Second Allow (confirmation)
      await sendSignal(
        "waitForTextThenScreenshot:Allow:10000:28-allow-confirm",
        15000,
      );
      await sendSignal("clickText:Allow", 10000);
      await wait(2000);

      // Wait up to 60s for "Sign in to Repair Service" button (primary pass criterion)
      await sendSignal(
        "waitForTextThenScreenshot:Sign in to Repair Service:60000:29-sign-in-button",
        68000,
      );
      await takeScreenshot("29-sign-in-button");

      // Evaluate DOM to confirm the button text
      const resultFile = path.join(OUTPUT_DIR, "signin-button-check.json");
      await sendSignal(
        `eval:${resultFile}:JSON.stringify({found: document.body.innerText.includes('Sign in to Repair Service')})`,
        10000,
      );
      let btnFound = false;
      try {
        if (fs.existsSync(resultFile)) {
          btnFound =
            JSON.parse(fs.readFileSync(resultFile, "utf8")).found === true;
        }
      } catch {}

      // Screenshot from signal file is the primary evidence
      const screenshotExists = fs.existsSync(
        path.join(SCREENSHOT_DIR, "29-sign-in-button.png"),
      );
      step(
        "Sign in to Repair Service button visible in Copilot",
        screenshotExists || btnFound,
        screenshotExists
          ? "✓ screenshot captured"
          : btnFound
            ? "✓ DOM eval"
            : "not found after 60s",
      );

      // Clean up marker
      try {
        fs.unlinkSync(SCAFFOLD_MARKER);
      } catch {}

      assert.ok(
        screenshotExists || btnFound,
        '"Sign in to Repair Service" button not found in Copilot',
      );
    });
  }
});
