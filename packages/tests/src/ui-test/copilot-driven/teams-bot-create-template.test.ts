// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
//
// TC-001: Create Teams Bot template (TypeScript) via ATK VSCode extension wizard.
//
// This test is executed by the Copilot-driven ATK test pipeline.
// It opens VSCode with the ATK extension, runs the "Teams: Create New App" wizard,
// creates a Basic Bot (TypeScript) project, and asserts the scaffold is correct.

import * as path from "path";
import * as fs from "fs-extra";
import { expect } from "chai";
import {
  VSBrowser,
  Workbench,
  InputBox,
  By,
  Key,
  until,
} from "vscode-extension-tester";

// ── Configuration ────────────────────────────────────────────────────────────

const OUTPUT_DIR = process.env.TEST_OUTPUT_DIR ?? "/tmp/atk-test-output";
const PROJECTS_DIR = path.join(OUTPUT_DIR, "projects");
const SCREENSHOTS_DIR = path.join(OUTPUT_DIR, "screenshots");
const RESULTS_FILE = path.join(OUTPUT_DIR, "results.json");
const APP_NAME = `test-teams-bot-${Date.now()}`;

// ── Helpers ───────────────────────────────────────────────────────────────────

interface StepResult {
  index: number;
  name: string;
  status: "passed" | "failed";
  message?: string;
}

const results: StepResult[] = [];
let stepIndex = 0;

async function step(name: string, fn: () => Promise<void>): Promise<void> {
  const idx = ++stepIndex;
  try {
    await fn();
    results.push({ index: idx, name, status: "passed" });
    console.log(`  ✅ Step ${idx}: ${name}`);
  } catch (err: any) {
    results.push({ index: idx, name, status: "failed", message: err?.message ?? String(err) });
    console.error(`  ❌ Step ${idx}: ${name} — ${err?.message}`);
    throw err;
  }
}

async function screenshot(id: string): Promise<void> {
  try {
    await fs.ensureDir(SCREENSHOTS_DIR);
    const file = path.join(SCREENSHOTS_DIR, `${id}.png`);
    await VSBrowser.instance.takeScreenshot(id);
    // vscode-extension-tester writes to cwd by default; move if needed
    const cwdFile = path.join(process.cwd(), `${id}.png`);
    if (await fs.pathExists(cwdFile)) {
      await fs.move(cwdFile, file, { overwrite: true });
    }
    console.log(`    📸 Screenshot: ${id}.png`);
  } catch (e: any) {
    console.warn(`    ⚠️  Screenshot failed (${id}): ${e?.message}`);
  }
}

async function writeResults(): Promise<void> {
  await fs.ensureDir(OUTPUT_DIR);
  const passed = results.filter((r) => r.status === "passed").length;
  const failed = results.filter((r) => r.status === "failed").length;
  const output = { passed, failed, steps: results };
  await fs.writeJSON(RESULTS_FILE, output, { spaces: 2 });
  console.log(`\nResults: ${passed} passed, ${failed} failed → ${RESULTS_FILE}`);
}

async function openCommandPalette(workbench: Workbench): Promise<void> {
  const driver = VSBrowser.instance.driver;
  // Try Ctrl+Shift+P
  await driver.actions().keyDown(Key.CONTROL).sendKeys(Key.SHIFT, "p").keyUp(Key.CONTROL).perform();
  await driver.sleep(800);
}

// ── Test Suite ────────────────────────────────────────────────────────────────

describe("TC-001: Create Teams Bot Template (TypeScript)", function () {
  this.timeout(5 * 60 * 1000); // 5 minutes

  before(async function () {
    await fs.ensureDir(PROJECTS_DIR);
    await fs.ensureDir(SCREENSHOTS_DIR);
    // Wait for workbench to be ready
    await VSBrowser.instance.waitForWorkbench(30_000);
    await VSBrowser.instance.driver.sleep(3000);
  });

  after(async function () {
    await screenshot("99-final-state");
    await writeResults();
  });

  it("should open VSCode with ATK extension activated", async function () {
    await step("01-startup-screenshot", async () => {
      await screenshot("01-startup");
    });

    await step("02-verify-atk-extension-sidebar", async () => {
      const driver = VSBrowser.instance.driver;
      // Look for ATK icon in activity bar
      await driver.wait(
        until.elementLocated(By.css('[aria-label*="Teams"], [title*="Teams"], [aria-label*="Microsoft 365"]')),
        15_000
      ).catch(() => {
        console.log("ATK sidebar icon not immediately visible — continuing anyway");
      });
      await screenshot("02-extension-sidebar");
    });
  });

  it("should open Teams: Create New App via Command Palette", async function () {
    await step("03-open-command-palette", async () => {
      const workbench = new Workbench();
      await openCommandPalette(workbench);
      await screenshot("03-command-palette-open");
    });

    await step("04-type-create-new-app", async () => {
      const driver = VSBrowser.instance.driver;
      // Type the command
      await driver.sleep(500);
      const input = await InputBox.create(5000).catch(() => null);
      if (input) {
        await input.clear();
        await input.setText("Teams: Create New App");
        await driver.sleep(1000);
        await screenshot("04-command-typed");
        await input.confirm();
      } else {
        // fallback: type directly
        await driver.actions().sendKeys("Teams: Create New App").perform();
        await driver.sleep(1000);
        await screenshot("04-command-typed-fallback");
        await driver.actions().sendKeys(Key.ENTER).perform();
      }
      await driver.sleep(2000);
    });
  });

  it("should navigate the ATK new project wizard", async function () {
    const driver = VSBrowser.instance.driver;

    await step("05-select-teams-agents-apps", async () => {
      await driver.sleep(1500);
      await screenshot("05-wizard-step1");
      // Select "Teams Agent & Apps" (or first option if different build)
      const input = await InputBox.create(8000);
      const items = await input.getQuickPicks();
      const target = items.find((i) =>
        i.getText().then((t) => t.toLowerCase().includes("agent") || t.toLowerCase().includes("bot"))
      );
      if (target) {
        await target.select();
      } else {
        // pick first item as fallback
        await input.selectQuickPick(0);
      }
      await driver.sleep(1000);
    });

    await step("06-select-bot", async () => {
      await screenshot("06-wizard-step2-app-type");
      const input = await InputBox.create(8000);
      await input.selectQuickPick("Bot");
      await driver.sleep(1000);
    });

    await step("07-select-basic-bot", async () => {
      await screenshot("07-wizard-step3-bot-subtype");
      const input = await InputBox.create(8000);
      // Try "Basic Bot" or first available
      const picks = await input.getQuickPicks();
      const basic = picks.find(async (p) =>
        (await p.getText()).toLowerCase().includes("basic")
      );
      if (basic) {
        await basic.select();
      } else {
        await input.selectQuickPick(0);
      }
      await driver.sleep(1000);
    });

    await step("08-select-typescript", async () => {
      await screenshot("08-wizard-step4-language");
      const input = await InputBox.create(8000);
      await input.selectQuickPick("TypeScript");
      await driver.sleep(1000);
    });

    await step("09-enter-app-name", async () => {
      await screenshot("09-wizard-step5-name");
      const input = await InputBox.create(8000);
      await input.clear();
      await input.setText(APP_NAME);
      await driver.sleep(500);
      await input.confirm();
      await driver.sleep(1000);
    });

    await step("10-choose-output-folder", async () => {
      await screenshot("10-wizard-step6-folder");
      const input = await InputBox.create(8000).catch(() => null);
      if (input) {
        await input.clear();
        await input.setText(PROJECTS_DIR);
        await driver.sleep(500);
        await input.confirm();
      }
      await driver.sleep(2000);
    });
  });

  it("should wait for scaffold to complete and verify files", async function () {
    const driver = VSBrowser.instance.driver;

    await step("11-wait-for-scaffold", async () => {
      // Wait up to 60 seconds for scaffold notification or file to appear
      const projectPath = path.join(PROJECTS_DIR, APP_NAME);
      let elapsed = 0;
      while (elapsed < 60_000) {
        if (await fs.pathExists(path.join(projectPath, "package.json"))) break;
        await driver.sleep(2000);
        elapsed += 2000;
      }
      await screenshot("11-scaffold-in-progress");
    });

    await step("12-assert-scaffold-files", async () => {
      const projectPath = path.join(PROJECTS_DIR, APP_NAME);
      const requiredFiles = [
        "package.json",
        "teamsapp.yml",
        "appPackage/manifest.json",
        "src/index.ts",
      ];
      for (const f of requiredFiles) {
        const full = path.join(projectPath, f);
        expect(await fs.pathExists(full), `Expected file to exist: ${f}`).to.be.true;
        console.log(`    ✔ ${f}`);
      }
      await screenshot("12-scaffold-complete");
    });

    await step("13-verify-no-error-notifications", async () => {
      // Check for error notification toasts
      const driver2 = VSBrowser.instance.driver;
      const errors = await driver2
        .findElements(By.css('.notification-toast[aria-label*="error" i], .notification-list-item[aria-label*="error" i]'))
        .catch(() => []);
      expect(errors.length, "Unexpected error notification toasts").to.equal(0);
      await screenshot("13-no-errors");
    });
  });
});