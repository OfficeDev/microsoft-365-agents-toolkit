// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * runTest.ts - Hybrid: @vscode/test-electron (activates extension + Mocha)
 *              + Playwright CDP connection (screenshots + UI interaction)
 *
 * Extended signal protocol:
 *   Signal file name: {timestamp}-{name}.signal
 *   Signal file content:
 *     screenshot:{destPath}    — take a screenshot
 *     click:{cssSelector}      — click matching element in QuickPick
 *     clickText:{text}         — click QuickPick item whose label contains text
 *     type:{text}              — type text into active input
 *     pressKey:{key}           — press a keyboard key (Enter, Escape, Tab, ...)
 *     waitForText:{text}       — wait until element with text is visible, then delete signal
 */
import * as path from "path";
import * as fs from "fs";
import * as cp from "child_process";
import * as os from "os";
import { downloadAndUnzipVSCode, runTests } from "@vscode/test-electron";
import { chromium } from "playwright";
import type { Browser, Page } from "playwright";

const HERE = __dirname;
const TESTS_ROOT = path.resolve(HERE, "..");

function sleep(ms: number) {
  return new Promise<void>((r) => setTimeout(r, ms));
}

// ── Signal watcher (screenshots + UI actions) ─────────────────────────────────
async function startSignalWatcher(
  signalDir: string,
  getPage: () => Page | null,
  stopFlag: { stop: boolean },
) {
  fs.mkdirSync(signalDir, { recursive: true });
  console.log("📡 Signal watcher started →", signalDir);

  while (!stopFlag.stop) {
    const signals = fs
      .readdirSync(signalDir)
      .filter((f) => f.endsWith(".signal"))
      .sort();

    for (const sig of signals) {
      const sigPath = path.join(signalDir, sig);
      let content = "";
      try {
        content = fs.readFileSync(sigPath, "utf8").trim();
      } catch {
        continue; // already deleted
      }

      const page = getPage();
      try {
        if (content.startsWith("screenshot:")) {
          const dest = content.slice("screenshot:".length);
          if (page) {
            await page.screenshot({ path: dest, fullPage: false });
            console.log(`  📸 ${path.basename(dest)}`);
          }
        } else if (content.startsWith("clickText:")) {
          const text = content.slice("clickText:".length);
          if (page) {
            // QuickPick item label selector
            const item = page
              .locator(".quick-input-list .monaco-list-row")
              .filter({ hasText: text });
            try {
              await item.first().waitFor({ timeout: 8000 });
              await item.first().click();
              console.log(`  🖱️ Clicked: "${text}"`);
              await sleep(500); // settle
            } catch {
              console.warn(
                `  ⚠️ clickText: "${text}" not found, trying pressKey ArrowDown+Enter`,
              );
              await page.keyboard.press("ArrowDown");
              await sleep(200);
              await page.keyboard.press("Enter");
            }
          }
        } else if (content.startsWith("click:")) {
          const selector = content.slice("click:".length);
          if (page) {
            try {
              await page.locator(selector).first().waitFor({ timeout: 8000 });
              await page.locator(selector).first().click();
              console.log(`  🖱️ Clicked selector: ${selector}`);
              await sleep(300);
            } catch (e) {
              console.warn(`  ⚠️ click: selector "${selector}" not found`);
            }
          }
        } else if (content.startsWith("type:")) {
          const text = content.slice("type:".length);
          if (page) {
            // Use fill() to set/clear input - avoids Ctrl+a shortcuts that can trigger QuickInput actions
            const input = page.locator(
              ".quick-input-box input, .quick-input-filter .input",
            );
            try {
              await input.first().waitFor({ timeout: 5000 });
              await input.first().fill(text); // fill() clears then sets without keyboard shortcuts
              console.log(`  ⌨️ Typed: "${text}"`);
              await sleep(300);
            } catch (e) {
              // Fallback: keyboard type
              await page.keyboard.type(text, { delay: 30 });
            }
          }
        } else if (content.startsWith("pressKey:")) {
          const key = content.slice("pressKey:".length);
          if (page) {
            await page.keyboard.press(key);
            console.log(`  ⌨️ Key: ${key}`);
            await sleep(300);
          }
        } else if (content.startsWith("waitForText:")) {
          const text = content.slice("waitForText:".length);
          if (page) {
            try {
              await page.waitForSelector(`text="${text}"`, { timeout: 10000 });
              console.log(`  ✅ Found text: "${text}"`);
            } catch {
              console.warn(
                `  ⚠️ waitForText: "${text}" not found within timeout`,
              );
            }
          }
        } else {
          // Legacy: treat content as a screenshot dest path
          if (page && content) {
            await page.screenshot({ path: content, fullPage: false });
            console.log(`  📸 (legacy) ${path.basename(content)}`);
          }
        }
      } catch (e) {
        console.warn(`  Signal error (${sig}):`, e);
      }

      try {
        fs.rmSync(sigPath, { force: true });
      } catch {
        /* ok */
      }
    }

    await sleep(150);
  }
}

// ── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  const extPath = process.env.ATK_EXT_PATH ?? "";

  const outputDir =
    process.env.TEST_OUTPUT_DIR ||
    path.resolve(TESTS_ROOT, "../../test-output");

  const screenshotDir = path.join(outputDir, "screenshots");
  const signalDir = path.join(outputDir, ".screenshot-signals");
  const userDataDir = path.join(outputDir, "vscode-user-data");

  for (const d of [screenshotDir, signalDir, userDataDir]) {
    fs.mkdirSync(d, { recursive: true });
  }

  // Clean old signals
  fs.readdirSync(signalDir).forEach((f) =>
    fs.rmSync(path.join(signalDir, f), { force: true }),
  );

  console.log("=== Playwright + test-electron Hybrid Runner ===");
  console.log("Ext:", extPath);
  console.log("Out:", outputDir);

  if (extPath && !fs.existsSync(extPath)) {
    console.error("ATK extension not found:", extPath);
    process.exit(1);
  }

  // Compile the Mocha suite
  const tmpOut = path.join(TESTS_ROOT, "out");
  fs.mkdirSync(tmpOut, { recursive: true });
  const tsconfigPath = path.join(HERE, "_tsconfig.build.json");
  const tsconfig = {
    compilerOptions: {
      module: "commonjs",
      target: "ES2020",
      lib: ["ES2020"],
      esModuleInterop: true,
      resolveJsonModule: true,
      strict: false,
      skipLibCheck: true,
      outDir: tmpOut,
      rootDir: HERE,
      types: ["node", "mocha"],
      typeRoots: [path.join(TESTS_ROOT, "node_modules", "@types")],
    },
    include: ["suite/**/*.ts", "*.test.ts"],
    exclude: ["node_modules", "runTest.ts"],
  };
  fs.writeFileSync(tsconfigPath, JSON.stringify(tsconfig, null, 2), "utf8");
  const tscBin = path.join(
    TESTS_ROOT,
    "node_modules",
    ".bin",
    process.platform === "win32" ? "tsc.CMD" : "tsc",
  );
  const compileResult = cp.spawnSync(tscBin, ["--project", tsconfigPath], {
    cwd: HERE,
    stdio: "inherit",
    shell: process.platform === "win32",
  });
  fs.rmSync(tsconfigPath, { force: true });
  if (compileResult.status !== 0) console.warn("tsc warnings – continuing");

  const extensionTestsPath = path.join(tmpOut, "suite", "index");
  if (!fs.existsSync(extensionTestsPath + ".js")) {
    console.error("Compiled suite missing:", extensionTestsPath + ".js");
    process.exit(1);
  }

  const CDP_PORT = 9229;

  let activePage: Page | null = null;
  const stopFlag = { stop: false };

  const watcherPromise = startSignalWatcher(
    signalDir,
    () => activePage,
    stopFlag,
  );

  const userExtDir =
    process.env.VSCODE_EXTENSIONS_DIR ||
    path.join(os.homedir(), ".vscode", "extensions");

  const yamlExtPresent =
    fs.existsSync(
      path.join(userExtDir, "redhat.vscode-yaml", "package.json"),
    ) ||
    fs
      .readdirSync(userExtDir)
      .some(
        (d) =>
          d.startsWith("redhat.vscode-yaml") &&
          fs.existsSync(path.join(userExtDir, d, "package.json")),
      );
  console.log(`Extensions dir: ${userExtDir} (yaml: ${yamlExtPresent})`);

  const vscodeTestOpts: any = {
    extensionTestsPath,
    launchArgs: [
      "--disable-workspace-trust",
      "--skip-welcome",
      "--skip-release-notes",
      `--user-data-dir=${userDataDir}`,
      "--no-sandbox",
      `--remote-debugging-port=${CDP_PORT}`,
      `--extensions-dir=${userExtDir}`,
      ...(yamlExtPresent ? [] : ["--install-extension", "redhat.vscode-yaml"]),
    ],
    version: "stable",
    extensionTestsEnv: {
      TEST_OUTPUT_DIR: outputDir,
      SCREENSHOT_SIGNAL_DIR: signalDir,
      SCREENSHOT_DIR: screenshotDir,
    },
  };
  if (extPath) {
    vscodeTestOpts.extensionDevelopmentPath = extPath;
  }
  const testRunPromise = runTests(vscodeTestOpts);

  let browser: Browser | null = null;
  let cdpConnected = false;
  const cdpStart = Date.now();

  while (!cdpConnected && Date.now() - cdpStart < 20000) {
    await sleep(1000);
    try {
      browser = await chromium.connectOverCDP(`http://localhost:${CDP_PORT}`, {
        timeout: 2000,
      });
      cdpConnected = true;
    } catch (_) {}
  }

  if (browser) {
    await sleep(1000);
    const contexts = browser.contexts();
    console.log(`CDP connected: ${contexts.length} context(s)`);
    const allPages = contexts.flatMap((c) => c.pages());
    console.log(`  Pages: ${allPages.length}`);

    activePage =
      allPages.find((p) => p.url().includes("vscode-app")) ??
      allPages.sort((a, b) => b.url().length - a.url().length)[0] ??
      null;

    if (activePage) {
      const title = await activePage.title().catch(() => "?");
      console.log(`  Active page: ${title}`);

      contexts.forEach((ctx) => {
        ctx.on("page", (newPage) => {
          console.log("  New page:", newPage.url());
          // Stay on main window for QuickPick (QuickPick is in the main window, not a separate page)
          // Only switch to new page if it's a webview panel
          if (newPage.url().includes("webview-panel")) {
            activePage = newPage;
          }
        });
      });
    }
  } else {
    console.warn("CDP connect failed - UI interaction will be skipped");
  }

  try {
    await testRunPromise;
    console.log("Test run completed successfully");
  } catch (e: any) {
    console.error("Test run failed:", e.message);
  } finally {
    stopFlag.stop = true;
    await sleep(300);
    if (browser) await browser.close().catch(() => {});
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
