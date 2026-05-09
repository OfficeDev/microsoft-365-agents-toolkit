// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * runTest.ts - Hybrid: @vscode/test-electron (activates extension + Mocha)
 *              + Playwright CDP connection (takes real screenshots externally)
 */
import * as path from "path";
import * as fs from "fs";
import * as cp from "child_process";
import * as os from "os";
import { downloadAndUnzipVSCode, runTests } from "@vscode/test-electron";
import { chromium } from "playwright";
import type { Browser, Page } from "playwright";

const HERE = __dirname;
const TESTS_ROOT = path.resolve(HERE, "../../..");

function sleep(ms: number) { return new Promise<void>((r) => setTimeout(r, ms)); }

// ── Screenshot signal watcher ─────────────────────────────────────────────────
async function startScreenshotWatcher(
  signalDir: string,
  getPage: () => Page | null,
  stopFlag: { stop: boolean }
) {
  fs.mkdirSync(signalDir, { recursive: true });
  console.log("📡 Screenshot watcher started →", signalDir);

  while (!stopFlag.stop) {
    const signals = fs.readdirSync(signalDir).filter((f) => f.endsWith(".signal"));
    for (const sig of signals) {
      const sigPath = path.join(signalDir, sig);
      try {
        const dest = fs.readFileSync(sigPath, "utf8").trim();
        const page = getPage();
        if (page) {
          await page.screenshot({ path: dest, fullPage: false });
          console.log(`  📸 ${path.basename(dest)}`);
        }
        fs.rmSync(sigPath, { force: true });
      } catch (e) {
        fs.rmSync(sigPath, { force: true });
      }
    }
    await sleep(150);
  }
}

// ── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  const extPath =
    process.env.ATK_EXT_PATH ||
    path.resolve(TESTS_ROOT, "../../packages/vscode-extension");

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
  fs.readdirSync(signalDir).forEach((f) => fs.rmSync(path.join(signalDir, f), { force: true }));

  console.log("=== Playwright + test-electron Hybrid Runner ===");
  console.log("Ext:", extPath);
  console.log("Out:", outputDir);

  if (!fs.existsSync(extPath)) {
    console.error("ATK extension not found:", extPath);
    process.exit(1);
  }

  // Compile the Mocha suite
  const tmpOut = path.join(TESTS_ROOT, "out", "copilot-driven");
  fs.mkdirSync(tmpOut, { recursive: true });
  const tsconfigPath = path.join(HERE, "_tsconfig.build.json");
  const tsconfig = {
    compilerOptions: {
      module: "commonjs", target: "ES2020", lib: ["ES2020"],
      esModuleInterop: true, resolveJsonModule: true,
      strict: false, skipLibCheck: true,
      outDir: tmpOut, rootDir: HERE,
      types: ["node", "mocha"],
      typeRoots: [path.join(TESTS_ROOT, "node_modules", "@types")],
    },
    include: ["suite/**/*.ts", "teams-bot-create-template.test.ts"],
    exclude: ["node_modules", "runTest.ts"],
  };
  fs.writeFileSync(tsconfigPath, JSON.stringify(tsconfig, null, 2), "utf8");
  const tscBin = path.join(TESTS_ROOT, "node_modules", ".bin",
    process.platform === "win32" ? "tsc.CMD" : "tsc");
  const compileResult = cp.spawnSync(tscBin, ["--project", tsconfigPath], {
    cwd: HERE, stdio: "inherit", shell: process.platform === "win32",
  });
  fs.rmSync(tsconfigPath, { force: true });
  if (compileResult.status !== 0) console.warn("tsc warnings – continuing");

  const extensionTestsPath = path.join(tmpOut, "suite", "index");
  if (!fs.existsSync(extensionTestsPath + ".js")) {
    console.error("Compiled suite missing:", extensionTestsPath + ".js");
    process.exit(1);
  }

  // CDP port for Playwright
  const CDP_PORT = 9229;

  // Track active Playwright page
  let activePage: Page | null = null;
  const stopFlag = { stop: false };

  // Start watcher BEFORE launching VSCode
  const watcherPromise = startScreenshotWatcher(signalDir, () => activePage, stopFlag);

  // Extensions dir: env override for Docker (VSCODE_EXTENSIONS_DIR) or user default
  const userExtDir = process.env.VSCODE_EXTENSIONS_DIR
    || path.join(os.homedir(), ".vscode", "extensions");

  // Check if yaml extension is present in the extensions dir
  const yamlExtPresent = fs.existsSync(path.join(userExtDir, "redhat.vscode-yaml", "package.json"))
    || fs.readdirSync(userExtDir).some(d => d.startsWith("redhat.vscode-yaml") && 
       fs.existsSync(path.join(userExtDir, d, "package.json")));
  console.log(`Extensions dir: ${userExtDir} (yaml: ${yamlExtPresent})`);

  // Launch VSCode via @vscode/test-electron
  const testRunPromise = runTests({
    extensionDevelopmentPath: extPath,
    extensionTestsPath,
    launchArgs: [
      "--disable-workspace-trust",
      "--skip-welcome",
      "--skip-release-notes",
      `--user-data-dir=${userDataDir}`,
      "--no-sandbox",
      `--remote-debugging-port=${CDP_PORT}`,
      // Include user extensions so redhat.vscode-yaml satisfies ATK's extensionDependencies
      `--extensions-dir=${userExtDir}`,
      // If yaml not in extensions dir, install it from marketplace (first run only)
      ...(yamlExtPresent ? [] : ["--install-extension", "redhat.vscode-yaml"]),
    ],
    version: "stable",
    extensionTestsEnv: {
      TEST_OUTPUT_DIR: outputDir,
      SCREENSHOT_SIGNAL_DIR: signalDir,
      SCREENSHOT_DIR: screenshotDir,
    },
  });

  // Connect CDP - poll until VSCode exposes the debugging endpoint
  let browser: Browser | null = null;
  let cdpConnected = false;
  const cdpStart = Date.now();

  // Poll for CDP availability (VSCode can take 8-15s to start)
  while (!cdpConnected && Date.now() - cdpStart < 20000) {
    await sleep(1000);
    try {
      browser = await chromium.connectOverCDP(`http://localhost:${CDP_PORT}`, { timeout: 2000 });
      cdpConnected = true;
    } catch (_) {
      // Not ready yet
    }
  }

  if (browser) {
    // Give VSCode UI a moment to settle
    await sleep(1000);
    const contexts = browser.contexts();
    console.log(`CDP connected: ${contexts.length} context(s)`);
    const allPages = contexts.flatMap((c) => c.pages());
    console.log(`  Pages: ${allPages.length}`);

    // Prefer main VSCode window
    activePage = allPages.find((p) => p.url().includes("vscode-app"))
      ?? allPages.sort((a, b) => b.url().length - a.url().length)[0]
      ?? null;

    if (activePage) {
      const title = await activePage.title().catch(() => "?");
      console.log(`  Active page: ${title}`);

      // Listen for new pages (wizard opens in new webview)
      contexts.forEach((ctx) => {
        ctx.on("page", (newPage) => {
          console.log("  New page:", newPage.url());
          // Prefer wizard pages over main window
          if (newPage.url().includes("webview") || newPage.url().includes("panel")) {
            activePage = newPage;
          }
        });
      });
    }
  } else {
    console.warn("CDP connect failed - screenshots will be skipped");
  }

  // Wait for test run to complete
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

main().catch((e) => { console.error(e); process.exit(1); });


