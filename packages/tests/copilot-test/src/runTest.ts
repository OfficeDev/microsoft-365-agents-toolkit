// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as path from "path";
import * as fs from "fs";
import * as cp from "child_process";
import * as os from "os";
import {
  downloadAndUnzipVSCode,
  runTests,
  runVSCodeCommand,
} from "@vscode/test-electron";
import { chromium } from "playwright";
import type { Browser, BrowserContext, Page } from "playwright";

const HERE = __dirname;
const TESTS_ROOT = path.resolve(HERE, "..");

function sleep(ms: number) {
  return new Promise<void>((r) => setTimeout(r, ms));
}

interface PagePair {
  mainPage: Page | null;
  galleryPage: Page | null;
}

async function evalInGalleryFrame(
  ctx: BrowserContext,
  mainPage: Page,
  galleryPage: Page | null,
  evalScript: string,
): Promise<string> {
  // Priority: S2 (gallery iframe in mainPage) > S1 (galleryPage CDP target) > S3 (ARIA)
  //
  // S2 is primary because the ATK gallery is a sidebar WebviewView whose iframes are
  // embedded directly in the main VS Code window — reliably found via mainPage.frames().
  // S1 (galleryPage) can be set to a wrong CDP target when other extensions (e.g.
  // redhat.vscode-yaml) open their own WebviewPanels after the gallery opens.

  const frames = mainPage.frames();
  console.log(`  Frames (${frames.length}):`);
  frames.forEach((f, i) => console.log(`    [${i}] ${f.url().slice(0, 100)}`));

  // S2: gallery iframe within mainPage
  // VS Code 1.90+ webview architecture:
  //   outer container = index.html (CSP wrapper)
  //   inner content   = fake.html  (actual React gallery content)
  const galleryFrame =
    frames.find((f) => f.url().includes("/fake.html")) ||
    frames.find((f) => f.url().startsWith("vscode-webview://"));

  if (galleryFrame) {
    console.log(
      `  Gallery frame (selected): ${galleryFrame.url().slice(0, 100)}`,
    );

    // S2a: CDP session on galleryFrame
    try {
      const frameSession = await ctx.newCDPSession(galleryFrame);
      const diagEval = await frameSession.send("Runtime.evaluate", {
        expression:
          'JSON.stringify({href: window.location.href.slice(0,70), elemCount: document.querySelectorAll("*").length, hasRoot: !!document.getElementById("root")})',
        returnByValue: true,
        awaitPromise: false,
      });
      if (diagEval.result?.value) {
        console.log("  [CDP diag]", diagEval.result.value.slice(0, 200));
      }
      const evalResult = await frameSession.send("Runtime.evaluate", {
        expression: evalScript,
        returnByValue: true,
        awaitPromise: false,
      });
      await frameSession.detach();
      if (evalResult.result?.type === "string") return evalResult.result.value;
      if (evalResult.result?.value !== undefined)
        return JSON.stringify(evalResult.result.value);
      if (evalResult.exceptionDetails) {
        console.warn("  CDP eval exception:", evalResult.exceptionDetails.text);
        return "";
      }
      return "";
    } catch (cdpErr: any) {
      console.warn(`  S2a CDP session eval failed: ${cdpErr.message}`);
    }

    // S2b: frame.evaluate fallback
    try {
      try {
        await galleryFrame.waitForSelector(
          ".sample-filter, .offlinePage, .sample-card, .ms-Link",
          { timeout: 5000 },
        );
      } catch {}
      const val = await galleryFrame.evaluate(evalScript);
      return typeof val === "string" ? val : JSON.stringify(val) ?? "";
    } catch (evalErr: any) {
      console.warn(`  S2b frame.evaluate failed: ${evalErr.message}`);
    }
  }

  // S1: galleryPage separate CDP target (fallback when gallery not found as iframe)
  if (galleryPage) {
    console.log(`  S1 fallback: galleryPage ${galleryPage.url().slice(0, 80)}`);
    try {
      await galleryPage.waitForSelector(
        ".sample-filter, .offlinePage, .sample-card, .ms-Link",
        { timeout: 5000 },
      );
    } catch {}
    try {
      const val = await galleryPage.evaluate(evalScript);
      return typeof val === "string" ? val : JSON.stringify(val) ?? "";
    } catch (s1Err: any) {
      console.warn(`  S1 galleryPage eval failed: ${s1Err.message}`);
    }
  } else {
    console.log(
      `  No gallery frame found (${frames.length} frames), no galleryPage`,
    );
  }

  // S3: ARIA snapshot fallback (Playwright 1.40+)
  console.log("  Falling back to ariaSnapshot");
  try {
    const snapshot = await mainPage.ariaSnapshot();
    const outputDir = process.env.TEST_OUTPUT_DIR || "/output";
    const snapFile = path.join(outputDir, "a11y-snapshot.txt");
    if (!fs.existsSync(snapFile)) {
      fs.writeFileSync(snapFile, snapshot, "utf8");
      console.log(`  ARIA snapshot saved (${snapshot.length} chars)`);
    }
    return `ACCESSIBILITY:${snapshot}`;
  } catch (axErr: any) {
    console.warn(`  ariaSnapshot failed: ${axErr.message}`);
    return "";
  }
}
async function startSignalWatcher(
  signalDir: string,
  screenshotDir: string,
  getPages: () => PagePair,
  getCtx: () => BrowserContext | null,
  stopFlag: { stop: boolean },
) {
  fs.mkdirSync(signalDir, { recursive: true });
  console.log("Signal watcher started:", signalDir);

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
        continue;
      }

      const { mainPage: page, galleryPage } = getPages();
      try {
        if (content.startsWith("screenshot:")) {
          const dest = content.slice("screenshot:".length);
          if (page) {
            await page.screenshot({ path: dest, fullPage: false });
            console.log(`  Screenshot: ${path.basename(dest)}`);
          }
        } else if (content.startsWith("clickText:")) {
          const text = content.slice("clickText:".length);
          if (page) {
            // Strategy 1: getByText (robust Playwright text locator, substring match)
            const byText = page.getByText(text, { exact: false }).filter({
              hasNot: page.locator('[class*="outline"], [class*="sidebar"]'),
            });
            try {
              await byText.first().waitFor({ timeout: 8000 });
              await byText.first().click();
              console.log(`  Clicked (getByText): "${text}"`);
              await sleep(500);
            } catch {
              // Strategy 2: try monaco-list-row with hasText
              const byRow = page
                .locator(".monaco-list-row")
                .filter({ hasText: text });
              try {
                await byRow.first().waitFor({ timeout: 5000 });
                await byRow.first().click();
                console.log(`  Clicked (monaco-list-row): "${text}"`);
                await sleep(500);
              } catch {
                // Strategy 3: type to filter QuickPick input, then Enter
                try {
                  const inputLoc = page.locator(
                    ".quick-input-filter .input, .quick-input-box .input",
                  );
                  await inputLoc.first().waitFor({ timeout: 3000 });
                  await inputLoc.first().fill(text);
                  await sleep(500);
                  await page.keyboard.press("Enter");
                  console.log(`  clickText (type+Enter): "${text}"`);
                  await sleep(300);
                } catch {
                  console.warn(`  clickText: "${text}" not found`);
                  await page.keyboard.press("ArrowDown");
                  await sleep(200);
                  await page.keyboard.press("Enter");
                }
              }
            }
          }
        } else if (content.startsWith("click:")) {
          const selector = content.slice("click:".length);
          if (page) {
            try {
              await page.locator(selector).first().waitFor({ timeout: 8000 });
              await page.locator(selector).first().click();
              console.log(`  Clicked selector: ${selector}`);
              await sleep(300);
            } catch (e) {
              console.warn(`  click: "${selector}" not found`);
            }
          }
        } else if (content.startsWith("type:")) {
          const text = content.slice("type:".length);
          if (page) {
            const input = page.locator(
              ".quick-input-box input, .quick-input-filter .input",
            );
            try {
              await input.first().waitFor({ timeout: 5000 });
              await input.first().fill(text);
              console.log(`  Typed: "${text}"`);
              await sleep(300);
            } catch (e) {
              await page.keyboard.type(text, { delay: 30 });
            }
          }
        } else if (content.startsWith("pressKey:")) {
          const key = content.slice("pressKey:".length);
          if (page) {
            await page.keyboard.press(key);
            console.log(`  Key: ${key}`);
            await sleep(300);
          }
        } else if (content.startsWith("waitForText:")) {
          const rest = content.slice("waitForText:".length);
          // Format: "waitForText:text" or "waitForText:text:timeoutMs"
          const lastColon = rest.lastIndexOf(":");
          let text = rest;
          let wftTimeout = 20000;
          if (lastColon > 0 && /^\d+$/.test(rest.slice(lastColon + 1))) {
            text = rest.slice(0, lastColon);
            wftTimeout = parseInt(rest.slice(lastColon + 1), 10);
          }
          if (page) {
            try {
              // Strategy 1: monaco-list-row (VSCode QuickPick items use virtualized list — getByText misses them)
              const byRow = page
                .locator(".monaco-list-row")
                .filter({ hasText: text });
              let found = false;
              try {
                // Use "attached" (not "visible") — QuickPick items may be in DOM but not "visible" per Playwright
                await byRow
                  .first()
                  .waitFor({ state: "attached", timeout: wftTimeout });
                found = true;
              } catch {
                /* fall through to strategy 2 */
              }
              if (!found) {
                // Strategy 2: getByText fallback (InputBox labels, panel text, etc.)
                await page
                  .getByText(text, { exact: false })
                  .first()
                  .waitFor({ timeout: 5000 });
              }
              console.log(`  Found text: "${text}"`);
            } catch {
              console.warn(`  waitForText: "${text}" not found`);
            }
          }
        } else if (content.startsWith("waitForTextThenScreenshot:")) {
          // Format: "waitForTextThenScreenshot:text:timeoutMs:screenshotName"
          // Waits for text to appear, immediately takes screenshot while QuickPick is visible,
          // then deletes the signal. Avoids the timing gap of separate waitForText+screenshot.
          const rest = content.slice("waitForTextThenScreenshot:".length);
          const parts = rest.split(":");
          const screenshotName = parts[parts.length - 1];
          const wftTimeout2 = parseInt(parts[parts.length - 2], 10) || 20000;
          const wftText = parts.slice(0, parts.length - 2).join(":");
          if (page) {
            try {
              const byRow2 = page
                .locator(".monaco-list-row")
                .filter({ hasText: wftText });
              let found2 = false;
              try {
                await byRow2
                  .first()
                  .waitFor({ state: "attached", timeout: wftTimeout2 });
                found2 = true;
              } catch {
                /* fall through */
              }
              if (!found2) {
                await page
                  .getByText(wftText, { exact: false })
                  .first()
                  .waitFor({ timeout: 5000 });
                found2 = true;
              }
              console.log(`  Found text: "${wftText}"`);
              await sleep(600);
              const dest = path.join(screenshotDir, `${screenshotName}.png`);
              await page.screenshot({ path: dest, fullPage: false });
              console.log(`  Screenshot (inline): ${screenshotName}.png`);
            } catch {
              console.warn(
                `  waitForTextThenScreenshot: "${wftText}" not found`,
              );
            }
          }
        } else if (content.startsWith("eval:")) {
          const rest = content.slice("eval:".length);
          // The signal format is "eval:${resultFile}:${evalScript}".
          // On Windows, resultFile contains a drive colon (e.g., C:\...).
          // Use ".result:" as the unambiguous boundary marker.
          const resultSuffix = ".result:";
          const resultEndIdx = rest.indexOf(resultSuffix);
          let resultFile: string;
          let evalScript: string;
          if (resultEndIdx !== -1) {
            resultFile = rest.slice(0, resultEndIdx + ".result".length);
            evalScript = rest.slice(resultEndIdx + resultSuffix.length);
          } else {
            // Fallback: first colon (works on Linux/macOS paths without drive letters)
            const colonIdx = rest.indexOf(":");
            if (colonIdx === -1) {
              console.warn("  Malformed eval signal:", content.slice(0, 80));
              resultFile = "";
              evalScript = "";
            } else {
              resultFile = rest.slice(0, colonIdx);
              evalScript = rest.slice(colonIdx + 1);
            }
          }
          if (resultFile && evalScript) {
            let result = "";
            if (page) {
              try {
                const ctx = getCtx();
                if (ctx) {
                  result = await evalInGalleryFrame(
                    ctx,
                    page,
                    galleryPage,
                    evalScript,
                  );
                  console.log(`  Eval result: ${result.slice(0, 120)}`);
                }
              } catch (evalErr) {
                result = `ERROR:${String(evalErr)}`;
                console.warn(`  Eval failed: ${evalErr}`);
              }
            }
            try {
              fs.writeFileSync(resultFile, result, "utf8");
            } catch (writeErr) {
              console.warn(`  Could not write eval result: ${writeErr}`);
            }
          }
        } else {
          if (page && content) {
            await page.screenshot({ path: content, fullPage: false });
            console.log(`  Screenshot (legacy): ${path.basename(content)}`);
          }
        }
      } catch (e) {
        console.warn(`  Signal error (${sig}):`, e);
      }

      try {
        fs.rmSync(sigPath, { force: true });
      } catch {}
    }

    await sleep(150);
  }
}

async function main() {
  // First positional arg is the test file filter (e.g. `npm test -- sample-app-a11y`).
  // Falls back to TEST_FILE env var (used by CI/Docker) then runs all tests if neither is set.
  const testFile = process.argv[2] || process.env.TEST_FILE || undefined;

  const extPath = process.env.ATK_EXT_PATH ?? "";
  const outputDir =
    process.env.TEST_OUTPUT_DIR ||
    path.resolve(TESTS_ROOT, "../../test-output");
  const screenshotDir = path.join(outputDir, "screenshots");
  const signalDir = path.join(outputDir, ".screenshot-signals");
  const userDataDir = path.join(outputDir, "vscode-user-data");

  // Use a fresh user data dir each run to prevent VS Code extension auto-updates from
  // competing for network bandwidth during template loading.
  try {
    fs.rmSync(userDataDir, { recursive: true, force: true });
  } catch {}
  for (const d of [screenshotDir, signalDir, userDataDir])
    fs.mkdirSync(d, { recursive: true });
  // Disable VS Code extension auto-updates to prevent github.copilot-chat from downloading
  // during the test run, which blocks the extension host and delays wizard QuickPick appearance.
  const vscodeUserSettings = path.join(userDataDir, "User");
  fs.mkdirSync(vscodeUserSettings, { recursive: true });
  fs.writeFileSync(
    path.join(vscodeUserSettings, "settings.json"),
    JSON.stringify(
      {
        "extensions.autoUpdate": false,
        "extensions.autoCheckUpdates": false,
        "update.mode": "none",
        "telemetry.telemetryLevel": "off",
      },
      null,
      2,
    ),
    "utf8",
  );

  fs.readdirSync(signalDir).forEach((f) =>
    fs.rmSync(path.join(signalDir, f), { force: true }),
  );

  console.log("=== Playwright + test-electron Runner ===");
  console.log("Ext:", extPath);
  console.log("Out:", outputDir);

  if (extPath && !fs.existsSync(extPath)) {
    console.error("ATK extension not found:", extPath);
    process.exit(1);
  }

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
  let mainPage: Page | null = null;
  let galleryPage: Page | null = null;
  let activeCtx: BrowserContext | null = null;
  const stopFlag = { stop: false };

  const watcherPromise = startSignalWatcher(
    signalDir,
    screenshotDir,
    () => ({ mainPage, galleryPage }),
    () => activeCtx,
    stopFlag,
  );

  // Use a local temp dir for VS Code extensions (avoids cross-OS mount issues with Docker volumes).
  const userExtDir = path.join(os.tmpdir(), "atk-test-vscode-ext");
  fs.mkdirSync(userExtDir, { recursive: true });

  // Install redhat.vscode-yaml dependency. If a local VSIX path is provided via
  // YAML_STUB_VSIX env var, use it; otherwise install by extension ID directly from
  // the Marketplace — VS Code CLI handles the download automatically.
  const yamlVsixPath = process.env.YAML_STUB_VSIX;
  const yamlExtDir = path.join(userExtDir, "redhat.vscode-yaml");
  if (!fs.existsSync(path.join(yamlExtDir, "package.json"))) {
    const installArg =
      yamlVsixPath && fs.existsSync(yamlVsixPath)
        ? yamlVsixPath
        : "redhat.vscode-yaml";
    try {
      await runVSCodeCommand([
        "--install-extension",
        installArg,
        `--extensions-dir=${userExtDir}`,
      ]);
      console.log(`YAML ext installed from: ${installArg}`);
    } catch (e: any) {
      console.warn("YAML extension install failed:", e.message);
    }
  } else {
    console.log(`YAML ext already present: ${yamlExtDir}`);
  }
  console.log(`Extensions dir: ${userExtDir}`);

  const vscodeTestOpts: any = {
    extensionTestsPath,
    launchArgs: [
      "--disable-workspace-trust",
      "--skip-welcome",
      "--skip-release-notes",
      `--user-data-dir=${userDataDir}`,
      "--no-sandbox",
      "--disable-gpu",
      "--disable-dev-shm-usage",
      `--remote-debugging-port=${CDP_PORT}`,
      `--extensions-dir=${userExtDir}`,
    ],
    version: "stable",
    extensionTestsEnv: {
      TEST_OUTPUT_DIR: outputDir,
      SCREENSHOT_SIGNAL_DIR: signalDir,
      SCREENSHOT_DIR: screenshotDir,
      ...(testFile ? { TEST_FILE: testFile } : {}),
      // Force local bundled templates — prevents ATK from downloading metadata from GitHub
      // (which takes 4+ minutes in Docker/CI). Without this, useLocalTemplate() returns false.
      TEMPLATE_VERSION: process.env.TEMPLATE_VERSION ?? "local",
    },
  };
  if (extPath) vscodeTestOpts.extensionDevelopmentPath = extPath;

  const testRunPromise = runTests(vscodeTestOpts);

  let browser: Browser | null = null;
  let cdpConnected = false;
  const cdpStart = Date.now();
  while (!cdpConnected && Date.now() - cdpStart < 90000) {
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
    allPages.forEach((p, i) => console.log(`  [${i}] ${p.url().slice(0, 80)}`));

    mainPage =
      allPages.find((p) => p.url().includes("vscode-app")) ??
      allPages.sort((a, b) => b.url().length - a.url().length)[0] ??
      null;
    activeCtx = contexts[0] ?? null;

    if (mainPage) {
      const title = await mainPage.title().catch(() => "?");
      console.log(`  Main page: ${title}`);

      // Log frames periodically for debugging
      const frameLogger = setInterval(() => {
        const frs = mainPage!.frames();
        console.log(`  Frames: ${frs.length}`);
        frs.forEach((f, i) =>
          console.log(`    [${i}] ${f.url().slice(0, 80)}`),
        );
      }, 10000);
      setTimeout(() => clearInterval(frameLogger), 120000);

      contexts.forEach((ctx) => {
        ctx.on("page", (newPage) => {
          const url = newPage.url();
          console.log("  New CDP page:", url.slice(0, 80));
          if (
            url.startsWith("vscode-webview://") ||
            url.includes("vscode-webview")
          ) {
            galleryPage = newPage;
            console.log("  galleryPage set:", url.slice(0, 80));
          }
          newPage.on("domcontentloaded", () => {
            const u = newPage.url();
            if (
              u.startsWith("vscode-webview://") ||
              u.includes("vscode-webview")
            ) {
              galleryPage = newPage;
            }
          });
          newPage.on("close", () => {
            if (galleryPage === newPage) {
              galleryPage = null;
            }
          });
        });
      });
    }
  } else {
    console.warn("CDP connect failed");
  }

  let testFailed = false;
  try {
    await testRunPromise;
    console.log("Test run completed");
  } catch (e: any) {
    // runTests() rejects when VS Code exits non-zero (i.e., Mocha reported failures).
    console.error("Test run failed:", e.message);
    testFailed = true;
  } finally {
    stopFlag.stop = true;
    await sleep(300);
    if (browser) await browser.close().catch(() => {});
  }
  if (testFailed) process.exit(1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
