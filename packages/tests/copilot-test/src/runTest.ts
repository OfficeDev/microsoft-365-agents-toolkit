// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as path from "path";
import * as fs from "fs";
import * as cp from "child_process";
import * as os from "os";
import { downloadAndUnzipVSCode, runTests } from "@vscode/test-electron";
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
  // S1: galleryPage = separate CDP target
  if (galleryPage) {
    try {
      await galleryPage.waitForSelector(
        ".sample-filter, .offlinePage, .sample-card, .ms-Link",
        { timeout: 30000 },
      );
    } catch {}
    const val = await galleryPage.evaluate(evalScript);
    return typeof val === "string" ? val : JSON.stringify(val);
  }

  // S2: gallery iframe within mainPage
  // VSCode webviews have TWO frames:
  //   outer: vscode-webview://xxx/ (empty container, no React, no #root)
  //   inner: vscode-webview://xxx/index.html (actual React gallery content)
  // Always target the inner index.html frame.
  const frames = mainPage.frames();
  const allWebviewFrames = frames.filter((f) => f.url().startsWith("vscode-webview://"));
  console.log(`  All frames (${frames.length}):`);
  frames.forEach((f, i) => console.log(`    [${i}] ${f.url().slice(0, 120)}`));
  console.log(`  Webview frames (${allWebviewFrames.length}):`);
  allWebviewFrames.forEach((f, i) => {
    console.log(`    [${i}] url=${f.url().slice(0, 100)} childFrames=${f.childFrames().length}`);
    f.childFrames().forEach((cf, ci) => console.log(`      child[${ci}] ${cf.url().slice(0, 100)}`));
  });
  // In VSCode 1.90+, webview architecture:
  //   outer container = index.html (CSP wrapper, has active-frame iframe)
  //   active-frame loads = fake.html (actual gallery content — confusingly named)
  //   fake.html = the frame that receives the extension webview.html content
  // So we must target fake.html, NOT index.html
  const galleryFrame =
    frames.find((f) => f.url().includes("/fake.html")) ||
    frames.find((f) => f.url().startsWith("vscode-webview://"));

  if (galleryFrame) {
    console.log(`  Gallery frame (selected): ${galleryFrame.url().slice(0, 100)}`);

    // S2a: newCDPSession(frame) — uses Runtime.evaluate directly, may differ from frame.evaluate
    try {
      const frameSession = await ctx.newCDPSession(galleryFrame);
      await sleep(2000);
      // Diagnostic: verify we are in the gallery JS context
      const diagEval = await frameSession.send("Runtime.evaluate", {
        expression: "JSON.stringify({href: window.location.href.slice(0,70), elemCount: document.querySelectorAll(\"*\").length, readyState: document.readyState, hasRoot: !!document.getElementById(\"root\"), html: document.documentElement.outerHTML.slice(0,1000)})",
        returnByValue: true, awaitPromise: false,
      });
      if (diagEval.result?.value) {
        const d = JSON.parse(diagEval.result.value);
        console.log("  [CDP diag]", JSON.stringify({href: d.href, elemCount: d.elemCount, readyState: d.readyState, hasRoot: d.hasRoot}));
        if (!d.hasRoot) {
          console.log("  [CDP html]", (d.html || "").slice(0, 1200));
          // Also check for iframe src in the container
          const iframeEval = await frameSession.send("Runtime.evaluate", {
            expression: "JSON.stringify({iframes: Array.from(document.querySelectorAll(\"iframe\")).map(f => ({id:f.id, src:f.src, name:f.name}))})",
            returnByValue: true, awaitPromise: false,
          });
          if (iframeEval.result?.value) console.log("  [CDP iframes]", iframeEval.result.value);
        }
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
        console.warn(
          "  CDP eval exception:",
          evalResult.exceptionDetails.text,
        );
        return "";
      }
      return "";
    } catch (cdpErr: any) {
      console.warn(`  CDP session eval failed: ${cdpErr.message}`);
    }

    // S2b: frame.evaluate fallback
    try {
      try {
        await galleryFrame.waitForSelector(
          ".sample-filter, .offlinePage, .sample-card, .ms-Link",
          { timeout: 15000 },
        );
      } catch {}
      const val = await galleryFrame.evaluate(evalScript);
      return typeof val === "string" ? val : JSON.stringify(val);
    } catch (evalErr: any) {
      console.warn(`  frame.evaluate failed: ${evalErr.message}`);
    }
  } else {
    console.log(
      `  No gallery frame (${frames.length} frames total), using main page`,
    );
    const val = await mainPage.evaluate(evalScript);
    return typeof val === "string" ? val : JSON.stringify(val);
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
            const item = page
              .locator(".quick-input-list .monaco-list-row")
              .filter({ hasText: text });
            try {
              await item.first().waitFor({ timeout: 8000 });
              await item.first().click();
              console.log(`  Clicked: "${text}"`);
              await sleep(500);
            } catch {
              console.warn(`  clickText: "${text}" not found`);
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
              await page.waitForSelector(`text="${text}"`, { timeout: wftTimeout });
              console.log(`  Found text: "${text}"`);
            } catch {
              console.warn(`  waitForText: "${text}" not found`);
            }
          }
        } else if (content.startsWith("eval:")) {
          const rest = content.slice("eval:".length);
          const colonIdx = rest.indexOf(":");
          if (colonIdx !== -1) {
            const resultFile = rest.slice(0, colonIdx);
            const evalScript = rest.slice(colonIdx + 1);
            let result = "";
            if (page) {
              try {
                const ctx = getCtx();
                if (ctx) {
                  result = await evalInGalleryFrame(ctx, page, galleryPage, evalScript);
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
  const extPath = process.env.ATK_EXT_PATH ?? "";
  const outputDir =
    process.env.TEST_OUTPUT_DIR ||
    path.resolve(TESTS_ROOT, "../../test-output");
  const screenshotDir = path.join(outputDir, "screenshots");
  const signalDir = path.join(outputDir, ".screenshot-signals");
  const userDataDir = path.join(outputDir, "vscode-user-data");

  for (const d of [screenshotDir, signalDir, userDataDir])
    fs.mkdirSync(d, { recursive: true });

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
      module: "commonjs", target: "ES2020", lib: ["ES2020"],
      esModuleInterop: true, resolveJsonModule: true, strict: false,
      skipLibCheck: true, outDir: tmpOut, rootDir: HERE,
      types: ["node", "mocha"],
      typeRoots: [path.join(TESTS_ROOT, "node_modules", "@types")],
    },
    include: ["suite/**/*.ts", "*.test.ts"],
    exclude: ["node_modules", "runTest.ts"],
  };
  fs.writeFileSync(tsconfigPath, JSON.stringify(tsconfig, null, 2), "utf8");
  const tscBin = path.join(
    TESTS_ROOT, "node_modules", ".bin",
    process.platform === "win32" ? "tsc.CMD" : "tsc",
  );
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

  const CDP_PORT = 9229;
  let mainPage: Page | null = null;
  let galleryPage: Page | null = null;
  let activeCtx: BrowserContext | null = null;
  const stopFlag = { stop: false };

  const watcherPromise = startSignalWatcher(
    signalDir,
    () => ({ mainPage, galleryPage }),
    () => activeCtx,
    stopFlag,
  );

  const userExtDir =
    process.env.VSCODE_EXTENSIONS_DIR ||
    path.join(os.homedir(), ".vscode", "extensions");
  const yamlExtPresent =
    fs.existsSync(path.join(userExtDir, "redhat.vscode-yaml", "package.json")) ||
    fs.readdirSync(userExtDir).some(
      (d) => d.startsWith("redhat.vscode-yaml") &&
        fs.existsSync(path.join(userExtDir, d, "package.json")),
    );
  console.log(`Extensions dir: ${userExtDir} (yaml: ${yamlExtPresent})`);

  const vscodeTestOpts: any = {
    extensionTestsPath,
    launchArgs: [
      "--disable-workspace-trust", "--skip-welcome", "--skip-release-notes",
      `--user-data-dir=${userDataDir}`, "--no-sandbox", "--disable-gpu",
      "--disable-dev-shm-usage",
      `--remote-debugging-port=${CDP_PORT}`,
      `--extensions-dir=${userExtDir}`,
      ...(yamlExtPresent ? [] : ["--install-extension", "redhat.vscode-yaml"]),
    ],
    version: "stable",
    extensionTestsEnv: {
      TEST_OUTPUT_DIR: outputDir,
      SCREENSHOT_SIGNAL_DIR: signalDir,
      SCREENSHOT_DIR: screenshotDir,
      ...(process.env.TEST_FILE ? { TEST_FILE: process.env.TEST_FILE } : {}),
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
      browser = await chromium.connectOverCDP(`http://localhost:${CDP_PORT}`, { timeout: 2000 });
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
        frs.forEach((f, i) => console.log(`    [${i}] ${f.url().slice(0, 80)}`));
      }, 10000);
      setTimeout(() => clearInterval(frameLogger), 120000);

      contexts.forEach((ctx) => {
        ctx.on("page", (newPage) => {
          const url = newPage.url();
          console.log("  New CDP page:", url.slice(0, 80));
          if (url.startsWith("vscode-webview://") || url.includes("vscode-webview")) {
            galleryPage = newPage;
            console.log("  galleryPage set:", url.slice(0, 80));
          }
          newPage.on("domcontentloaded", () => {
            const u = newPage.url();
            if (u.startsWith("vscode-webview://") || u.includes("vscode-webview")) {
              galleryPage = newPage;
            }
          });
          newPage.on("close", () => {
            if (galleryPage === newPage) { galleryPage = null; }
          });
        });
      });
    }
  } else {
    console.warn("CDP connect failed");
  }

  try {
    await testRunPromise;
    console.log("Test run completed");
  } catch (e: any) {
    console.error("Test run failed:", e.message);
  } finally {
    stopFlag.stop = true;
    await sleep(300);
    if (browser) await browser.close().catch(() => {});
  }
}

main().catch((e) => { console.error(e); process.exit(1); });









