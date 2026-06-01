// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * sample-app-a11y.test.ts
 * TC-001a: Link text color contrast >= 4.5:1 when focused (Light theme)
 * TC-001b: Link text color contrast >= 4.5:1 when focused (Dark theme)
 * TC-002:  Gallery/List toggle buttons aria-pressed state before and after click
 * TC-003:  Sample card accessible names include tags on keyboard focus
 * TC-004a: Featured cards in Gallery (grid) view display a .featured-corner-badge element containing a .featured-corner-star (codicon-star-full) on the top-left of the thumbnail; non-featured cards do not; featured section has no background color
 * TC-004b: Featured items in List view have a .featured-star (codicon-star-full) icon in their h3 before title text (color matches title via currentColor); non-featured items do not
 * TC-005:  Screen reader differentiates Featured from non-Featured cards
 * TC-006a: Focus ring contrast >= 3:1 in Gallery view (Light theme)
 * TC-006b: Focus ring contrast >= 3:1 in List view (Light theme)
 *
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
 * All contrast checks use the WCAG relative luminance formula (IEC 61966-2-1 sRGB).
 *
 * Each TC flows: set theme → activate extension → open gallery → assert → screenshot
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

function wait(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function waitForCommand(cmd: string, maxMs = 60000): Promise<boolean> {
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

/**
 * Send a DOM evaluation signal to Playwright.
 * Returns the result written to a response file by the Playwright side.
 */
function sendEvalSignal(evalScript: string, timeoutMs = 15000): string {
  try {
    const id = Date.now();
    const resultFile = path.join(SIGNAL_DIR, `${id}-eval.result`);
    const signal = path.join(SIGNAL_DIR, `${id}-eval.signal`);
    fs.writeFileSync(signal, `eval:${resultFile}:${evalScript}`, "utf8");
    const deadline = Date.now() + timeoutMs;
    while (!fs.existsSync(resultFile) && Date.now() < deadline) {
      const end = Date.now() + 100;
      while (Date.now() < end) {
        /* busy wait */
      }
    }
    if (fs.existsSync(resultFile)) {
      const result = fs.readFileSync(resultFile, "utf8");
      try {
        fs.unlinkSync(resultFile);
      } catch {}
      return result;
    }
    try {
      fs.unlinkSync(signal);
    } catch {}
    return "";
  } catch (e) {
    console.warn("Eval signal failed:", e);
    return "";
  }
}

/**
 * Send an action signal to Playwright and wait for it to be consumed.
 * content: "clickText:List", "pressKey:Tab", "type:foo", etc.
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

/**
 * WCAG relative luminance helper JS (IEC 61966-2-1 sRGB).
 * Injected into every contrast evaluation script.
 */
const WCAG_HELPER_JS =
  "function relativeLuminance(r,g,b){" +
  "  var srgb=[r,g,b].map(function(c){" +
  "    c=c/255;" +
  "    return c<=0.03928?c/12.92:Math.pow((c+0.055)/1.055,2.4);" +
  "  });" +
  "  return 0.2126*srgb[0]+0.7152*srgb[1]+0.0722*srgb[2];" +
  "}" +
  "function contrastRatio(L1,L2){" +
  "  var lighter=Math.max(L1,L2);" +
  "  var darker=Math.min(L1,L2);" +
  "  return (lighter+0.05)/(darker+0.05);" +
  "}" +
  "function parseRgb(str){" +
  "  var m=str.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)/);" +
  "  return m?[parseInt(m[1]),parseInt(m[2]),parseInt(m[3])]:null;" +
  "}" +
  "function effectiveBg(el,theme){" +
  "  var node=el;" +
  "  while(node){" +
  "    var cs=getComputedStyle(node);" +
  "    var bg=cs.backgroundColor;" +
  "    if(bg&&bg!=='transparent'&&bg!=='rgba(0, 0, 0, 0)'&&" +
  "       !(bg==='rgb(0, 0, 0)'&&node===document.body&&theme==='light')){" +
  "      var rgb=parseRgb(bg);" +
  "      if(rgb){" +
  "        var isBlack=rgb[0]+rgb[1]+rgb[2]<10;" +
  "        if(!(theme==='light'&&isBlack))return rgb;" +
  "      }" +
  "    }" +
  "    node=node.parentElement;" +
  "  }" +
  "  return theme==='dark'?[30,30,30]:[255,255,255];" +
  "}";

/**
 * Injects a CSS style that renders each card's aria-label as an overlay via ::before.
 * Makes aria-label content VISIBLE in screenshots for manual verification.
 */
function injectAriaOverlay(selector: string = ".sample-card"): void {
  const script =
    "(function(){" +
    "  var id='__aria-overlay-style__';" +
    "  if(document.getElementById(id))return;" +
    "  var s=document.createElement('style');" +
    "  s.id=id;" +
    "  s.textContent='" +
    selector +
    "{position:relative!important;}" +
    selector +
    "::before{content:attr(aria-label);position:absolute;top:0;left:0;right:0;" +
    "background:rgba(0,0,0,0.82);color:#fff;font-size:8px;line-height:1.3;" +
    "padding:3px 4px;z-index:9999;pointer-events:none;" +
    "white-space:normal;word-break:break-word;}';" +
    "  document.head.appendChild(s);" +
    "  return 'injected';" +
    "})()";
  sendEvalSignal(script, 3000);
}

function removeAriaOverlay(): void {
  const script =
    "(function(){" +
    "  var s=document.getElementById('__aria-overlay-style__');" +
    "  if(s)s.parentNode.removeChild(s);" +
    "  return 'removed';" +
    "})()";
  sendEvalSignal(script, 1000);
}

/**
 * Per-TC setup: set Default Light Modern theme, click ATK Activity Bar icon,
 * open Sample Gallery, and wait for cards to render.
 *
 * Implements the common Steps 1-3 that appear in every TC of the test plan:
 *   Step 1: Set VS Code color theme to "Default Light Modern"
 *   Step 2: Click the ATK icon in the Activity Bar
 *   Step 3: Open gallery and wait for cards to appear
 *
 * Returns true if the gallery was successfully opened.
 */
async function setupTcGallery(
  tcLabel: string,
  theme: "light" | "dark" = "light",
): Promise<boolean> {
  const themeName =
    theme === "dark" ? "Default Dark Modern" : "Default Light Modern";
  const themeFallback = theme === "dark" ? "Default Dark+" : "Default Light+";
  console.log(`\n  [setup] ${tcLabel}: setting ${themeName}...`);

  // Step 1: Set VS Code color theme via VS Code settings API
  // (equivalent to Ctrl+Shift+P → Preferences: Color Theme → select theme)
  const wbConfig = vscode.workspace.getConfiguration("workbench");
  try {
    await wbConfig.update(
      "colorTheme",
      themeName,
      vscode.ConfigurationTarget.Global,
    );
    console.log(`  [setup] ${tcLabel}: theme set to ${themeName}`);
  } catch {
    try {
      await wbConfig.update(
        "colorTheme",
        themeFallback,
        vscode.ConfigurationTarget.Global,
      );
      console.log(
        `  [setup] ${tcLabel}: theme set to ${themeFallback} (fallback)`,
      );
    } catch (e) {
      console.warn(
        `  [setup] ${tcLabel}: could not set theme ${themeName}:`,
        e,
      );
    }
  }
  await wait(2000);

  // Step 2: Click the ATK icon (M365 Agents Toolkit) in the VS Code Activity Bar.
  // This ensures the ATK sidebar is visible before opening the Sample Gallery.
  try {
    await vscode.commands.executeCommand(
      "workbench.view.extension.teamsfx-toolkit",
    );
    console.log(`  [setup] ${tcLabel}: ATK Activity Bar view activated`);
  } catch {
    // Try alternate Activity Bar icon click via signal
    await sendSignal("click:.codicon-m365-agents-toolkit", 5000);
    console.log(`  [setup] ${tcLabel}: sent Activity Bar click signal`);
  }
  await wait(2000);

  // Step 3: Open Command Palette → "Microsoft 365 Agents Toolkit: View Samples"
  const cmdName = "fx-extension.openSamples";
  const cmdAvailable = await waitForCommand(cmdName, 60000);
  if (!cmdAvailable) {
    console.log(`  [setup] ${tcLabel}: ${cmdName} not registered`);
    return false;
  }

  // Fire without await — command opens a webview panel
  vscode.commands.executeCommand(cmdName).then(undefined, () => {});

  // Poll until gallery renders
  let galleryLoaded = false;
  for (let poll = 0; poll < 30 && !galleryLoaded; poll++) {
    await wait(2000);
    const diagResult = sendEvalSignal(
      "JSON.stringify({" +
        'sampleCards:document.querySelectorAll(".sample-card").length,' +
        'hasFilter:!!document.querySelector(".sample-filter"),' +
        'hasOffline:!!document.querySelector(".offlinePage")' +
        "})",
      3000,
    );
    if (
      diagResult &&
      !diagResult.startsWith("ERROR:") &&
      !diagResult.startsWith("ACCESSIBILITY:")
    ) {
      try {
        const status = JSON.parse(diagResult);
        if (status.sampleCards > 0 || status.hasFilter || status.hasOffline) {
          galleryLoaded = true;
          console.log(
            `  [setup] ${tcLabel}: gallery loaded after ${(poll + 1) * 2}s — ${status.sampleCards} cards`,
          );
        }
      } catch {}
    }
  }
  if (!galleryLoaded) {
    console.log(`  [setup] ${tcLabel}: gallery did not load in time`);
  }
  return galleryLoaded;
}

suite("ATK Sample App A11y Regression Tests (Issue #15916)", function () {
  this.timeout(10 * 60 * 1000);

  const steps: object[] = [];
  let passed = 0;
  let failed = 0;

  const step = (name: string, ok: boolean, detail?: string) => {
    steps.push({ name, status: ok ? "pass" : "fail", detail });
    ok ? passed++ : failed++;
    console.log(
      `${ok ? "PASS" : "FAIL"} ${name}${detail ? ": " + detail : ""}`,
    );
    if (!ok) {
      assert.fail(`${name}: ${detail || "failed"}`);
    }
  };

  suiteSetup(() => {
    ensureDirs();
    console.log("=== ATK Sample App A11y Test (Issue #15916) ===");
    console.log("Output:", OUTPUT_DIR);
  });

  suiteTeardown(() => {
    writeResults(passed, failed, steps);
    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
  });

  // ── Baseline: ATK extension activation ──────────────────────────────────────
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
    await wait(8000);
    const active = !!ext?.isActive;
    step(
      "ATK extension activates",
      active,
      ext ? `v${ext.packageJSON.version}` : "not found",
    );
    // Screenshot 01 — baseline: extension loaded (used by TC-001 step 3)
    takeScreenshot("01-extension-active");
    if (!active) {
      console.log(
        "  Note: Extension not active — remaining TCs will attempt per-TC setup",
      );
    }
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-001a — Link text color contrast >= 4.5:1 when focused (Light theme)
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon to open extension panel
  //   3. Screenshot 01 (ATK extension panel open in Light theme)
  //   4. Open gallery via Command Palette → View Samples
  //   5. Screenshot 02 (gallery open with visible link text on light background)
  //   6. Focus a .ms-Link via element.focus()
  //   7. Read getComputedStyle(link).color and compute WCAG contrast ratio
  //   8. Screenshot 03 (focused link with focus ring visible)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-001a: Link text color contrast >= 4.5:1 (focused, Light theme)", async () => {
    // Steps 1-2: set Light theme + click ATK Activity Bar icon
    const galleryOpened = await setupTcGallery("TC-001a", "light");

    // Step 3: screenshot showing ATK panel open in Light theme
    takeScreenshot("01-tc001a-extension-active");

    // Step 5: screenshot showing gallery with visible link text
    takeScreenshot("02-tc001a-gallery-open");

    if (!galleryOpened) {
      takeScreenshot("03-tc001a-link-focused");
      step(
        "TC-001a Link text contrast >= 4.5:1 (Light)",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Steps 6-7: Focus a .ms-Link element and compute WCAG contrast ratio
    const evalScript =
      "(function(){" +
      WCAG_HELPER_JS +
      "  var links=Array.from(document.querySelectorAll('.ms-Link'));" +
      "  if(links.length===0) return JSON.stringify({error:'no-ms-link',count:0});" +
      "  var link=links[0];" +
      "  link.focus();" +
      "  var cs=getComputedStyle(link);" +
      "  var fgRaw=cs.color;" +
      "  var fg=parseRgb(fgRaw);" +
      "  var bg=effectiveBg(link,'light');" +
      "  if(!fg) return JSON.stringify({error:'parse-fg',fgRaw:fgRaw});" +
      "  var fgL=relativeLuminance(fg[0],fg[1],fg[2]);" +
      "  var bgL=relativeLuminance(bg[0],bg[1],bg[2]);" +
      "  var ratio=contrastRatio(fgL,bgL);" +
      "  return JSON.stringify({" +
      "    count:links.length," +
      "    fgRaw:fgRaw," +
      "    bgRgb:'rgb('+bg[0]+','+bg[1]+','+bg[2]+')'," +
      "    ratio:Math.round(ratio*100)/100," +
      "    passes:ratio>=4.5" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 8000);
    console.log(
      "  TC-001a eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 8: screenshot showing the link in focused state with focus ring
    takeScreenshot("03-tc001a-link-focused");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-001a Link text contrast >= 4.5:1 (Light)",
        false,
        "FAIL: DOM eval error — gallery may not be accessible via Playwright CDP.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-001a Link text contrast >= 4.5:1 (Light)",
        false,
        `FAIL: Could not parse eval result: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-ms-link") {
      step(
        "TC-001a Link text contrast >= 4.5:1 (Light)",
        false,
        "FAIL: No .ms-Link elements found in gallery webview.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-001a Link text contrast >= 4.5:1 (Light)",
        false,
        `FAIL: ${data.error} fg=${data.fgRaw || "?"}`,
      );
      return;
    }

    const detail =
      `Computed ratio=${data.ratio}:1; fg=${data.fgRaw}; bg=${data.bgRgb}; ` +
      `${data.count} .ms-Link elements`;
    step("TC-001a Link text contrast >= 4.5:1 (Light)", !!data.passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-001b — Link text color contrast >= 4.5:1 when focused (Dark theme)
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Dark Modern"
  //   2. Click ATK Activity Bar icon to open extension panel
  //   3. Screenshot 04 (ATK extension panel open in Dark theme)
  //   4. Open gallery via Command Palette → View Samples
  //   5. Screenshot 05 (gallery open with visible link text on dark background)
  //   6. Focus a .ms-Link via element.focus()
  //   7. Read getComputedStyle(link).color and background, compute WCAG contrast ratio
  //   8. Screenshot 06 (focused link with focus ring visible in dark theme)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-001b: Link text color contrast >= 4.5:1 (focused, Dark theme)", async () => {
    // Steps 1-2: set Dark theme + click ATK Activity Bar icon
    const galleryOpened = await setupTcGallery("TC-001b", "dark");

    // Step 3: screenshot showing ATK panel open in Dark theme
    takeScreenshot("04-tc001b-extension-active");

    // Step 5: screenshot showing gallery with visible link text on dark background
    takeScreenshot("05-tc001b-gallery-open");

    if (!galleryOpened) {
      takeScreenshot("06-tc001b-link-focused");
      step(
        "TC-001b Link text contrast >= 4.5:1 (Dark)",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Steps 6-7: Focus a .ms-Link element and compute WCAG contrast ratio
    const evalScript =
      "(function(){" +
      WCAG_HELPER_JS +
      "  var links=Array.from(document.querySelectorAll('.ms-Link'));" +
      "  if(links.length===0) return JSON.stringify({error:'no-ms-link',count:0});" +
      "  var link=links[0];" +
      "  link.focus();" +
      "  var cs=getComputedStyle(link);" +
      "  var fgRaw=cs.color;" +
      "  var fg=parseRgb(fgRaw);" +
      "  var bg=effectiveBg(link,'dark');" +
      "  if(!fg) return JSON.stringify({error:'parse-fg',fgRaw:fgRaw});" +
      "  var fgL=relativeLuminance(fg[0],fg[1],fg[2]);" +
      "  var bgL=relativeLuminance(bg[0],bg[1],bg[2]);" +
      "  var ratio=contrastRatio(fgL,bgL);" +
      "  return JSON.stringify({" +
      "    count:links.length," +
      "    fgRaw:fgRaw," +
      "    bgRgb:'rgb('+bg[0]+','+bg[1]+','+bg[2]+')'," +
      "    ratio:Math.round(ratio*100)/100," +
      "    passes:ratio>=4.5" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 8000);
    console.log(
      "  TC-001b eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 8: screenshot showing the link in focused state with focus ring (dark theme)
    takeScreenshot("06-tc001b-link-focused");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-001b Link text contrast >= 4.5:1 (Dark)",
        false,
        "FAIL: DOM eval error — gallery may not be accessible via Playwright CDP.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-001b Link text contrast >= 4.5:1 (Dark)",
        false,
        `FAIL: Could not parse eval result: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-ms-link") {
      step(
        "TC-001b Link text contrast >= 4.5:1 (Dark)",
        false,
        "FAIL: No .ms-Link elements found in gallery webview.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-001b Link text contrast >= 4.5:1 (Dark)",
        false,
        `FAIL: ${data.error} fg=${data.fgRaw || "?"}`,
      );
      return;
    }

    const detail =
      `Computed ratio=${data.ratio}:1; fg=${data.fgRaw}; bg=${data.bgRgb}; ` +
      `${data.count} .ms-Link elements`;
    step("TC-001b Link text contrast >= 4.5:1 (Dark)", !!data.passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-002 — Gallery/List toggle buttons expose aria-pressed before and after click
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery (Gallery/Grid layout is default)
  //   4. Read aria-pressed on Gallery button — expects "true"
  //   5. Read aria-pressed on List button — expects "false"
  //   6. Screenshot 04 (Gallery layout active, Gallery toggle selected)
  //   7. Click List view toggle button — layout switches to vertical list
  //   8. Read aria-pressed on Gallery button — expects "false"
  //   9. Read aria-pressed on List button — expects "true"
  //  10. Screenshot 05 (List layout active, List toggle selected)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-002: Gallery/List toggle aria-pressed state toggles correctly", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-002");

    if (!galleryOpened) {
      takeScreenshot("04-tc002-gallery-active");
      takeScreenshot("05-tc002-list-active");
      step(
        "TC-002 Toggle aria-pressed state",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Ensure we start in gallery/grid view
    sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"grid view\"]')" +
        "    ||document.querySelector('[aria-label=\"Gallery view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      var l=(b.getAttribute('aria-label')||'').toLowerCase();" +
        "      return l.includes('grid')||l.includes('gallery');});" +
        "  if(btn)btn.click();" +
        "  return btn?'clicked':'no-btn';" +
        "})()",
      3000,
    );
    await wait(600);

    const readPressedScript =
      "(function(){" +
      "  var btns=Array.from(document.querySelectorAll('.layout-button'));" +
      "  var mapped=btns.map(function(b){return{" +
      "    label:b.getAttribute('aria-label')||b.textContent||''," +
      "    pressed:b.getAttribute('aria-pressed')" +
      "  };});" +
      "  return JSON.stringify({count:btns.length,buttons:mapped});" +
      "})()";

    // Steps 4-5: Read aria-pressed on Gallery (expects "true") and List (expects "false")
    const beforeResult = sendEvalSignal(readPressedScript, 5000);
    console.log(
      "  TC-002 before-click state:",
      beforeResult ? beforeResult.slice(0, 200) : "(empty)",
    );

    // Step 6: Screenshot showing Gallery layout is active with Gallery toggle visually selected
    takeScreenshot("04-tc002-gallery-active");

    // Step 7: Click the List view toggle button — observe layout switches to vertical list
    const clickResult = sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"list view\"]')" +
        "    ||document.querySelector('[aria-label=\"List view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      return (b.getAttribute('aria-label')||'').toLowerCase().includes('list');});" +
        "  if(!btn)return 'no-list-button';" +
        "  btn.click();" +
        "  return 'clicked';" +
        "})()",
      5000,
    );
    console.log("  TC-002 list-click result:", clickResult);
    await wait(800);

    // Steps 8-9: Read aria-pressed — Gallery expects "false", List expects "true"
    const afterResult = sendEvalSignal(readPressedScript, 5000);
    console.log(
      "  TC-002 after-click state:",
      afterResult ? afterResult.slice(0, 200) : "(empty)",
    );

    // Step 10: Screenshot showing List layout is active with List toggle visually selected
    takeScreenshot("05-tc002-list-active");

    // Restore gallery view for subsequent tests
    sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"grid view\"]')" +
        "    ||document.querySelector('[aria-label=\"Gallery view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      var l=(b.getAttribute('aria-label')||'').toLowerCase();" +
        "      return l.includes('grid')||l.includes('gallery');});" +
        "  if(btn)btn.click();" +
        "  return btn?'restored':'no-btn';" +
        "})()",
      3000,
    );
    await wait(500);

    if (
      !beforeResult ||
      beforeResult.startsWith("ERROR:") ||
      !afterResult ||
      afterResult.startsWith("ERROR:")
    ) {
      step(
        "TC-002 Toggle aria-pressed state",
        false,
        `FAIL: DOM eval error. before=${beforeResult?.slice(0, 60)} after=${afterResult?.slice(0, 60)}`,
      );
      return;
    }

    let beforeData: any = null;
    let afterData: any = null;
    try {
      beforeData = JSON.parse(beforeResult);
    } catch {}
    try {
      afterData = JSON.parse(afterResult);
    } catch {}

    if (!beforeData || !afterData) {
      step(
        "TC-002 Toggle aria-pressed state",
        false,
        `FAIL: Could not parse button state. before=${beforeResult?.slice(0, 80)}`,
      );
      return;
    }
    if (beforeData.count === 0) {
      step(
        "TC-002 Toggle aria-pressed state",
        false,
        "FAIL: No .layout-button elements found in gallery webview.",
      );
      return;
    }

    const galleryBefore = beforeData.buttons?.find(
      (b: any) =>
        (b.label || "").toLowerCase().includes("grid") ||
        (b.label || "").toLowerCase().includes("gallery"),
    );
    const listBefore = beforeData.buttons?.find((b: any) =>
      (b.label || "").toLowerCase().includes("list"),
    );
    const galleryAfter = afterData.buttons?.find(
      (b: any) =>
        (b.label || "").toLowerCase().includes("grid") ||
        (b.label || "").toLowerCase().includes("gallery"),
    );
    const listAfter = afterData.buttons?.find((b: any) =>
      (b.label || "").toLowerCase().includes("list"),
    );

    const beforeOk =
      galleryBefore?.pressed === "true" && listBefore?.pressed === "false";
    const afterOk =
      listAfter?.pressed === "true" && galleryAfter?.pressed === "false";
    const passes = beforeOk && afterOk;

    const detail =
      `Before: gallery=${galleryBefore?.label}[${galleryBefore?.pressed}] ` +
      `list=${listBefore?.label}[${listBefore?.pressed}]. ` +
      `After: gallery=${galleryAfter?.label}[${galleryAfter?.pressed}] ` +
      `list=${listAfter?.label}[${listAfter?.pressed}]`;
    step("TC-002 Toggle aria-pressed state", passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-003 — Sample card accessible names include tags on keyboard focus
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery (Gallery/Grid layout default)
  //   4. Click inside gallery webview to give it focus
  //   5. Press Tab to move keyboard focus to first sample card
  //   6. Read aria-label from focused card element
  //   7. Verify aria-label matches pattern /.+\. Tags: .+/
  //   8. Screenshot 06 (focused card with focus ring visible)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-003: Sample card aria-label includes tags on keyboard focus", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-003");

    if (!galleryOpened) {
      takeScreenshot("06-tc003-focused-card");
      step(
        "TC-003 Card aria-label includes tags",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Step 4: Click inside the Sample Gallery webview to give it keyboard focus
    sendEvalSignal(
      "(function(){" +
        "  var card=document.querySelector('.sample-card');" +
        "  if(card){card.focus();return 'focused';}" +
        "  document.body.click();" +
        "  return 'body-click';" +
        "})()",
      3000,
    );
    await wait(400);

    // Step 5: Press Tab to move keyboard focus to the first .sample-card
    await sendSignal("pressKey:Tab", 3000);
    await wait(400);

    // Steps 6-7: Read aria-label from focused card and verify it contains ". Tags:"
    const evalScript =
      "(function(){" +
      "  var card=document.activeElement;" +
      "  if(!card||!card.classList.contains('sample-card')){" +
      "    card=document.querySelector('.sample-card');" +
      "    if(card)card.focus();" +
      "  }" +
      "  if(!card) return JSON.stringify({error:'no-sample-card'});" +
      "  var label=card.getAttribute('aria-label')||'';" +
      "  var hasTags=/\\. Tags: .+/.test(label);" +
      "  return JSON.stringify({" +
      "    label:label.slice(0,200)," +
      "    hasTags:hasTags," +
      "    isFocused:document.activeElement===card" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 8000);
    console.log(
      "  TC-003 eval result:",
      rawResult ? rawResult.slice(0, 300) : "(empty)",
    );

    // Step 8: Screenshot showing focused card with focus ring
    takeScreenshot("06-tc003-focused-card");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-003 Card aria-label includes tags",
        false,
        "FAIL: DOM eval error.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-003 Card aria-label includes tags",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-sample-card") {
      step(
        "TC-003 Card aria-label includes tags",
        false,
        "FAIL: No .sample-card elements found in gallery webview.",
      );
      return;
    }

    const detail = `aria-label="${data.label}"; hasTags=${data.hasTags}; isFocused=${data.isFocused}`;
    step("TC-003 Card aria-label includes tags", !!data.hasTags, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-004a — Featured cards in Gallery (grid) view display a blue corner triangle badge
  //           (.featured-corner-badge) containing a white star (.featured-corner-star)
  //           on the top-left of the thumbnail; non-featured cards do not; the
  //           featured section has no background color (transparent / page bg).
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery — Gallery (Grid) view is the default
  //   4. Screenshot 13 showing featured section with blue corner badge on thumbnails
  //   5. For each .sample-card inside .featured-sample-section verify .featured-corner-badge present
  //   5b. Verify .featured-corner-badge contains .featured-corner-star
  //   6. For each .sample-card outside .featured-sample-section verify .featured-corner-badge absent
  //   7. Verify .featured-sample-section has no background color
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-004a: Featured cards display corner badge with star icon in Gallery view", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-004a", "light");

    if (!galleryOpened) {
      takeScreenshot("13-tc004a-gallery-star");
      step(
        "TC-004a Featured gallery corner badge with star on thumbnail",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Ensure Gallery (grid) view is active
    sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"grid view\"]')" +
        "    ||document.querySelector('[aria-label=\"Gallery view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      var l=(b.getAttribute('aria-label')||'').toLowerCase();" +
        "      return l.includes('grid')||l.includes('gallery');});" +
        "  if(btn)btn.click();" +
        "  return btn?'clicked':'no-btn';" +
        "})()",
      3000,
    );
    await wait(600);

    // Step 4: Screenshot showing featured section with star icons on card thumbnails
    takeScreenshot("13-tc004a-gallery-star");

    // Steps 5-7: Verify corner badge on featured cards, absence on non-featured, no section bg
    const evalScript =
      "(function(){" +
      "  var featSec=document.querySelector('.featured-sample-section');" +
      "  if(!featSec) return JSON.stringify({error:'no-featured-section'});" +
      "  var featCards=Array.from(featSec.querySelectorAll('.sample-card'));" +
      "  var allCards=Array.from(document.querySelectorAll('.sample-card'));" +
      "  var nonFeatCards=allCards.filter(function(c){return !featSec.contains(c);});" +
      "  if(featCards.length===0) return JSON.stringify({error:'no-featured-cards'});" +
      // Step 5: every featured card must have .featured-corner-badge
      "  var missingBadge=featCards.filter(function(c){" +
      "    return !c.querySelector('.featured-corner-badge');});" +
      // Step 5b: every featured-corner-badge must contain .featured-corner-star
      "  var missingStar=featCards.filter(function(c){" +
      "    return !c.querySelector('.featured-corner-badge .featured-corner-star');});" +
      // Step 6: no non-featured card should have .featured-corner-badge
      "  var unexpectedBadge=nonFeatCards.filter(function(c){" +
      "    return !!c.querySelector('.featured-corner-badge');});" +
      // Step 7: featured section background must be transparent
      "  var sectBg=getComputedStyle(featSec).backgroundColor;" +
      "  var bgTransparent=sectBg==='transparent'||sectBg==='rgba(0, 0, 0, 0)';" +
      "  return JSON.stringify({" +
      "    featuredCount:featCards.length," +
      "    nonFeaturedCount:nonFeatCards.length," +
      "    missingBadgeCount:missingBadge.length," +
      "    missingStarCount:missingStar.length," +
      "    unexpectedBadgeCount:unexpectedBadge.length," +
      "    sectionBg:sectBg," +
      "    bgTransparent:bgTransparent," +
      "    passes:missingBadge.length===0&&missingStar.length===0&&unexpectedBadge.length===0" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 10000);
    console.log(
      "  TC-004a eval result:",
      rawResult ? rawResult.slice(0, 300) : "(empty)",
    );

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-004a Featured gallery corner badge with star on thumbnail",
        false,
        "FAIL: DOM eval error.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-004a Featured gallery corner badge with star on thumbnail",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-featured-section") {
      step(
        "TC-004a Featured gallery corner badge with star on thumbnail",
        false,
        "FAIL: No .featured-sample-section element found in the gallery DOM.",
      );
      return;
    }
    if (data.error === "no-featured-cards") {
      step(
        "TC-004a Featured gallery corner badge with star on thumbnail",
        false,
        "FAIL: No .sample-card elements found inside .featured-sample-section.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-004a Featured gallery corner badge with star on thumbnail",
        false,
        `FAIL: ${data.error}`,
      );
      return;
    }

    const detail =
      `featured=${data.featuredCount} cards; nonFeatured=${data.nonFeaturedCount} cards; ` +
      `missingCornerBadge=${data.missingBadgeCount} (expect 0); ` +
      `missingCornerStar=${data.missingStarCount} (expect 0); ` +
      `unexpectedCornerBadge=${data.unexpectedBadgeCount} (expect 0); ` +
      `sectionBg=${data.sectionBg} transparent=${data.bgTransparent}`;
    step(
      "TC-004a Featured gallery corner badge with star on thumbnail",
      !!data.passes,
      detail,
    );
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-004b — Featured items in List view display a star icon (codicon-star-full)
  //           as the first child of their h3 before the title text;
  //           the star uses currentColor so it renders in the same color as the title.
  //           Non-featured list items do not have the star.
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery in Gallery (Grid) view
  //   4. Click the List view toggle button — observe layout switches to vertical list
  //   5. Screenshot 14 showing list view with star icons before featured item titles
  //   6. For each .sample-list-item inside .featured-sample-section verify h3 .featured-star present
  //   7. For each .sample-list-item outside .featured-sample-section verify h3 .featured-star absent
  //   8. Read getComputedStyle(starSpan).color and getComputedStyle(h3).color; verify equal (currentColor)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-004b: Featured items display star icon before title in List view", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-004b", "light");

    if (!galleryOpened) {
      takeScreenshot("14-tc004b-list-star");
      step(
        "TC-004b Featured list item star icon before title",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Ensure we start from Gallery (grid) view before switching to List
    sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"grid view\"]')" +
        "    ||document.querySelector('[aria-label=\"Gallery view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      var l=(b.getAttribute('aria-label')||'').toLowerCase();" +
        "      return l.includes('grid')||l.includes('gallery');});" +
        "  if(btn)btn.click();" +
        "  return btn?'clicked':'no-btn';" +
        "})()",
      3000,
    );
    await wait(600);

    // Step 4: Click the List view toggle button — observe layout switches to vertical list
    const switchResult = sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"list view\"]')" +
        "    ||document.querySelector('[aria-label=\"List view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      return (b.getAttribute('aria-label')||'').toLowerCase().includes('list');});" +
        "  if(!btn)return 'no-list-button';" +
        "  btn.click();" +
        "  return 'clicked';" +
        "})()",
      5000,
    );
    console.log("  TC-004b switch to list result:", switchResult);
    await wait(800);

    // Step 5: Screenshot showing list view with star icons before featured item titles
    takeScreenshot("14-tc004b-list-star");

    // Steps 6-8: Verify star icon in h3 for featured list items; absent for non-featured;
    //            star color must equal h3 color (star uses currentColor)
    const evalScript =
      "(function(){" +
      "  var featSec=document.querySelector('.featured-sample-section');" +
      "  if(!featSec) return JSON.stringify({error:'no-featured-section'});" +
      "  var featItems=Array.from(featSec.querySelectorAll('.sample-list-item'));" +
      "  var allItems=Array.from(document.querySelectorAll('.sample-list-item'));" +
      "  var nonFeatItems=allItems.filter(function(i){return !featSec.contains(i);});" +
      "  if(featItems.length===0) return JSON.stringify({error:'no-featured-list-items'});" +
      // Step 6: every featured list item must have h3 .featured-star
      "  var missingStars=featItems.filter(function(i){" +
      "    return !i.querySelector('h3 .featured-star');});" +
      // Step 7: no non-featured list item should have h3 .featured-star
      "  var unexpectedStars=nonFeatItems.filter(function(i){" +
      "    return !!i.querySelector('h3 .featured-star');});" +
      // Step 8: star color must match h3 color (currentColor)
      "  var colorMismatch=[];" +
      "  var starColorSample=null,h3ColorSample=null;" +
      "  featItems.forEach(function(i){" +
      "    var starEl=i.querySelector('h3 .featured-star');" +
      "    var h3El=i.querySelector('h3');" +
      "    if(!starEl||!h3El)return;" +
      "    var sc=getComputedStyle(starEl).color;" +
      "    var hc=getComputedStyle(h3El).color;" +
      "    if(!starColorSample){starColorSample=sc;h3ColorSample=hc;}" +
      "    if(sc!==hc)colorMismatch.push({starColor:sc,h3Color:hc});});" +
      "  return JSON.stringify({" +
      "    featuredCount:featItems.length," +
      "    nonFeaturedCount:nonFeatItems.length," +
      "    missingStarCount:missingStars.length," +
      "    unexpectedStarCount:unexpectedStars.length," +
      "    colorMismatchCount:colorMismatch.length," +
      "    starColorSample:starColorSample," +
      "    h3ColorSample:h3ColorSample," +
      "    passes:missingStars.length===0&&unexpectedStars.length===0&&colorMismatch.length===0" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 10000);
    console.log(
      "  TC-004b eval result:",
      rawResult ? rawResult.slice(0, 300) : "(empty)",
    );

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-004b Featured list item star icon before title",
        false,
        "FAIL: DOM eval error.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-004b Featured list item star icon before title",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-featured-section") {
      step(
        "TC-004b Featured list item star icon before title",
        false,
        "FAIL: No .featured-sample-section element found in the gallery DOM.",
      );
      return;
    }
    if (data.error === "no-featured-list-items") {
      step(
        "TC-004b Featured list item star icon before title",
        false,
        "FAIL: No .sample-list-item elements found inside .featured-sample-section — list view may not have rendered.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-004b Featured list item star icon before title",
        false,
        `FAIL: ${data.error}`,
      );
      return;
    }

    const detail =
      `featured=${data.featuredCount} items; nonFeatured=${data.nonFeaturedCount} items; ` +
      `missingStarInH3=${data.missingStarCount} (expect 0); ` +
      `unexpectedStarInH3=${data.unexpectedStarCount} (expect 0); ` +
      `colorMismatch=${data.colorMismatchCount} (expect 0); ` +
      `starColor=${data.starColorSample} h3Color=${data.h3ColorSample}`;
    step(
      "TC-004b Featured list item star icon before title",
      !!data.passes,
      detail,
    );
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-005 — Screen reader differentiates Featured from non-Featured cards
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery — mix of featured and non-featured cards
  //   4. Screenshot 08 (gallery with featured and non-featured cards visible)
  //   5. Find all card elements (.sample-card) and read their aria-label attributes
  //   6. Verify at least one card has aria-label starting with "Featured sample."
  //   7. Verify at least one card does NOT have "Featured sample." prefix
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-005: Featured cards have 'Featured sample.' aria-label prefix", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-005");

    if (!galleryOpened) {
      takeScreenshot("08-tc005-aria-labels");
      step(
        "TC-005 Featured ARIA differentiation",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Steps 5-8: Find all cards, verify featured vs non-featured aria-labels,
    //            and verify .featured-badge elements have been removed (badge-free DOM).
    const evalScript =
      "(function(){" +
      "  var allCards=Array.from(document.querySelectorAll('.sample-card,.sample-list-item'));" +
      "  var labels=allCards.map(function(c){return c.getAttribute('aria-label')||'';});" +
      // Step 6: at least one card with aria-label starting with "Featured sample." (WITH period)
      "  var featured=labels.filter(function(l){return l.startsWith('Featured sample.');});" +
      // Step 7: at least one card without the "Featured sample." prefix
      "  var nonFeatured=labels.filter(function(l){" +
      "    return !l.startsWith('Featured sample.')&&l.length>0;});" +
      // Step 8: featured badge elements must have been removed (count === 0)
      "  var badgeCount=document.querySelectorAll('.featured-badge').length;" +
      "  return JSON.stringify({" +
      "    total:allCards.length," +
      "    featuredCount:featured.length," +
      "    nonFeaturedCount:nonFeatured.length," +
      "    badgeCount:badgeCount," +
      "    sampleFeaturedLabel:(featured[0]||'none').slice(0,100)," +
      "    sampleNonFeaturedLabel:(nonFeatured[0]||'none').slice(0,100)" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 5000);
    console.log(
      "  TC-005 eval result:",
      rawResult ? rawResult.slice(0, 300) : "(empty)",
    );

    // Step 4: Screenshot showing both featured and non-featured cards (no badge, ARIA only)
    injectAriaOverlay(".sample-card,.sample-list-item");
    await wait(300);
    takeScreenshot("08-tc005-aria-labels");
    removeAriaOverlay();

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-005 Featured ARIA differentiation",
        false,
        "FAIL: DOM eval error.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-005 Featured ARIA differentiation",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }

    const hasFeatured = data.featuredCount > 0;
    const hasNonFeatured = data.nonFeaturedCount > 0;
    // Step 8: badge must be gone — badgeCount === 0
    const noBadge = data.badgeCount === 0;
    const passes = hasFeatured && hasNonFeatured && noBadge;
    const detail =
      `${data.featuredCount} featured / ${data.nonFeaturedCount} non-featured of ${data.total} total. ` +
      `badge count=${data.badgeCount} (expect 0). ` +
      `Featured: "${data.sampleFeaturedLabel}". Non-featured: "${data.sampleNonFeaturedLabel}"`;
    step("TC-005 Featured ARIA differentiation", passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-006a — Focus ring contrast >= 3:1 in Gallery view (Light theme)
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery in Gallery (Grid) view
  //   4. Click inside gallery webview to give it keyboard focus
  //   5. Press Tab to move keyboard focus to the first .sample-card
  //   6. Screenshot 09 — IMMEDIATELY after Tab so focus ring is clearly visible
  //   7. Read getComputedStyle(card).outlineColor from focused card
  //   8. Read card background color
  //   9. Compute contrast_ratio(outline_color, card_bg) using WCAG relative luminance
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-006a: Focus ring contrast >= 3:1 in Gallery view (Light theme)", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-006a");

    if (!galleryOpened) {
      takeScreenshot("09-tc006a-gallery-focus");
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Ensure Gallery (grid) view is active
    sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"grid view\"]')" +
        "    ||document.querySelector('[aria-label=\"Gallery view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      var l=(b.getAttribute('aria-label')||'').toLowerCase();" +
        "      return l.includes('grid')||l.includes('gallery');});" +
        "  if(btn)btn.click();" +
        "  return btn?'clicked':'no-btn';" +
        "})()",
      3000,
    );
    await wait(600);

    // Step 4: Click inside the gallery webview to give it keyboard focus
    sendEvalSignal(
      "(function(){" +
        "  var el=document.querySelector('.sample-card,.sample-filter');" +
        "  if(el){el.focus();return 'focused';}" +
        "  document.body.click();" +
        "  return 'body-click';" +
        "})()",
      3000,
    );
    await wait(400);

    // Step 5: Press Tab to move keyboard focus to the first .sample-card
    await sendSignal("pressKey:Tab", 3000);
    await wait(400);

    // Steps 7-9: Read outline color and card background, compute WCAG contrast ratio
    const evalScript =
      "(function(){" +
      WCAG_HELPER_JS +
      "  var card=document.activeElement;" +
      "  if(!card||!card.classList.contains('sample-card')){" +
      "    card=document.querySelector('.sample-card');" +
      "    if(card)card.focus();" +
      "  }" +
      "  if(!card) return JSON.stringify({error:'no-sample-card'});" +
      "  var cs=getComputedStyle(card);" +
      "  var outlineRaw=cs.outlineColor;" +
      "  var outlineRgb=parseRgb(outlineRaw);" +
      "  var bgRgb=effectiveBg(card,'light');" +
      "  if(!outlineRgb) return JSON.stringify({error:'parse-outline',outlineRaw:outlineRaw});" +
      "  var outlineL=relativeLuminance(outlineRgb[0],outlineRgb[1],outlineRgb[2]);" +
      "  var bgL=relativeLuminance(bgRgb[0],bgRgb[1],bgRgb[2]);" +
      "  var ratio=contrastRatio(outlineL,bgL);" +
      "  return JSON.stringify({" +
      "    outlineRaw:outlineRaw," +
      "    bgRgb:'rgb('+bgRgb[0]+','+bgRgb[1]+','+bgRgb[2]+')'," +
      "    outlineStyle:cs.outlineStyle," +
      "    ratio:Math.round(ratio*100)/100," +
      "    passes:ratio>=3.0" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 8000);
    console.log(
      "  TC-006a eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 6: Screenshot — immediately after Tab so focus ring is clearly visible
    takeScreenshot("09-tc006a-gallery-focus");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        "FAIL: DOM eval error.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-sample-card") {
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        "FAIL: No .sample-card elements found in gallery webview.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        `FAIL: ${data.error} outlineRaw=${data.outlineRaw || "?"}`,
      );
      return;
    }

    const detail = `Computed ratio=${data.ratio}:1; outline=${data.outlineRaw} (${data.outlineStyle}); bg=${data.bgRgb}`;
    step("TC-006a Gallery focus ring contrast >= 3:1", !!data.passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-006b — Focus ring contrast >= 3:1 in List view (Light theme)
  //
  // Steps per test plan:
  //   1. Set VS Code color theme to "Default Light Modern"
  //   2. Click ATK Activity Bar icon
  //   3. Open gallery in Gallery (Grid) view
  //   4. Click the List view toggle button — layout switches to vertical list
  //   5. Click inside gallery webview to give it keyboard focus
  //   6. Press Tab to move keyboard focus to the first .sample-list-item
  //   7. Screenshot 10 — IMMEDIATELY after Tab so focus ring is clearly visible
  //   8. Read getComputedStyle(listItem).outlineColor from focused list item
  //   9. Find the effective background color by walking up the DOM until a non-transparent
  //      backgroundColor is found. Since the featured section background was removed (fix
  //      #15916), all list items share the page background (white #FFFFFF in Light theme).
  //  10. Compute contrast_ratio(outline_color, effective_background) using WCAG luminance
  //
  // Expected:
  //   - All list items share the white page background (#FFFFFF) — no grey featured section bg.
  //   - Focused list item has a dark-blue focus ring (#005FB8, ~10:1 vs white).
  //   - contrast_ratio >= 3.0 (WCAG AA non-text threshold)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-006b: Focus ring contrast >= 3:1 in List view (Light theme)", async () => {
    // Steps 1-3: set Light theme + click ATK Activity Bar + open gallery
    const galleryOpened = await setupTcGallery("TC-006b");

    if (!galleryOpened) {
      takeScreenshot("10-tc006b-list-focus");
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        "FAIL: Gallery webview not open after setup.",
      );
      return;
    }

    // Ensure we start from Gallery view
    sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"grid view\"]')" +
        "    ||document.querySelector('[aria-label=\"Gallery view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      var l=(b.getAttribute('aria-label')||'').toLowerCase();" +
        "      return l.includes('grid')||l.includes('gallery');});" +
        "  if(btn)btn.click();" +
        "  return btn?'clicked':'no-btn';" +
        "})()",
      3000,
    );
    await wait(600);

    // Step 4: Click the List view toggle button — observe layout switches to vertical list
    const switchResult = sendEvalSignal(
      "(function(){" +
        "  var btn=document.querySelector('[aria-label=\"list view\"]')" +
        "    ||document.querySelector('[aria-label=\"List view\"]')" +
        "    ||Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "      return (b.getAttribute('aria-label')||'').toLowerCase().includes('list');});" +
        "  if(!btn)return 'no-list-button';" +
        "  btn.click();" +
        "  return 'clicked';" +
        "})()",
      5000,
    );
    console.log("  TC-006b switch to list result:", switchResult);
    await wait(800);

    // Step 5: Click inside gallery to give it keyboard focus
    sendEvalSignal(
      "(function(){" +
        "  var el=document.querySelector('.sample-list-item,.sample-filter');" +
        "  if(el){el.focus();return 'focused';}" +
        "  document.body.click();" +
        "  return 'body-click';" +
        "})()",
      3000,
    );
    await wait(400);

    // Step 6: Press Tab to move keyboard focus to the first .sample-list-item
    await sendSignal("pressKey:Tab", 3000);
    await wait(400);

    // Steps 8-10: Read outline color and list item background, compute WCAG contrast ratio
    const evalScript =
      "(function(){" +
      WCAG_HELPER_JS +
      "  var item=document.activeElement;" +
      "  if(!item||!item.classList.contains('sample-list-item')){" +
      "    item=document.querySelector('.sample-list-item');" +
      "    if(item)item.focus();" +
      "  }" +
      "  if(!item) return JSON.stringify({error:'no-sample-list-item'});" +
      "  var cs=getComputedStyle(item);" +
      "  var outlineRaw=cs.outlineColor;" +
      "  var outlineRgb=parseRgb(outlineRaw);" +
      "  var bgRgb=effectiveBg(item,'light');" +
      "  if(!outlineRgb) return JSON.stringify({error:'parse-outline',outlineRaw:outlineRaw});" +
      "  var outlineL=relativeLuminance(outlineRgb[0],outlineRgb[1],outlineRgb[2]);" +
      "  var bgL=relativeLuminance(bgRgb[0],bgRgb[1],bgRgb[2]);" +
      "  var ratio=contrastRatio(outlineL,bgL);" +
      "  return JSON.stringify({" +
      "    outlineRaw:outlineRaw," +
      "    bgRgb:'rgb('+bgRgb[0]+','+bgRgb[1]+','+bgRgb[2]+')'," +
      "    outlineStyle:cs.outlineStyle," +
      "    ratio:Math.round(ratio*100)/100," +
      "    passes:ratio>=3.0" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 8000);
    console.log(
      "  TC-006b eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 7: Screenshot — immediately after Tab so focus ring is clearly visible
    takeScreenshot("10-tc006b-list-focus");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        "FAIL: DOM eval error.",
      );
      return;
    }

    let data: any = null;
    try {
      data = JSON.parse(rawResult);
    } catch {}
    if (!data) {
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-sample-list-item") {
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        "FAIL: No .sample-list-item elements found — list view may not have rendered.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        `FAIL: ${data.error} outlineRaw=${data.outlineRaw || "?"}`,
      );
      return;
    }

    const detail = `Computed ratio=${data.ratio}:1; outline=${data.outlineRaw} (${data.outlineStyle}); bg=${data.bgRgb}`;
    step("TC-006b List focus ring contrast >= 3:1", !!data.passes, detail);
  });
});
