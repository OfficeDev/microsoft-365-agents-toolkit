// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * sample-app-a11y.test.ts
 * TC-001:  Link text color contrast >= 4.5:1 when focused (Light theme)
 * TC-002:  Gallery/List toggle buttons aria-pressed state before and after click
 * TC-003:  Sample card accessible names include tags on keyboard focus
 * TC-004:  Featured badge non-text contrast >= 3:1 (WCAG 1.4.11)
 * TC-005:  Screen reader differentiates Featured from non-Featured cards
 * TC-006a: Focus ring contrast >= 3:1 in Gallery view (Light theme)
 * TC-006b: Focus ring contrast >= 3:1 in List view (Light theme)
 *
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
 * All contrast checks use the WCAG relative luminance formula (IEC 61966-2-1 sRGB).
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

async function waitForCommand(cmd: string, maxMs = 30000): Promise<boolean> {
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
  "function effectiveBg(el){" +
  "  var node=el;" +
  "  while(node){" +
  "    var cs=getComputedStyle(node);" +
  "    var bg=cs.backgroundColor;" +
  "    if(bg&&bg!=='transparent'&&bg!=='rgba(0, 0, 0, 0)'){" +
  "      var rgb=parseRgb(bg);" +
  "      if(rgb) return rgb;" +
  "    }" +
  "    node=node.parentElement;" +
  "  }" +
  "  return [255,255,255];" +
  "}";

/**
 * Injects a CSS style that renders each card\'s aria-label as an overlay via ::before.
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

suite("ATK Sample App A11y Regression Tests (Issue #15916)", function () {
  this.timeout(10 * 60 * 1000);

  const steps: object[] = [];
  let passed = 0;
  let failed = 0;
  let galleryOpened = false;

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

  // ── Screenshot 01: VS Code with ATK extension activated ─────────────────────
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
    // Screenshot 01: Baseline — extension loaded before gallery opens
    takeScreenshot("01-extension-active");
    if (!active && ext) {
      console.log(
        "  Note: Extension found but not active — source-level checks used below",
      );
    }
  });

  // ── Screenshot 02: Sample Gallery panel with sample cards ───────────────────
  test("Open Sample Gallery panel", async () => {
    const cmdName = "fx-extension.openSamples";
    // Poll 60s for command registration after extension activation
    const available = await waitForCommand(cmdName, 60000);
    step(
      "fx-extension.openSamples registered",
      available,
      available
        ? "command found"
        : "command not registered — source-level checks used for remaining TCs",
    );

    if (!available) {
      const allCmds = await vscode.commands.getCommands(true);
      const fxCmds = allCmds.filter((c) => c.startsWith("fx-extension"));
      console.log(
        `  [diag] fx-extension.* commands (${fxCmds.length}):`,
        fxCmds.slice(0, 20).join(", ") || "(none)",
      );
    }

    if (available) {
      // Fire without await — command opens a webview panel
      vscode.commands.executeCommand(cmdName).then(undefined, () => {});
      // Poll until gallery renders (.sample-filter or cards present)
      const galleryReadyScript =
        "JSON.stringify({" +
        "sampleCards:document.querySelectorAll(\".sample-card\").length," +
        "hasFilter:!!document.querySelector(\".sample-filter\")," +
        "hasOffline:!!document.querySelector(\".offlinePage\")," +
        "elemCount:document.querySelectorAll(\"*\").length" +
        "})";
      let pollCount = 0;
      let galleryLoaded = false;
      while (pollCount < 30 && !galleryLoaded) {
        await wait(2000);
        pollCount++;
        const diagResult = sendEvalSignal(galleryReadyScript, 3000);
        console.log(
          `  [poll ${pollCount}] Gallery status: ${diagResult.slice(0, 180)}`,
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
                `  Gallery loaded after ${pollCount * 2}s: ${status.sampleCards} cards`,
              );
            }
          } catch {}
        }
      }
      if (!galleryLoaded) {
        const bodyDump = sendEvalSignal(
          "JSON.stringify({href:window.location.href.slice(0,80),bodyText:document.body.innerText.slice(0,300)})",
          3000,
        );
        console.log("  [diag] DOM dump:", bodyDump.slice(0, 500));
      }
      galleryOpened = true;
    } else {
      vscode.commands
        .executeCommand("workbench.view.extension.teamsfx-toolkit")
        .then(undefined, () => {});
      await wait(2000);
    }

    // Screenshot 02: Sample Gallery panel with grid of sample cards
    takeScreenshot("02-gallery-open");
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-001 — Link text color contrast >= 4.5:1 when focused (Light theme)
  // Steps 1-6: open gallery → focus .ms-Link → compute WCAG contrast ratio
  // Screenshot 03: focused link with visible focus indicator
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-001: Link text color contrast >= 4.5:1 (focused, Light theme)", async () => {
    if (!galleryOpened) {
      takeScreenshot("03-tc001-link-focused");
      step(
        "TC-001 Link text contrast >= 4.5:1",
        false,
        "FAIL: Gallery webview not open.",
      );
      return;
    }

    // Step 3: Focus a .ms-Link element via element.focus()
    // Step 4: Read getComputedStyle(link).color from the focused link
    // Step 5: Compute contrast_ratio(link_color, background_color) via WCAG luminance
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
      "  var bg=effectiveBg(link);" +
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
      "  TC-001 eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 6: Screenshot showing the link in focused state with visible focus indicator
    takeScreenshot("03-tc001-link-focused");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-001 Link text contrast >= 4.5:1",
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
        "TC-001 Link text contrast >= 4.5:1",
        false,
        `FAIL: Could not parse eval result: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-ms-link") {
      step(
        "TC-001 Link text contrast >= 4.5:1",
        false,
        "FAIL: No .ms-Link elements found in gallery webview.",
      );
      return;
    }
    if (data.error) {
      step(
        "TC-001 Link text contrast >= 4.5:1",
        false,
        `FAIL: ${data.error} fg=${data.fgRaw || "?"}`,
      );
      return;
    }

    const detail =
      `Computed ratio=${data.ratio}:1; fg=${data.fgRaw}; bg=${data.bgRgb}; ` +
      `${data.count} .ms-Link elements`;
    step("TC-001 Link text contrast >= 4.5:1", !!data.passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-002 — Gallery/List toggle buttons expose aria-pressed before and after click
  // Steps 1-8: read initial state → screenshot → click List → verify toggled → screenshot
  // Screenshot 04: Gallery layout active (before click)
  // Screenshot 05: List layout active (after click)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-002: Gallery/List toggle aria-pressed state toggles correctly", async () => {
    if (!galleryOpened) {
      takeScreenshot("04-tc002-gallery-active");
      takeScreenshot("05-tc002-list-active");
      step(
        "TC-002 Toggle aria-pressed state",
        false,
        "FAIL: Gallery webview not open.",
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

    // Steps 2-3: Read aria-pressed on Gallery (expects "true") and List (expects "false")
    const beforeResult = sendEvalSignal(readPressedScript, 5000);
    console.log(
      "  TC-002 before-click state:",
      beforeResult ? beforeResult.slice(0, 200) : "(empty)",
    );

    // Step 4: Screenshot showing Gallery layout is active
    takeScreenshot("04-tc002-gallery-active");

    // Step 5: User clicks the List view toggle button
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

    // Steps 6-7: Verify List aria-pressed="true", Gallery aria-pressed="false"
    const afterResult = sendEvalSignal(readPressedScript, 5000);
    console.log(
      "  TC-002 after-click state:",
      afterResult ? afterResult.slice(0, 200) : "(empty)",
    );

    // Step 8: Screenshot showing List layout is active
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

    // Pass criteria from test plan:
    // Before click: Gallery aria-pressed="true", List aria-pressed="false"
    // After click:  List aria-pressed="true", Gallery aria-pressed="false"
    const beforeOk =
      galleryBefore?.pressed === "true" && listBefore?.pressed === "false";
    const afterOk =
      listAfter?.pressed === "true" && galleryAfter?.pressed === "false";
    const clickOk = clickResult === "clicked";

    const passes = beforeOk && afterOk && clickOk;
    const detail =
      `Before: gallery.pressed=${galleryBefore?.pressed} list.pressed=${listBefore?.pressed}. ` +
      `Click=${clickResult}. ` +
      `After: gallery.pressed=${galleryAfter?.pressed} list.pressed=${listAfter?.pressed}. ` +
      `Buttons: ${JSON.stringify(beforeData.buttons)}`;
    step("TC-002 Toggle aria-pressed state", passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-003 — Sample card accessible names include tags on keyboard focus
  // Steps 1-5: open gallery → press Tab → read aria-label → verify ". Tags:" pattern
  // Screenshot 06: focused card with focus ring visible (aria overlay)
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-003: Sample card aria-label includes tags on keyboard focus", async () => {
    if (!galleryOpened) {
      takeScreenshot("06-tc003-focused-card");
      step(
        "TC-003 Card aria-label includes tags",
        false,
        "FAIL: Gallery webview not open.",
      );
      return;
    }

    // Step 2: User presses Tab to move keyboard focus onto the first sample card
    await sendSignal("pressKey:Tab", 3000);
    await wait(400);

    // Step 3: Read aria-label from the focused card element
    // Step 4: Verify the aria-label matches pattern /.+\. Tags: .+/
    const evalScript =
      "(function(){" +
      "  var cards=Array.from(document.querySelectorAll('.sample-card,.sample-list-item'));" +
      "  var withTags=cards.filter(function(c){" +
      "    return /\\.\\s*Tags:/i.test(c.getAttribute('aria-label')||'');" +
      "  });" +
      "  var focused=document.activeElement;" +
      "  var focusedLabel=(focused&&(focused.classList.contains('sample-card')||" +
      "    focused.classList.contains('sample-list-item')))" +
      "    ?focused.getAttribute('aria-label')||'':'';" +
      "  var sampleLabel=withTags.length>0?(withTags[0].getAttribute('aria-label')||''):'';" +
      "  return JSON.stringify({" +
      "    total:cards.length," +
      "    withTags:withTags.length," +
      "    focusedLabel:focusedLabel.slice(0,120)," +
      "    sampleLabel:sampleLabel.slice(0,120)" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 5000);
    console.log(
      "  TC-003 eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 5: Screenshot showing the focused card with focus ring visible
    injectAriaOverlay(".sample-card,.sample-list-item");
    await wait(300);
    takeScreenshot("06-tc003-focused-card");
    removeAriaOverlay();

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

    const hasTagsInLabel = data.withTags > 0;
    const detail =
      `${data.withTags}/${data.total} cards have ". Tags:" in aria-label. ` +
      `Sample: "${(data.sampleLabel || data.focusedLabel || "").slice(0, 100)}"`;
    step("TC-003 Card aria-label includes tags", hasTagsInLabel, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-004 — Featured badge non-text contrast >= 3:1 against card background (WCAG 1.4.11)
  // Steps 1-6: open gallery → read badge bg color → read card bg → compute WCAG ratio
  // Screenshot 07: featured and non-featured cards side by side
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-004: Featured badge non-text contrast >= 3:1 (WCAG 1.4.11)", async () => {
    if (!galleryOpened) {
      takeScreenshot("07-tc004-featured-badge");
      step(
        "TC-004 Featured badge contrast >= 3:1",
        false,
        "FAIL: Gallery webview not open.",
      );
      return;
    }

    // Step 3: Read getComputedStyle(badge).backgroundColor from a .featured-badge element
    // Step 4: Read card background color from the parent .sample-card
    // Step 5: Compute contrast_ratio(badge_bg, card_bg) via WCAG luminance
    const evalScript =
      "(function(){" +
      WCAG_HELPER_JS +
      "  var badge=document.querySelector('.featured-badge');" +
      "  if(!badge){" +
      "    var fSec=document.querySelector('.featured-sample-section');" +
      "    return JSON.stringify({error:'no-badge',hasFeaturedSection:!!fSec});" +
      "  }" +
      "  var badgeCs=getComputedStyle(badge);" +
      "  var badgeBgRaw=badgeCs.backgroundColor;" +
      "  var card=badge.closest('.sample-card')||badge.parentElement;" +
      "  var cardBg=card?effectiveBg(card):[255,255,255];" +
      "  var badgeBgRgb=parseRgb(badgeBgRaw);" +
      "  if(!badgeBgRgb) return JSON.stringify({error:'parse-badge-bg',badgeBgRaw:badgeBgRaw});" +
      "  var badgeL=relativeLuminance(badgeBgRgb[0],badgeBgRgb[1],badgeBgRgb[2]);" +
      "  var cardL=relativeLuminance(cardBg[0],cardBg[1],cardBg[2]);" +
      "  var ratio=contrastRatio(badgeL,cardL);" +
      "  return JSON.stringify({" +
      "    badgeBgRaw:badgeBgRaw," +
      "    cardBgRgb:'rgb('+cardBg[0]+','+cardBg[1]+','+cardBg[2]+')'," +
      "    ratio:Math.round(ratio*100)/100," +
      "    passes:ratio>=3.0" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 8000);
    console.log(
      "  TC-004 eval result:",
      rawResult ? rawResult.slice(0, 200) : "(empty)",
    );

    // Step 6: Screenshot showing both featured and non-featured cards
    takeScreenshot("07-tc004-featured-badge");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step(
        "TC-004 Featured badge contrast >= 3:1",
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
        "TC-004 Featured badge contrast >= 3:1",
        false,
        `FAIL: Parse error: ${rawResult.slice(0, 100)}`,
      );
      return;
    }
    if (data.error === "no-badge") {
      step(
        "TC-004 Featured badge contrast >= 3:1",
        false,
        `FAIL: No .featured-badge element found. hasFeaturedSection=${data.hasFeaturedSection}`,
      );
      return;
    }
    if (data.error) {
      step(
        "TC-004 Featured badge contrast >= 3:1",
        false,
        `FAIL: ${data.error} badgeBgRaw=${data.badgeBgRaw || "?"}`,
      );
      return;
    }

    const detail =
      `Computed ratio=${data.ratio}:1; badge bg=${data.badgeBgRaw}; card bg=${data.cardBgRgb}`;
    step("TC-004 Featured badge contrast >= 3:1", !!data.passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-005 — Screen reader differentiates Featured from non-Featured cards
  // Steps 1-4: open gallery → read all card aria-labels → verify both prefixed and non-prefixed
  // Screenshot 08: gallery showing featured and non-featured cards
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-005: Featured cards have 'Featured sample.' aria-label prefix", async () => {
    if (!galleryOpened) {
      takeScreenshot("08-tc005-aria-labels");
      step(
        "TC-005 Featured ARIA differentiation",
        false,
        "FAIL: Gallery webview not open.",
      );
      return;
    }

    // Steps 2-4: Find all card elements, read aria-labels, verify both prefixed and non-prefixed exist
    const evalScript =
      "(function(){" +
      "  var allCards=Array.from(document.querySelectorAll('.sample-card,.sample-list-item'));" +
      "  var labels=allCards.map(function(c){return c.getAttribute('aria-label')||'';});" +
      "  var featured=labels.filter(function(l){return l.startsWith('Featured sample');});" +
      "  var nonFeatured=labels.filter(function(l){" +
      "    return !l.startsWith('Featured sample')&&l.length>0;});" +
      "  return JSON.stringify({" +
      "    total:allCards.length," +
      "    featuredCount:featured.length," +
      "    nonFeaturedCount:nonFeatured.length," +
      "    sampleFeaturedLabel:(featured[0]||'none').slice(0,100)," +
      "    sampleNonFeaturedLabel:(nonFeatured[0]||'none').slice(0,100)" +
      "  });" +
      "})()";

    const rawResult = sendEvalSignal(evalScript, 5000);
    console.log(
      "  TC-005 eval result:",
      rawResult ? rawResult.slice(0, 300) : "(empty)",
    );

    // Step 5: Screenshot
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
    const passes = hasFeatured && hasNonFeatured;
    const detail =
      `${data.featuredCount} featured / ${data.nonFeaturedCount} non-featured of ${data.total} total. ` +
      `Featured: "${data.sampleFeaturedLabel}". Non-featured: "${data.sampleNonFeaturedLabel}"`;
    step("TC-005 Featured ARIA differentiation", passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-006a — Focus ring contrast >= 3:1 in Gallery view (Light theme)
  // Steps 1-6: set Light theme → Tab to first .sample-card → compute WCAG contrast ratio
  // Screenshot 09: Gallery view with focused card and focus ring visible
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-006a: Focus ring contrast >= 3:1 in Gallery view (Light theme)", async () => {
    if (!galleryOpened) {
      takeScreenshot("09-tc006a-gallery-focus");
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        "FAIL: Gallery webview not open.",
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

    // Step 1: Set VS Code theme to a Light variant
    const wbConfig = vscode.workspace.getConfiguration("workbench");
    const originalTheme =
      wbConfig.get<string>("colorTheme") ?? "Default Dark Modern";
    try {
      await wbConfig.update(
        "colorTheme",
        "Default Light Modern",
        vscode.ConfigurationTarget.Global,
      );
    } catch {
      try {
        await wbConfig.update(
          "colorTheme",
          "Default Light+",
          vscode.ConfigurationTarget.Global,
        );
      } catch (e) {
        console.warn("  TC-006a: Could not switch to light theme:", e);
      }
    }
    await wait(3000);

    // Step 2: User presses Tab to move keyboard focus to the first .sample-card
    await sendSignal("pressKey:Tab", 3000);
    await wait(400);

    // Steps 3-5: Read outline color and card background, compute WCAG contrast ratio
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
      "  var isLight=document.body.classList.contains('vscode-light');" +
      "  var outlineRgb=parseRgb(outlineRaw);" +
      "  var bgRgb=effectiveBg(card);" +
      "  if(!outlineRgb) return JSON.stringify({error:'parse-outline',outlineRaw:outlineRaw,isLight:isLight});" +
      "  var outlineL=relativeLuminance(outlineRgb[0],outlineRgb[1],outlineRgb[2]);" +
      "  var bgL=relativeLuminance(bgRgb[0],bgRgb[1],bgRgb[2]);" +
      "  var ratio=contrastRatio(outlineL,bgL);" +
      "  return JSON.stringify({" +
      "    outlineRaw:outlineRaw," +
      "    bgRgb:'rgb('+bgRgb[0]+','+bgRgb[1]+','+bgRgb[2]+')'," +
      "    isLight:isLight," +
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

    // Step 6: Screenshot — after Tab so focus ring is clearly visible
    takeScreenshot("09-tc006a-gallery-focus");

    // Restore original theme
    try {
      await wbConfig.update(
        "colorTheme",
        originalTheme,
        vscode.ConfigurationTarget.Global,
      );
    } catch (e) {
      console.warn("  TC-006a: Could not restore theme:", e);
    }

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
    if (data.error) {
      step(
        "TC-006a Gallery focus ring contrast >= 3:1",
        false,
        `FAIL: ${data.error}. outlineRaw=${data.outlineRaw || "?"} isLight=${data.isLight}`,
      );
      return;
    }

    const detail =
      `Computed ratio=${data.ratio}:1; outline=${data.outlineRaw} (${data.outlineStyle}); ` +
      `bg=${data.bgRgb}; isLight=${data.isLight}`;
    step("TC-006a Gallery focus ring contrast >= 3:1", !!data.passes, detail);
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // TC-006b — Focus ring contrast >= 3:1 in List view (Light theme)
  // Steps 1-7: set Light theme → click List toggle → Tab to first .sample-list-item →
  //            compute WCAG contrast ratio
  // Screenshot 10: List view with focused list item and focus ring visible
  // ─────────────────────────────────────────────────────────────────────────────
  test("TC-006b: Focus ring contrast >= 3:1 in List view (Light theme)", async () => {
    if (!galleryOpened) {
      takeScreenshot("10-tc006b-list-focus");
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        "FAIL: Gallery webview not open.",
      );
      return;
    }

    // Step 1: Set VS Code theme to a Light variant
    const wbConfig = vscode.workspace.getConfiguration("workbench");
    const originalTheme =
      wbConfig.get<string>("colorTheme") ?? "Default Dark Modern";
    try {
      await wbConfig.update(
        "colorTheme",
        "Default Light Modern",
        vscode.ConfigurationTarget.Global,
      );
    } catch {
      try {
        await wbConfig.update(
          "colorTheme",
          "Default Light+",
          vscode.ConfigurationTarget.Global,
        );
      } catch (e) {
        console.warn("  TC-006b: Could not switch to light theme:", e);
      }
    }
    await wait(3000);

    // Step 2: User clicks the List view toggle button to switch to List layout
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
    console.log("  TC-006b list-click result:", switchResult);
    await wait(800);

    // Step 3: User presses Tab to move keyboard focus to the first .sample-list-item
    await sendSignal("pressKey:Tab", 3000);
    await wait(400);

    // Steps 4-6: Read outline color and list item background, compute WCAG contrast ratio
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
      "  var isLight=document.body.classList.contains('vscode-light');" +
      "  var outlineRgb=parseRgb(outlineRaw);" +
      "  var bgRgb=effectiveBg(item);" +
      "  if(!outlineRgb) return JSON.stringify({error:'parse-outline',outlineRaw:outlineRaw,isLight:isLight});" +
      "  var outlineL=relativeLuminance(outlineRgb[0],outlineRgb[1],outlineRgb[2]);" +
      "  var bgL=relativeLuminance(bgRgb[0],bgRgb[1],bgRgb[2]);" +
      "  var ratio=contrastRatio(outlineL,bgL);" +
      "  return JSON.stringify({" +
      "    outlineRaw:outlineRaw," +
      "    bgRgb:'rgb('+bgRgb[0]+','+bgRgb[1]+','+bgRgb[2]+')'," +
      "    isLight:isLight," +
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

    // Step 7: Screenshot — after Tab so focus ring is clearly visible
    takeScreenshot("10-tc006b-list-focus");

    // Restore original theme and gallery view
    try {
      await wbConfig.update(
        "colorTheme",
        originalTheme,
        vscode.ConfigurationTarget.Global,
      );
    } catch (e) {
      console.warn("  TC-006b: Could not restore theme:", e);
    }
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
    if (data.error) {
      step(
        "TC-006b List focus ring contrast >= 3:1",
        false,
        `FAIL: ${data.error}. outlineRaw=${data.outlineRaw || "?"} isLight=${data.isLight}; switchResult=${switchResult}`,
      );
      return;
    }

    const detail =
      `Computed ratio=${data.ratio}:1; outline=${data.outlineRaw} (${data.outlineStyle}); ` +
      `bg=${data.bgRgb}; isLight=${data.isLight}; switchResult=${switchResult}`;
    step("TC-006b List focus ring contrast >= 3:1", !!data.passes, detail);
  });

  test("Final state", async () => {
    await wait(1000);
    takeScreenshot("11-final-state");
    step("Final state captured", true);
  });
});
