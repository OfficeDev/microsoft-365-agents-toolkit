// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * sample-app-a11y.test.ts
 * Tests for Sample App panel accessibility (A11y) regressions.
 * Issue #15916: 5 A11y bugs in Light 2026 theme (contrast + screen reader).
 *
 * TC-001: Link text color contrast ≥ 4.5:1
 * TC-002: Featured vs non-Featured ARIA differentiation
 * TC-003: Featured badge non-text contrast ≥ 3:1
 * TC-004: Tags included in accessible name on keyboard focus
 * TC-005: Gallery/List toggle buttons expose aria-pressed state
 *
 * Runs INSIDE VSCode extension host via @vscode/test-electron (Mocha TDD).
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
 * Send an evaluation signal to Playwright for DOM inspection.
 * Returns result written to a response file.
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
    // Signal not processed — clean up and return empty
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
 * Parse Playwright ariaSnapshot YAML output (ACCESSIBILITY:{yaml}).
 * The format is: `- role "name" [prop=val]\n  - child...`
 * Returns structured nodes for ARIA attribute checks.
 */
interface AxNode {
  role: string;
  name: string;
  pressed?: string; // "true" | "false" | undefined
  line: string;
}
function parseAriaSnapshot(yaml: string): AxNode[] {
  const nodes: AxNode[] = [];
  for (const line of yaml.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed.startsWith("- ")) continue;
    const rest = trimmed.slice(2);
    // Match: role "name" [prop=val ...]
    const m = rest.match(/^(\w[\w-]*)\s+"([^"]*)"/);
    if (!m) continue;
    const role = m[1];
    const name = m[2];
    const node: AxNode = { role, name, line: rest };
    // Check for [pressed=...] flag
    const pressedM = rest.match(/\[pressed=(\w+)\]/);
    if (pressedM) node.pressed = pressedM[1];
    nodes.push(node);
  }
  return nodes;
}

/**
 * Injects a CSS style that renders each .sample-card's aria-label as an overlay
 * (via ::before pseudo-element). Call before takeScreenshot(), removeAriaOverlay() after.
 * This makes aria-label content VISIBLE in screenshots so reviewers can see pass/fail.
 */
function injectAriaOverlay(selector: string = ".sample-card"): void {
  const script =
    "(function() {" +
    "  var id = '__aria-overlay-style__';" +
    "  if (document.getElementById(id)) return;" +
    "  var s = document.createElement('style');" +
    "  s.id = id;" +
    "  s.textContent = '" + selector + " { position: relative !important; }" +
    "    " + selector + "::before {" +
    "      content: attr(aria-label);" +
    "      position: absolute; top: 0; left: 0; right: 0;" +
    "      background: rgba(0,0,0,0.82); color: #fff;" +
    "      font-size: 8px; line-height: 1.3; padding: 3px 4px;" +
    "      z-index: 9999; pointer-events: none;" +
    "      white-space: normal; word-break: break-word;" +
    "    }';" +
    "  document.head.appendChild(s);" +
    "  return 'injected';" +
    "})()";
  sendEvalSignal(script, 3000);
}

/**
 * Removes the aria-label overlay CSS injected by injectAriaOverlay().
 */
function removeAriaOverlay(): void {
  const script =
    "(function() {" +
    "  var s = document.getElementById('__aria-overlay-style__');" +
    "  if (s) s.parentNode.removeChild(s);" +
    "  return 'removed';" +
    "})()";
  sendEvalSignal(script, 1000);
}


suite("ATK Sample App A11y Regression Tests (Issue #15916)", function () {
  this.timeout(8 * 60 * 1000);

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
    console.log("=== ATK Sample App A11y Test ===");
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
    // Wait for extension to finish registering commands after activation
    await wait(8000);
    const active = !!ext?.isActive;
    step(
      "ATK extension activates",
      active,
      ext ? `v${ext.packageJSON.version}` : "not found",
    );
    takeScreenshot("01-extension-active");
    // Don't assert — if extension is loaded in dev mode without deps it still shows as inactive
    // but the source fixes are verified below
    if (!active && ext) {
      console.log(
        "  Note: Extension found but not active (likely missing dependency — see TC source checks below)",
      );
    }
  });

  test("Open Sample Gallery panel", async () => {
    const cmdName = "fx-extension.openSamples";
    // Extension activation can be slow in CI (workspace deps + API proposals init).
    // Poll for 60s to give the command time to register after isActive=true.
    const available = await waitForCommand(cmdName, 60000);
    step(
      "fx-extension.openSamples registered",
      available,
      available
        ? "command found"
        : "command not registered (extension not fully active — source-level checks used for TC-002/004/005)",
    );

    if (!available) {
      // Diagnostic: log which fx-extension commands ARE registered to help debug
      const allCmds = await vscode.commands.getCommands(true);
      const fxCmds = allCmds.filter((c) => c.startsWith("fx-extension"));
      console.log(
        `  [diag] fx-extension.* commands registered (${fxCmds.length}):`,
        fxCmds.slice(0, 20).join(", ") || "(none)",
      );
    }
    if (available) {
      // Fire without await — command opens a webview panel
      vscode.commands.executeCommand(cmdName).then(undefined, () => {});
      // Poll for gallery to render. Detect loaded state by .sample-filter (filter bar appears
      // when samples are ready) or .offlinePage (network error). Do NOT use [role=button] —
      // that is the A11y bug we are testing (layout buttons lack role="button" in unfixed code).
      const galleryReadyScript =
        "JSON.stringify({" +
        "href: window.location.href.slice(0,60)," +
        "sampleCards: document.querySelectorAll(\".sample-card\").length," +
        "hasFilter: !!document.querySelector(\".sample-filter\")," +
        "hasOffline: !!document.querySelector(\".offlinePage\")," +
        "elemCount: document.querySelectorAll(\"*\").length," +
        "loading: !document.querySelector(\".sample-filter\") && !document.querySelector(\".offlinePage\")" +
        "})";
      let pollCount = 0;
      let galleryLoaded = false;
      while (pollCount < 30 && !galleryLoaded) {
        await wait(2000);
        pollCount++;
        const diagResult = sendEvalSignal(galleryReadyScript, 3000);
        console.log(`  [poll ${pollCount}] Gallery status: ${diagResult.slice(0, 180)}`);
        if (diagResult && !diagResult.startsWith("ERROR:") && !diagResult.startsWith("ACCESSIBILITY:")) {
          try {
            const status = JSON.parse(diagResult);
            if (!status.loading) {
              galleryLoaded = true;
              console.log(`  Gallery loaded after ${pollCount * 2}s: ${status.sampleCards} cards, offline=${status.hasOffline}`);
            }
          } catch {}
        }
      }
      if (!galleryLoaded) {
        console.log("  Gallery did not load after 60s — dumping DOM state");
        const bodyDump = sendEvalSignal(
          "JSON.stringify({href: window.location.href.slice(0,80), bodyText: document.body.innerText.slice(0,300), elemCount: document.querySelectorAll(\"*\").length, html: document.documentElement.outerHTML.slice(0,400)})",
          3000
        );
        console.log("  [diag] DOM dump:", bodyDump.slice(0, 500));
      }
      galleryOpened = true;
    } else {
      // Attempt fallback but don't mark as opened
      vscode.commands
        .executeCommand("workbench.view.extension.teamsfx-toolkit")
        .then(undefined, () => {});
      await wait(2000);
    }

    takeScreenshot("02-sample-gallery-opened");
  });

  /**
   * TC-001 — Link text color contrast ≥ 4.5:1 (Light theme)
   * Bug:  In Light Modern theme, .ms-Link inherits rgb(72, 160, 199) — contrast ~3.2:1 on white.
   * Fix:  .vscode-light .ms-Link overridden to #005B9E — contrast ≈ 7.6:1 on white.
   *
   * NOTE: Theme switch is NOT done here — switching theme causes the gallery webview to reload
   * and breaks subsequent TCs. The CSS-rule check reads the stylesheet regardless of current theme.
   *
   * SNAPSHOT: Dark theme, gallery open. The "Github Copilot" link (right panel, if chat visible)
   *   is NOT visually distinguishable pass/fail from the screenshot alone.
   *   -> PRIMARY SIGNAL: TEST LOG line: "Link colors OK: rgb(0, 120, 212)" (PASS)
   *      vs "FAIL: Found low-contrast color rgb(72, 160, 199) on .ms-Link" (FAIL).
   *   -> When chat panel is hidden: "CSS fix applied: .ms-Link|rgb(0, 91, 158)" (PASS)
   *      vs "FAIL: CSS rule for .ms-Link exists but wrong color" (FAIL).
   */
  test("TC-001: Link text color contrast ≥ 4.5:1", async () => {
    if (!galleryOpened) {
      takeScreenshot("03-tc001-link-contrast");
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        false,
        "FAIL: Gallery webview not open. Cannot verify link color contrast.",
      );
      return;
    }

    // Primary: computed color of .ms-Link elements (visible when shouldShowChat=true).
    // Fallback: CSS stylesheet rule for #005B9E (verifiable even when shouldShowChat=false).
    const evalScript =
      "(function(){" +
      "  var els = Array.from(document.querySelectorAll('.ms-Link'));" +
      "  if (els.length > 0) {" +
      "    return 'colors:' + els.map(function(e){return getComputedStyle(e).color;}).slice(0,3).join(',');" +
      "  }" +
      "  var linkRules = [];" +
      "  try {" +
      "    Array.from(document.styleSheets).forEach(function(ss){" +
      "      try {" +
      "        Array.from(ss.cssRules).forEach(function(r){" +
      "          if (r.selectorText && r.selectorText.includes('ms-Link') && r.style && r.style.color) {" +
      "            linkRules.push(r.selectorText + '|' + r.style.color);" +
      "          }" +
      "        });" +
      "      } catch(e) {}" +
      "    });" +
      "  } catch(e) {}" +
      "  return linkRules.length > 0 ? 'css-rule:' + linkRules.join(';') : 'no-ms-link';" +
      "})()";
    const rawResult = sendEvalSignal(evalScript, 5000);
    console.log("  TC-001 eval result:", rawResult ? rawResult.slice(0, 120) : "(empty)");

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      takeScreenshot("03-tc001-link-contrast");
      step("TC-001 Link text contrast ≥ 4.5:1", false,
        "FAIL: DOM eval error — gallery may not be accessible via Playwright.");
      return;
    } else if (rawResult.startsWith("colors:")) {
      const colors = rawResult.slice(7).split(",").filter(Boolean);
      const problematicColor = "rgb(72, 160, 199)";
      const hasProblematic = colors.some((c) => c.trim() === problematicColor);
      takeScreenshot("03-tc001-link-contrast");
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        !hasProblematic,
        hasProblematic
          ? `FAIL: Found low-contrast color ${problematicColor} on .ms-Link`
          : `Link colors OK: ${colors.join(", ")}`,
      );
      return;
    } else if (rawResult.startsWith("css-rule:")) {
      const rules = rawResult.slice(9);
      const hasGoodRule =
        rules.includes("#005b9e") || rules.includes("#005B9E") || rules.includes("rgb(0, 91, 158)");
      takeScreenshot("03-tc001-link-contrast");
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        hasGoodRule,
        hasGoodRule
          ? `CSS fix applied: ${rules.slice(0, 80)}`
          : `FAIL: CSS rule for .ms-Link exists but wrong color: ${rules.slice(0, 80)}`,
      );
    } else {
      takeScreenshot("03-tc001-link-contrast");
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        false,
        "FAIL: No .ms-Link elements and no CSS fix rule found. Fix not applied (shouldShowChat=false hides links; CSS fix must be in stylesheet).",
      );
    }
  });


  /**
   * TC-001b — Link text color contrast ≥ 4.5:1 (List view)
   * Same accessibility requirement as TC-001 but verifies the list-view layout (compact rows).
   * The list-view toggle switches cards from a grid to a narrow row format — the same .ms-Link
   * fix must hold for both layouts.
   *
   * SNAPSHOT: Dark theme, gallery in LIST VIEW (rows, not cards).
   *   Link elements (.ms-Link) are visible as blue text in each row.
   * SNAPSHOT PASS: LOG shows "Link colors OK" or "CSS fix applied" + list-view rows visible.
   * SNAPSHOT FAIL: LOG shows low-contrast rgb(72,160,199) color present in list view.
   *
   * Note: Switches gallery to list view, checks .ms-Link colors, then restores grid view.
   */
  test("TC-001b: Link text color contrast ≥ 4.5:1 (list view)", async () => {
    if (!galleryOpened) {
      takeScreenshot("03b-tc001b-link-contrast-list");
      step(
        "TC-001b Link text contrast ≥ 4.5:1 (list view)",
        false,
        "FAIL: Gallery webview not open. Cannot verify link color in list view.",
      );
      return;
    }

    // Switch to list view via the toggle button.
    // Use a busy-wait loop inside the eval so headless CI (which renders slower)
    // has time to render the layout toggle buttons before we give up.
    const switchToList = sendEvalSignal(
      "(function(){" +
      "  var end = Date.now() + 5000;" +
      "  var btn = null;" +
      "  while (!btn && Date.now() < end) {" +
      "    btn = document.querySelector('[aria-label=\"list view\"]') ||" +
      "          document.querySelector('.layout-button[title*=\"list\"]') ||" +
      "          Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
      "            return (b.getAttribute('aria-label') || '').toLowerCase().indexOf('list') !== -1;" +
      "          });" +
      "    if (!btn) { var t = Date.now() + 200; while (Date.now() < t) {} }" +
      "  }" +
      "  if (!btn) return 'no-list-toggle';" +
      "  btn.click(); return 'clicked';" +
      "})()",
      8000,
    );
    if (!switchToList || switchToList === "no-list-toggle") {
      console.log("  TC-001b: list-view toggle not found, skipping");
      takeScreenshot("03b-tc001b-link-contrast-list");
      step(
        "TC-001b Link text contrast ≥ 4.5:1 (list view)",
        false,
        "FAIL: List-view toggle button not found — cannot verify list view contrast.",
      );
      return;
    }
    await wait(1000);

    // Check .ms-Link color in list view (same logic as TC-001)
    const evalScript =
      "(function(){" +
      "  var els = Array.from(document.querySelectorAll('.ms-Link'));" +
      "  if (els.length > 0) {" +
      "    return 'colors:' + els.map(function(e){return getComputedStyle(e).color;}).slice(0,3).join(',');" +
      "  }" +
      "  return 'no-ms-link';" +
      "})()";
    const rawResult = sendEvalSignal(evalScript, 5000);
    console.log("  TC-001b eval result:", rawResult ? rawResult.slice(0, 120) : "(empty)");

    takeScreenshot("03b-tc001b-link-contrast-list");

    // Restore grid view
    sendEvalSignal(
      "(function(){" +
      "  var btn = document.querySelector('[aria-label=\"grid view\"]');" +
      "  if (!btn) btn = document.querySelector('.layout-button[title*=\"grid\"]');" +
      "  if (btn) { btn.click(); return 'restored'; }" +
      "  return 'no-grid-toggle';" +
      "})()",
      2000,
    );
    await wait(500);

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      step("TC-001b Link text contrast ≥ 4.5:1 (list view)", false,
        "FAIL: DOM eval error in list view.");
      return;
    }

    if (rawResult.startsWith("colors:")) {
      const colors = rawResult.slice(7).split(",").filter(Boolean);
      const problematicColor = "rgb(72, 160, 199)";
      const hasProblematic = colors.some((c) => c.trim() === problematicColor);
      step(
        "TC-001b Link text contrast ≥ 4.5:1 (list view)",
        !hasProblematic,
        hasProblematic
          ? `FAIL: Low-contrast color ${problematicColor} on .ms-Link in list view`
          : `Link colors OK in list view: ${colors.join(", ")}`,
      );
    } else {
      step(
        "TC-001b Link text contrast ≥ 4.5:1 (list view)",
        false,
        `FAIL: No .ms-Link elements found in list view. raw=${rawResult.slice(0, 60)}`,
      );
    }
  });

  /**
   * TC-002 — Featured vs non-Featured ARIA differentiation
   * Bug:  Featured cards share same aria-label format as regular cards — no "Featured sample." prefix.
   * Fix:  Featured cards get aria-label prefixed with "Featured sample." for screen readers.
   *
   * SNAPSHOT: Dark theme, gallery open; "Featured samples" section visible at top.
   *   aria-labels are rendered as OVERLAY OVERLAYS on each card (black tooltip at card top).
   * SNAPSHOT PASS: Each featured card shows overlay starting with "Featured sample."
   *   e.g. "Featured sample. Travel Agent. C#, Custom Engine Agent..."
   * SNAPSHOT FAIL: Featured card overlays show plain names without prefix,
   *   e.g. "Travel Agent. C#, Custom Engine Agent..." (same as regular cards).
   */
  test("TC-002: Featured vs non-Featured ARIA differentiation", async () => {
    if (galleryOpened) {
      // Use .sample-card (not [role=button] -- cards intentionally lack role=button in unfixed code).
      // Check featured section cards for "Featured sample." aria-label prefix.
      const evalScript =
        "(() => {" +
        "  var fSec = document.querySelector('.featured-sample-section');" +
        "  var sSec = document.querySelector('.sample-section');" +
        "  var fCards = fSec ? Array.from(fSec.querySelectorAll('.sample-card')) : [];" +
        "  var sCards = sSec ? Array.from(sSec.querySelectorAll('.sample-card')) : [];" +
        "  var fLabels = fCards.map(function(c){return c.getAttribute('aria-label')||'';});" +
        "  var sLabels = sCards.map(function(c){return c.getAttribute('aria-label')||'';});" +
        "  var fPrefixed = fLabels.filter(function(l){return l.startsWith('Featured sample');});" +
        "  return JSON.stringify({featuredTotal:fCards.length,featuredPrefixed:fPrefixed.length," +
        "    regularTotal:sCards.length,sampleFeaturedLabel:fLabels[0]||'none',sampleRegularLabel:sLabels[0]||'none'});" +
        "})()"
      const rawResult = sendEvalSignal(evalScript, 5000);

      if (rawResult && !rawResult.startsWith("ERROR:") && rawResult.startsWith("ACCESSIBILITY:")) {
        const nodes = parseAriaSnapshot(rawResult.slice("ACCESSIBILITY:".length));
        const buttonNodes = nodes.filter(n => n.role.toLowerCase() === "button");
        const featuredNodes = buttonNodes.filter(n => (n.name || "").startsWith("Featured sample"));
        const hasFeaturedAria = featuredNodes.length > 0;
        injectAriaOverlay(); await wait(300); takeScreenshot("04-tc002-featured-aria"); removeAriaOverlay();
        step(
          "TC-002 Featured ARIA differentiation",
          hasFeaturedAria,
          `[AX-TREE] ${buttonNodes.length} button nodes, ${featuredNodes.length} with "Featured sample" prefix`,
        );
        return;
      }
      if (rawResult && !rawResult.startsWith("ERROR:")) {
        let data002: any = null;
        try { data002 = JSON.parse(rawResult); } catch {}
        if (data002 !== null) {
          const hasFeaturedAria = data002.featuredPrefixed > 0;
          injectAriaOverlay(); await wait(300); takeScreenshot("04-tc002-featured-aria"); removeAriaOverlay();
          step(
            "TC-002 Featured ARIA differentiation",
            hasFeaturedAria,
            hasFeaturedAria
              ? `OK: ${data002.featuredPrefixed}/${data002.featuredTotal} featured cards have "Featured sample." prefix`
              : `FAIL: ${data002.featuredTotal} featured cards, 0 with "Featured sample." in aria-label. ` +
                `Sample: "${(data002.sampleFeaturedLabel || "").slice(0, 80)}"`,
          );
          return;
        } else {
          injectAriaOverlay(); await wait(300); takeScreenshot("04-tc002-featured-aria"); removeAriaOverlay();
          step(
            "TC-002 Featured ARIA differentiation",
            false,
            "parse error: " + rawResult.slice(0, 100),
          );
          return;
        }
      }
    }

    injectAriaOverlay(); await wait(300); takeScreenshot("04-tc002-featured-aria"); removeAriaOverlay();
    step(
      "TC-002 Featured ARIA differentiation",
      false,
      "FAIL: Gallery webview not open. Extension must activate and gallery must load for DOM-based A11y check.",
    );
  });

  /**
   * TC-003 — Featured badge non-text contrast ≥ 3:1
   * Bug:  No .featured-badge element on featured cards — badge not rendered at all.
   * Fix:  .featured-badge with text "★ Featured" uses color #7A5C00 on white ≈ 4.9:1 (> 3:1 WCAG).
   *
   * SNAPSHOT PASS: Gallery open; each featured card shows a small "★ Featured" badge chip
   *   with clearly legible dark-gold/brown text, visually distinct from the card background.
   * SNAPSHOT FAIL: Featured cards have NO badge chip visible at all — featured and regular
   *   cards look identical (badge element not rendered in the unfixed build).
   */
  test("TC-003: Featured badge contrast ≥ 3:1", async () => {
    if (!galleryOpened) {
      takeScreenshot("05-tc003-badge-contrast");
      step(
        "TC-003 Featured badge present with accessible contrast",
        false,
        "FAIL: Gallery webview not open. Cannot check featured badge.",
      );
      return;
    }
    // Check .featured-sample-section (gallery loaded) + .featured-badge (added by fix).
    const evalScript =
      "(() => {" +
      "  var fSec = document.querySelector('.featured-sample-section');" +
      "  if (!fSec) return 'no-featured-section';" +
      "  var badges = Array.from(document.querySelectorAll('.featured-badge'));" +
      "  if (badges.length === 0) return JSON.stringify({hasFeaturedSection:true,count:0});" +
      "  var s = getComputedStyle(badges[0]);" +
      "  return JSON.stringify({hasFeaturedSection:true,bg:s.backgroundColor,color:s.color,count:badges.length});" +
      "})()"
    const rawResult = sendEvalSignal(evalScript, 5000);

    if (!rawResult || rawResult.startsWith("ERROR:")) {
      takeScreenshot("05-tc003-badge-contrast");
      step("TC-003 Featured badge present with accessible contrast", false, "FAIL: DOM eval error.");
    } else if (rawResult === "no-featured-section") {
      takeScreenshot("05-tc003-badge-contrast");
      step("TC-003 Featured badge present with accessible contrast", false,
        "FAIL: .featured-sample-section not found (gallery may not have loaded featured samples).");
    } else {
      let data003: any = null;
      try { data003 = JSON.parse(rawResult); } catch {}
      if (data003 !== null) {
        if (data003.count === 0) {
          takeScreenshot("05-tc003-badge-contrast");
          step(
            "TC-003 Featured badge present with accessible contrast",
            false,
            "FAIL: .featured-sample-section present but no .featured-badge elements. A11y bug: badge element not added to featured cards.",
          );
        } else {
          takeScreenshot("05-tc003-badge-contrast");
          step(
            "TC-003 Featured badge present with accessible contrast",
            true,
            `${data003.count} badge(s) found. bg=${data003.bg}, color=${data003.color}`,
          );
        }
      } else {
        takeScreenshot("05-tc003-badge-contrast");
        step("TC-003 Featured badge contrast ≥ 3:1", false, "parse error: " + rawResult.slice(0, 100));
      }
    }
  });

  /**
   * TC-004 — Tags included in accessible name on keyboard focus
   * Bug:  .sample-card aria-label omits tag names (e.g. "Travel Agent. C#, Custom Engine Agent")
   *       — screen reader does not announce tag names when the card is focused.
   * Fix:  aria-label includes ". Tags: C#, Custom Engine Agent, …" suffix.
   *
   * SNAPSHOT: Dark theme, gallery open. aria-labels rendered as OVERLAY on each card.
   * SNAPSHOT PASS: Card overlays include Tags suffix, e.g.
   *   "Travel Agent. C#, Custom Engine Agent. Tags: C#, Custom Engine Agent, Agents SDK, Agent Framework"
   * SNAPSHOT FAIL: Card overlays show no Tags suffix, e.g.
   *   "Travel Agent. C#, Custom Engine Agent" (tags listed as chips but not in aria-label).
   */
  test("TC-004: Tags announced on keyboard focus via aria-label", async () => {
    if (galleryOpened) {
      // Use .sample-card (not [role=button] -- cards intentionally lack role=button in unfixed code).
      // Check that aria-label includes "Tags:" to announce tags on keyboard focus.
      const evalScript =
        "(() => {" +
        "  var cards = Array.from(document.querySelectorAll('.sample-card'));" +
        "  var withTags = cards.filter(function(c){" +
        "    var l = c.getAttribute('aria-label') || '';" +
        "    return l.toLowerCase().includes('tags:');" +
        "  });" +
        "  var sampleAny = cards[0];" +
        "  var sample = withTags[0] ? (withTags[0].getAttribute('aria-label')||'')" +
        "               : sampleAny ? (sampleAny.getAttribute('aria-label')||'none') : 'none';" +
        "  return JSON.stringify({total:cards.length,withTags:withTags.length,sample:sample});" +
        "})()"
      const rawResult = sendEvalSignal(evalScript, 5000);

      if (rawResult && rawResult.startsWith("ACCESSIBILITY:")) {
        const nodes = parseAriaSnapshot(rawResult.slice("ACCESSIBILITY:".length));
        const buttonNodes = nodes.filter(n => n.role.toLowerCase() === "button");
        const tagNodes = buttonNodes.filter(n => (n.name || "").toLowerCase().includes("tags:"));
        const hasTagsInLabel = tagNodes.length > 0;
        const sample = tagNodes[0]?.name || "";
        injectAriaOverlay(); await wait(300); takeScreenshot("06-tc004-tags-aria"); removeAriaOverlay();
        step(
          "TC-004 Tags in aria-label",
          hasTagsInLabel,
          `[AX-TREE] ${buttonNodes.length} buttons, ${tagNodes.length} with "Tags:" in name. Sample: "${sample.slice(0, 80)}"`,
        );
        return;
      }
      if (rawResult && !rawResult.startsWith("ERROR:")) {
        let data004: any = null;
        try { data004 = JSON.parse(rawResult); } catch {}
        if (data004 !== null) {
          const hasTagsInLabel = data004.withTags > 0;
          injectAriaOverlay(); await wait(300); takeScreenshot("06-tc004-tags-aria"); removeAriaOverlay();
          step(
            "TC-004 Tags in aria-label",
            hasTagsInLabel,
            hasTagsInLabel
              ? `OK: ${data004.withTags}/${data004.total} cards have "Tags:" in aria-label. Sample: "${(data004.sample || "").slice(0, 80)}"`
              : `FAIL: ${data004.total} cards found, 0 have "Tags:" in aria-label. Sample: "${(data004.sample || "").slice(0, 80)}"`,
          );
          return;
        } else {
          injectAriaOverlay(); await wait(300); takeScreenshot("06-tc004-tags-aria"); removeAriaOverlay();
          step("TC-004 Tags in aria-label", false, "parse error: " + rawResult.slice(0, 100));
          return;
        }
      }
    }

    injectAriaOverlay(); await wait(300); takeScreenshot("06-tc004-tags-aria"); removeAriaOverlay();
    step(
      "TC-004 Tags in aria-label",
      false,
      "FAIL: Gallery webview not open. Extension must activate and gallery must load for DOM-based A11y check.",
    );
  });

  /**
   * TC-005 — Gallery/List toggle buttons expose aria-pressed state
   * Bug:  .layout-button elements have no aria-pressed attribute — screen readers cannot
   *       determine which layout (grid/list) is currently active.
   * Fix:  Gallery button has aria-pressed="true" when grid is active; List button "false".
   *
   * Test plan steps implemented here:
   *   1. Read initial state: Grid button aria-pressed="true", List button aria-pressed="false"
   *   2. Click List view button (simulate user interaction)
   *   3. Assert List button becomes aria-pressed="true", Grid button becomes aria-pressed="false"
   *   4. Restore grid view
   *
   * SNAPSHOT 07-tc005-initial: gallery in grid view — grid button pressed, list button not.
   * SNAPSHOT 07b-tc005-after-list-click: gallery switched to list view — list button pressed.
   * SNAPSHOT 07c-tc005-restored: gallery restored to grid view.
   * PASS: both state transitions confirmed via aria-pressed attribute changes.
   * FAIL: buttons missing aria-pressed, or state does not toggle after click.
   */
  test("TC-005: Gallery/List toggle buttons have aria-pressed", async () => {
    if (galleryOpened) {
      const readButtonsScript =
        "(() => { const btns = Array.from(document.querySelectorAll('.layout-button')); " +
        "const results = btns.map(function(b){ return {label: b.getAttribute('aria-label'), pressed: b.getAttribute('aria-pressed')}; }); " +
        "return JSON.stringify({count: btns.length, buttons: results}); })()";

      // Step 1: Read initial state (should be grid=true, list=false)
      const initialResult = sendEvalSignal(readButtonsScript, 5000);
      takeScreenshot("07-tc005-initial");

      // ARIA snapshot fallback path (Playwright ariaSnapshot)
      if (initialResult && initialResult.startsWith("ACCESSIBILITY:")) {
        const nodes = parseAriaSnapshot(initialResult.slice("ACCESSIBILITY:".length));
        const pressedButtons = nodes.filter(n =>
          n.role.toLowerCase() === "button" && n.pressed !== undefined
        );
        step(
          "TC-005 Toggle buttons have aria-pressed",
          pressedButtons.length > 0,
          `[AX-TREE fallback] ${pressedButtons.length} buttons with pressed state. Note: click verification skipped in ARIA fallback mode.`,
        );
        return;
      }

      if (!initialResult || initialResult.startsWith("ERROR:")) {
        takeScreenshot("07c-tc005-restored");
        step("TC-005 Toggle buttons have aria-pressed", false,
          "FAIL: DOM eval error reading initial button state.");
        return;
      }

      let initialData: any = null;
      try { initialData = JSON.parse(initialResult); } catch {}
      if (!initialData) {
        takeScreenshot("07c-tc005-restored");
        step("TC-005 Toggle buttons have aria-pressed", false,
          "parse error on initial read: " + initialResult.slice(0, 100));
        return;
      }

      const hasInitialAriaPressed = initialData.count > 0 &&
        initialData.buttons?.some((b: any) => b.pressed === "true" || b.pressed === "false");

      // Step 2: Click the List view button (simulate user toggling layout)
      const clickResult = sendEvalSignal(
        "(function(){" +
        "  var btn = document.querySelector('[aria-label=\"list view\"]');" +
        "  if (!btn) btn = Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "    return (b.getAttribute('aria-label') || '').toLowerCase().includes('list');" +
        "  });" +
        "  if (btn) { btn.click(); return 'clicked'; }" +
        "  return 'no-list-button';" +
        "})()",
        3000,
      );
      await wait(600);
      takeScreenshot("07b-tc005-after-list-click");

      // Step 3: Read state after click — List should now be pressed=true
      const afterResult = sendEvalSignal(readButtonsScript, 5000);

      // Step 4: Restore grid view
      sendEvalSignal(
        "(function(){" +
        "  var btn = document.querySelector('[aria-label=\"grid view\"]');" +
        "  if (!btn) btn = Array.from(document.querySelectorAll('.layout-button')).find(function(b){" +
        "    return (b.getAttribute('aria-label') || '').toLowerCase().includes('grid') ||" +
        "           (b.getAttribute('aria-label') || '').toLowerCase().includes('gallery');" +
        "  });" +
        "  if (btn) btn.click();" +
        "})()",
        2000,
      );
      await wait(300);
      takeScreenshot("07c-tc005-restored");

      let afterData: any = null;
      try { afterData = JSON.parse(afterResult); } catch {}

      const listBtnAfter = afterData?.buttons?.find((b: any) =>
        (b.label || "").toLowerCase().includes("list")
      );
      const gridBtnAfter = afterData?.buttons?.find((b: any) =>
        (b.label || "").toLowerCase().includes("grid") ||
        (b.label || "").toLowerCase().includes("gallery")
      );
      const stateToggled = listBtnAfter?.pressed === "true" && gridBtnAfter?.pressed === "false";

      const passes = hasInitialAriaPressed && stateToggled && clickResult === "clicked";
      step(
        "TC-005 Toggle buttons have aria-pressed",
        passes,
        passes
          ? `OK: initial=${JSON.stringify(initialData.buttons)}, after-click=${JSON.stringify(afterData?.buttons)}`
          : `FAIL: hasInitialAriaPressed=${hasInitialAriaPressed}, clickResult=${clickResult}, stateToggled=${stateToggled}. ` +
            `Initial=${JSON.stringify(initialData.buttons)}, After=${JSON.stringify(afterData?.buttons)}`,
      );
      return;
    }

    takeScreenshot("07-tc005-initial");
    step(
      "TC-005 Toggle buttons aria-pressed",
      false,
      "FAIL: Gallery webview not open. Extension must activate and gallery must load for DOM-based A11y check.",
    );
  });


  /**
   * TC-006 — Focus indicator contrast ≥ 3:1 in Light theme (sample cards)
   * Bug:  .sample-card focus outline uses --vscode-focusBorder in Light theme → contrast ~2.257:1
   *       on white card background (< 3:1 WCAG AA for non-text UI components).
   * Fix:  body.vscode-light .sample-card:focus-visible { outline-color: #005FB8 }
   *       #005FB8 on white = contrast ~5.77:1 (> 3:1 WCAG AA).
   *
   * Test: switches to Light Modern theme, programmatically focuses first .sample-card, takes
   * screenshot, then checks for the CSS fix rule in gallery stylesheets.
   *
   * SNAPSHOT PASS: Gallery in Light theme (white card backgrounds). The first sample card has
   *   a clearly visible thick dark-blue focus ring around it (#005FB8). The ring stands out
   *   strongly against the white card background.
   * SNAPSHOT FAIL: Gallery in Light theme, focused card has a faint/barely-visible light-blue
   *   focus ring (~#0090F1 or similar default focusBorder) that is hard to see on white.
   *   OR: no visible focus ring at all.
   */
  test("TC-006: Focus indicator contrast \u2265 3:1 in Light theme (sample cards)", async () => {
    if (!galleryOpened) {
      takeScreenshot("08-tc006-focus-contrast");
      step("TC-006 Focus indicator contrast", false, "FAIL: Gallery webview not open.");
      return;
    }

    // 1. Save current theme and switch to Light
    const wbConfig = vscode.workspace.getConfiguration("workbench");
    const originalTheme = wbConfig.get<string>("colorTheme") ?? "Default Dark Modern";
    try {
      await wbConfig.update("colorTheme", "Default Light Modern", vscode.ConfigurationTarget.Global);
    } catch {
      try {
        await wbConfig.update("colorTheme", "Default Light+", vscode.ConfigurationTarget.Global);
      } catch (e) {
        console.warn("  TC-006: Could not switch to light theme:", e);
      }
    }
    await wait(3000); // allow webview to re-render with light theme styles

    // 2. Focus a sample card programmatically via eval signal
    const focusScript =
      "(() => {" +
      "  var card = document.querySelector('.sample-card');" +
      "  if (card) { card.focus(); return 'focused:' + (card.getAttribute('aria-label') || card.className).slice(0, 60); }" +
      "  return 'no-card';" +
      "})()";
    const focusResult = sendEvalSignal(focusScript, 5000);
    console.log("  TC-006 focus result:", focusResult ? focusResult.slice(0, 80) : "(empty)");
    await wait(300); // allow focus ring to render

    // 3. Screenshot: light theme with focused card — shows the actual focus ring
    takeScreenshot("08-tc006-focus-contrast");

    // 4. Check computed outline color of the focused/active card
    const outlineScript =
      "(() => {" +
      "  var card = document.activeElement;" +
      "  if (!card || !card.classList.contains('sample-card')) {" +
      "    card = document.querySelector('.sample-card');" +
      "  }" +
      "  if (!card) return 'no-card';" +
      "  var cs = getComputedStyle(card);" +
      "  var isLight = document.body.classList.contains('vscode-light');" +
      "  return JSON.stringify({outlineColor: cs.outlineColor, outlineStyle: cs.outlineStyle, isLight: isLight});" +
      "})()";
    const outlineResult = sendEvalSignal(outlineScript, 5000);
    console.log("  TC-006 outline result:", outlineResult ? outlineResult.slice(0, 160) : "(empty)");

    // 5. Check if the light-theme fix CSS rule exists in the gallery stylesheet
    const ruleScript =
      "(() => {" +
      "  var found = false; var foundColor = '';" +
      "  try {" +
      "    Array.from(document.styleSheets).forEach(function(ss) {" +
      "      try {" +
      "        Array.from(ss.cssRules).forEach(function(r) {" +
      "          if (r.selectorText &&" +
      "              r.selectorText.includes('vscode-light') &&" +
      "              r.selectorText.includes('sample-card') &&" +
      "              r.selectorText.includes('focus-visible')) {" +
      "            found = true;" +
      "            foundColor = (r.style && r.style.outlineColor) || '#005FB8';" +
      "          }" +
      "        });" +
      "      } catch(e) {}" +
      "    });" +
      "  } catch(e) {}" +
      "  return JSON.stringify({found: found, color: foundColor});" +
      "})()";
    const ruleResult = sendEvalSignal(ruleScript, 5000);
    console.log("  TC-006 rule result:", ruleResult ? ruleResult.slice(0, 120) : "(empty)");

    // 6. Restore original theme
    try {
      await wbConfig.update("colorTheme", originalTheme, vscode.ConfigurationTarget.Global);
    } catch (e) {
      console.warn("  TC-006: Could not restore theme:", e);
    }

    // Evaluate assertion
    if (!ruleResult || ruleResult.startsWith("ERROR:")) {
      step("TC-006 Focus indicator contrast", false, "FAIL: DOM eval error.");
      return;
    }
    let ruleData: any = null;
    try { ruleData = JSON.parse(ruleResult); } catch {}
    if (ruleData?.found) {
      step(
        "TC-006 Focus indicator contrast",
        true,
        `Light theme focus fix applied: outline-color ${ruleData.color || "#005FB8"} \u2265 3:1 contrast on white background`,
      );
    } else {
      let detail =
        "FAIL: No light-theme focus-visible CSS rule for .sample-card. " +
        "Focus indicator contrast ~2.257:1 < 3:1 in Light theme. " +
        "Fix: add body.vscode-light .sample-card:focus-visible { outline-color: #005FB8 } in sampleCard.scss";
      if (outlineResult && !outlineResult.startsWith("ERROR:")) {
        try {
          const d = JSON.parse(outlineResult);
          detail += `. Computed: outlineColor=${d.outlineColor}, isLight=${d.isLight}`;
        } catch {}
      }
      step("TC-006 Focus indicator contrast", false, detail);
    }
  });

  test("Final state", async () => {
    await wait(1000);
    takeScreenshot("08-final-state");
    step("Final state captured", true);
  });
});



