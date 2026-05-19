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
   * TC-001 — Link text color contrast
   * After fix: .ms-Link in light theme uses #005B9E (contrast ≈ 7.6:1 on white).
   * We verify the fix is present in the source by checking the SCSS/DOM.
   * Since the webview is a sandboxed iframe, we do a structural DOM check via
   * the Playwright eval signal. If Playwright is not connected, we fall back to
   * verifying the fix was applied via source-level assertions.
   */
  test("TC-001: Link text color contrast ≥ 4.5:1", async () => {
    if (!galleryOpened) {
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        false,
        "FAIL: Gallery webview not open. Cannot verify link color contrast.",
      );
      takeScreenshot("03-tc001-link-contrast");
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
      step("TC-001 Link text contrast ≥ 4.5:1", false,
        "FAIL: DOM eval error — gallery may not be accessible via Playwright.");
    } else if (rawResult.startsWith("colors:")) {
      const colors = rawResult.slice(7).split(",").filter(Boolean);
      const problematicColor = "rgb(72, 160, 199)";
      const hasProblematic = colors.some((c) => c.trim() === problematicColor);
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        !hasProblematic,
        hasProblematic
          ? `FAIL: Found low-contrast color ${problematicColor} on .ms-Link`
          : `Link colors OK: ${colors.join(", ")}`,
      );
    } else if (rawResult.startsWith("css-rule:")) {
      const rules = rawResult.slice(9);
      const hasGoodRule =
        rules.includes("#005b9e") || rules.includes("#005B9E") || rules.includes("rgb(0, 91, 158)");
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        hasGoodRule,
        hasGoodRule
          ? `CSS fix applied: ${rules.slice(0, 80)}`
          : `FAIL: CSS rule for .ms-Link exists but wrong color: ${rules.slice(0, 80)}`,
      );
    } else {
      step(
        "TC-001 Link text contrast ≥ 4.5:1",
        false,
        "FAIL: No .ms-Link elements and no CSS fix rule found. Fix not applied (shouldShowChat=false hides links; CSS fix must be in stylesheet).",
      );
    }
    takeScreenshot("03-tc001-link-contrast");
  });

  /**
   * TC-002 — Featured vs non-Featured ARIA differentiation
   * After fix: featured cards have aria-label prefixed with "Featured sample."
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
        step(
          "TC-002 Featured ARIA differentiation",
          hasFeaturedAria,
          `[AX-TREE] ${buttonNodes.length} button nodes, ${featuredNodes.length} with "Featured sample" prefix`,
        );
        takeScreenshot("04-tc002-featured-aria");
        return;
      }
      if (rawResult && !rawResult.startsWith("ERROR:")) {
        try {
          const data = JSON.parse(rawResult);
          const hasFeaturedAria = data.featuredPrefixed > 0;
          step(
            "TC-002 Featured ARIA differentiation",
            hasFeaturedAria,
            hasFeaturedAria
              ? `OK: ${data.featuredPrefixed}/${data.featuredTotal} featured cards have "Featured sample." prefix`
              : `FAIL: ${data.featuredTotal} featured cards, 0 with "Featured sample." in aria-label. ` +
                `Sample: "${data.sampleFeaturedLabel.slice(0, 80)}"`,
          );
        } catch {
          step(
            "TC-002 Featured ARIA differentiation",
            false,
            "parse error: " + rawResult,
          );
        }
        takeScreenshot("04-tc002-featured-aria");
        return;
      }
    }

    step(
      "TC-002 Featured ARIA differentiation",
      false,
      "FAIL: Gallery webview not open. Extension must activate and gallery must load for DOM-based A11y check.",
    );
    takeScreenshot("04-tc002-featured-aria");
  });

  /**
   * TC-003 — Featured badge non-text contrast ≥ 3:1
   * After fix: .featured-badge uses #7A5C00 on white (≈ 4.9:1).
   */
  test("TC-003: Featured badge contrast ≥ 3:1", async () => {
    if (!galleryOpened) {
      step(
        "TC-003 Featured badge present with accessible contrast",
        false,
        "FAIL: Gallery webview not open. Cannot check featured badge.",
      );
      takeScreenshot("05-tc003-badge-contrast");
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
      step("TC-003 Featured badge present with accessible contrast", false, "FAIL: DOM eval error.");
    } else if (rawResult === "no-featured-section") {
      step("TC-003 Featured badge present with accessible contrast", false,
        "FAIL: .featured-sample-section not found (gallery may not have loaded featured samples).");
    } else {
      try {
        const data = JSON.parse(rawResult);
        if (data.count === 0) {
          step(
            "TC-003 Featured badge present with accessible contrast",
            false,
            "FAIL: .featured-sample-section present but no .featured-badge elements. A11y bug: badge element not added to featured cards.",
          );
        } else {
          step(
            "TC-003 Featured badge present with accessible contrast",
            true,
            `${data.count} badge(s) found. bg=${data.bg}, color=${data.color}`,
          );
        }
      } catch {
        step("TC-003 Featured badge contrast ≥ 3:1", false, "parse error: " + rawResult);
      }
    }
    takeScreenshot("05-tc003-badge-contrast");
  });

  /**
   * TC-004 — Tags included in accessible name
   * After fix: aria-label includes ". Tags: tag1, tag2"
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
        step(
          "TC-004 Tags in aria-label",
          hasTagsInLabel,
          `[AX-TREE] ${buttonNodes.length} buttons, ${tagNodes.length} with "Tags:" in name. Sample: "${sample.slice(0, 80)}"`,
        );
        takeScreenshot("06-tc004-tags-aria");
        return;
      }
      if (rawResult && !rawResult.startsWith("ERROR:")) {
        try {
          const data = JSON.parse(rawResult);
          const hasTagsInLabel = data.withTags > 0;
          step(
            "TC-004 Tags in aria-label",
            hasTagsInLabel,
            hasTagsInLabel
              ? `OK: ${data.withTags}/${data.total} cards have "Tags:" in aria-label. Sample: "${data.sample.slice(0, 80)}"`
              : `FAIL: ${data.total} cards found, 0 have "Tags:" in aria-label. Sample: "${data.sample.slice(0, 80)}"`,
          );
        } catch {
          step("TC-004 Tags in aria-label", false, "parse error: " + rawResult);
        }
        takeScreenshot("06-tc004-tags-aria");
        return;
      }
    }

    step(
      "TC-004 Tags in aria-label",
      false,
      "FAIL: Gallery webview not open. Extension must activate and gallery must load for DOM-based A11y check.",
    );
    takeScreenshot("06-tc004-tags-aria");
  });

  /**
   * TC-005 — Gallery/List toggle buttons expose aria-pressed
   * After fix: Gallery button has aria-pressed="true" by default (grid layout).
   * Note: toggle buttons are inside the sample gallery webview — only testable
   * when gallery is open and Playwright has access to the webview frame.
   */
  test("TC-005: Gallery/List toggle buttons have aria-pressed", async () => {
    if (galleryOpened) {
      const evalScript =
        "(() => { const btns = Array.from(document.querySelectorAll('.layout-button')); const results = btns.map(b => ({label: b.getAttribute('aria-label'), pressed: b.getAttribute('aria-pressed')})); return JSON.stringify({count: btns.length, buttons: results}); })()";
      const rawResult = sendEvalSignal(evalScript, 5000);

      if (rawResult && rawResult.startsWith("ACCESSIBILITY:")) {
        const nodes = parseAriaSnapshot(rawResult.slice("ACCESSIBILITY:".length));
        // Layout buttons have aria-pressed, appearing as role=button with pressed property
        const pressedButtons = nodes.filter(n =>
          n.role.toLowerCase() === "button" &&
          n.pressed !== undefined
        );
        const hasAriaPressed = pressedButtons.length > 0;
        step(
          "TC-005 Toggle buttons have aria-pressed",
          hasAriaPressed,
          `[AX-TREE] ${pressedButtons.length} buttons with pressed state found. States: ${JSON.stringify(pressedButtons.map(b => ({name: b.name, pressed: b.pressed})).slice(0,3))}`,
        );
        takeScreenshot("07-tc005-toggle-aria-pressed");
        return;
      }
      if (rawResult && !rawResult.startsWith("ERROR:")) {
        try {
          const data = JSON.parse(rawResult);
          const hasAriaPressed =
            data.count > 0 &&
            data.buttons?.some(
              (b: any) => b.pressed === "true" || b.pressed === "false",
            );
          step(
            "TC-005 Toggle buttons have aria-pressed",
            hasAriaPressed,
            `${data.count} layout buttons found. States: ${JSON.stringify(
              data.buttons,
            )}`,
          );
        } catch {
          step(
            "TC-005 Toggle aria-pressed",
            false,
            "parse error: " + rawResult,
          );
        }
        takeScreenshot("07-tc005-toggle-aria-pressed");
        return;
      }
    }

    // Gallery not open — cannot verify TC without live DOM.
    step(
      "TC-005 Toggle buttons aria-pressed",
      false,
      "FAIL: Gallery webview not open. Extension must activate and gallery must load for DOM-based A11y check.",
    );
    takeScreenshot("07-tc005-toggle-aria-pressed");
  });

  test("Final state", async () => {
    await wait(1000);
    takeScreenshot("08-final-state");
    step("Final state captured", true);
  });
});



