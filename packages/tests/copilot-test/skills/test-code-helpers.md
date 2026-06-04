# Test Code Helpers

Boilerplate for every `.test.ts` file in `packages/tests/copilot-test/src/`.

---

## Imports and directory constants

```typescript
import * as vscode from "vscode";
import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

const OUTPUT_DIR     = process.env.TEST_OUTPUT_DIR       || path.join(os.tmpdir(), "atk-test-output");
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR        || path.join(OUTPUT_DIR, "screenshots");
const SIGNAL_DIR     = process.env.SCREENSHOT_SIGNAL_DIR || path.join(OUTPUT_DIR, ".screenshot-signals");

function ensureDirs() {
  [OUTPUT_DIR, SCREENSHOT_DIR, SIGNAL_DIR].forEach(d => fs.mkdirSync(d, { recursive: true }));
}
```

---

## Screenshot helpers

Always use `takeElementScreenshot` — it clips to the element's bounding box so screenshots
show exactly the relevant UI surface, not the whole window.

```typescript
/** Full-page screenshot — fallback only. Prefer takeElementScreenshot(). */
async function takeScreenshot(name: string): Promise<void> {
  const dest   = path.join(SCREENSHOT_DIR, `${name}.png`);
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
  fs.writeFileSync(signal, `screenshot:${dest}`, "utf8");
  await _waitForSignal(signal, 8000);
}

/**
 * Element-level screenshot.
 * Playwright resolves `selector`, calls element.boundingBox() to get (x, y, w, h),
 * then clips the screenshot to that region.
 * Falls back to full-page if the element is not found or has no bounding box.
 *
 * Common selectors:
 *   .quick-input-widget                              — wizard / command-palette QuickPick
 *   .quick-input-list                                — QuickPick options list only
 *   [id='workbench.view.extension.teamsfx'] .pane-body  — ATK sidebar (all panes)
 *   .view-pane[aria-label*='Accounts']               — Accounts pane only
 *   .explorer-viewlet                                — File explorer
 *   .debug-toolbar                                   — Debug toolbar
 *   .notifications-toasts                            — Notification toasts
 *   #workbench\\.parts\\.statusbar                   — Status bar
 */
async function takeElementScreenshot(name: string, selector: string): Promise<void> {
  const dest   = path.join(SCREENSHOT_DIR, `${name}.png`);
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
  fs.writeFileSync(signal, `screenshotElement:${selector}:${dest}`, "utf8");
  await _waitForSignal(signal, 8000);
}

async function _waitForSignal(signal: string, timeoutMs: number): Promise<void> {
  await new Promise<void>(resolve => {
    const deadline = Date.now() + timeoutMs;
    const iv = setInterval(() => {
      if (!fs.existsSync(signal) || Date.now() >= deadline) {
        clearInterval(iv);
        try { if (fs.existsSync(signal)) fs.unlinkSync(signal); } catch {}
        resolve();
      }
    }, 100);
  });
}
```

---

## Signal and interaction helpers

```typescript
/**
 * Send an action signal to Playwright and wait for it to be processed.
 * Uses setInterval so the extension host event loop stays free.
 */
async function sendSignal(content: string, timeoutMs = 15000): Promise<void> {
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-action.signal`);
  fs.writeFileSync(signal, content, "utf8");
  await _waitForSignal(signal, timeoutMs);
}

function wait(ms: number): Promise<void> {
  return new Promise(r => setTimeout(r, ms));
}
```

---

## Step 0: ATK activation check (mandatory in every test)

```typescript
/**
 * Poll until the ATK extension is active. Must be the first test in every suite.
 * ATK activates on workspaceContains:m365agents*.yml — the workspace must have
 * that file before VS Code launches, or activation will not fire automatically.
 */
async function waitForATKActivation(timeoutMs = 60000): Promise<boolean> {
  const extId = "TeamsDevApp.ms-teams-vscode-extension";
  const deadline = Date.now() + timeoutMs;
  let ext = vscode.extensions.getExtension(extId);
  while (!ext?.isActive && Date.now() < deadline) {
    await wait(1000);
    ext = vscode.extensions.getExtension(extId);
  }
  if (ext && !ext.isActive) {
    try { await ext.activate(); } catch {}
  }
  return !!ext?.isActive;
}

// In suiteSetup or first test:
test("Step 0: ATK extension is active", async () => {
  const active = await waitForATKActivation();
  await takeElementScreenshot("00-atk-active",
    "[id='workbench.view.extension.teamsfx'] .pane-body");
  assert.ok(active, "ATK extension must be active before any test step runs");
});
```

---

## Results output

```typescript
function writeResults(passed: number, failed: number, steps: object[]) {
  fs.writeFileSync(
    path.join(OUTPUT_DIR, "results.json"),
    JSON.stringify({ passed, failed, steps }, null, 2),
    "utf8",
  );
}
```

---

## Signal reference

| Signal content | Playwright action |
|---|---|
| `screenshot:<dest>` | Full-page screenshot |
| `screenshotElement:<cssSelector>:<dest>` | `element.boundingBox()` → clip screenshot to element region |
| `clickText:<text>` | Click element containing text (tries getByText, monaco-list-row, then QuickPick filter) |
| `click:<selector>` | Click CSS selector |
| `type:<text>` | Type text into focused input |
| `pressKey:<key>` | Press keyboard key |
| `waitForTextThenScreenshot:<text>:<timeoutMs>:<name>` | Wait for text to appear, take full-page screenshot, delete signal |
| `eval:<resultFile>:<script>` | Evaluate JS in webview, write result JSON to resultFile |
