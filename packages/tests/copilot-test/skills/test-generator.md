---
name: atk-copilot-test-generator
description: >
  ATK generator agent: reads a test plan and generates .test.ts code for the
  @vscode/test-electron + Playwright CDP test framework. Used by atk-copilot-test-generator.yml.
---

# ATK Copilot Test Generator Skill

## Role

You are the **test code generator**. The label agent has already:
1. Fixed any product code needed.
2. Created/updated the test plan.
3. Committed and pushed to a branch.

Your job: read that test plan → write `.test.ts` code that faithfully implements it →
commit and push → write the script basename to `/tmp/script-name.txt`.

You do NOT run the test. The runner workflow handles execution.

---

## Workflow Overview

Every test, regardless of route, follows this top-level sequence:

```
Open VS Code  →  Step 0: Verify ATK activated  →  Route-specific steps  →  Write results.json
```

**Success condition per route:**

| Route | Primary success signal | Required screenshots |
|---|---|---|
| Scaffold | All expected files exist under project dir | wizard step (element), file-check (explorer element) |
| Local Debug | Debug session attaches (breakpoint or terminal active) | accounts view (element), debug toolbar (element) |
| Remote Provision | `teamsApp/create` output shows app ID in env file | accounts view (element), provisioned resource (element) |
| Remote Deploy | Deploy completes with no error notification | accounts view (element), notification (element) |

**Always take element-level screenshots** (not full-page) — see `takeElementScreenshot()` below.

---

## Routes

### Route: Scaffold (no auth required)

Scaffold creates a new project from the ATK wizard. The wizard tree is defined in
`templates/src/ui/` — consult the relevant `*.ts` file for the exact option labels for
each project type before writing the click sequence.

**Key files:**
- `templates/src/ui/root.ts` — first page (project-type) options
- `templates/src/ui/da.ts` — Declarative Agent sub-tree
- `templates/src/ui/cea.ts` — Custom Engine Agent sub-tree
- `templates/src/ui/teams.ts` — Teams Apps sub-tree

**Pattern:**

```typescript
// 1. Generate a unique app name to avoid collisions
const appName = `da-${Date.now()}`;

// 2. Fire wizard WITHOUT await (wizard blocks until user completes it)
vscode.commands.executeCommand("fx-extension.create").catch(() => {});
await wait(500);  // yield event loop so command reaches extension host

// 3. Each wizard step: waitForText → screenshotElement(.quick-input-widget) → clickText
await sendSignal("waitForTextThenScreenshot:Declarative Agent:60000:02-step1-project-type", 68000);
await takeElementScreenshot("02-step1-project-type", ".quick-input-widget");
await sendSignal("clickText:Declarative Agent", 10000);
await wait(800);

// ... repeat for each step (options come from templates/src/ui/*.ts)

// 4. App name step: type random name
await sendSignal("waitForTextThenScreenshot:Application Name:15000:07-step-appname", 23000);
await takeElementScreenshot("07-step-appname", ".quick-input-widget");
await sendSignal(`type:${appName}`, 8000);
await wait(300);
await sendSignal("pressKey:Enter", 5000);

// 5. Wait for scaffold (ATK opens a new VS Code window; close it)
await wait(30000);
await takeElementScreenshot("08-scaffold-progress", ".notifications-toasts");
await wait(60000);
// Close the new window that ATK opens for the scaffolded project
await vscode.commands.executeCommand("workbench.action.closeWindow");

// 6. Open the scaffolded project in the current window
const projectDir = findScaffoldedDir(appName);  // search ~/AgentsToolkitProjects + homedir
await vscode.commands.executeCommand("vscode.openFolder", vscode.Uri.file(projectDir));
await wait(5000);

// 7. Verify files — poll sentinel, then check each, screenshot explorer per check
const sentinel = path.join(projectDir, "m365agents.yml");
for (let i = 0; i < 60 && !fs.existsSync(sentinel); i++) await wait(1000);

for (const relPath of EXPECTED_FILES) {
  const exists = fs.existsSync(path.join(projectDir, relPath));
  step(`File: ${relPath}`, exists);
  await takeElementScreenshot(`verify-${relPath.replace(/\//g, "-")}`, ".explorer-viewlet");
}
```

**Always expected in scaffolded output** (regardless of template):
- `m365agents.yml`
- `appPackage/manifest.json`

Additional files depend on the specific template — document them in the test plan.

---

### Route: Local Debug (requires scaffold + M365 auth)

**Preconditions:**
1. A scaffolded project is open as the workspace (either run scaffold route first, or open an existing project).
2. M365 auth is completed — see **Auth Reference** below.

**Pattern after auth:**

```typescript
// Open debug config picker (does NOT block — no await)
vscode.commands.executeCommand("workbench.action.debug.selectandstart").catch(() => {});
await wait(500);

// Select the debug config for this template (label comes from .vscode/launch.json)
await sendSignal("waitForTextThenScreenshot:Debug in Teams:20000:10-debug-picker", 28000);
await takeElementScreenshot("10-debug-picker", ".quick-input-widget");
await sendSignal("clickText:Debug in Teams", 10000);

// Wait for debug session to start
await sendSignal("waitForTextThenScreenshot:Debug Console:120000:11-debug-started", 130000);
await takeElementScreenshot("11-debug-toolbar", ".debug-toolbar");
```

---

### Route: Remote Provision / Deploy (requires M365 + Azure auth)

Follow the same auth pattern for both M365 and Azure (see **Auth Reference**), then:

```typescript
// Provision
vscode.commands.executeCommand("fx-extension.provision").catch(() => {});
// Deploy
vscode.commands.executeCommand("fx-extension.deploy").catch(() => {});
```

Screenshot the notification toast on completion:
```typescript
await takeElementScreenshot("provision-done", ".notifications-toasts");
```

---

## Auth Reference

**Source patches** (must be applied + rebuilt before running auth-dependent tests):

| File | Patch | Effect |
|---|---|---|
| `packages/vscode-extension/src/commonlib/m365Login.ts` | `if (process.env["TEAMSFX_AUTO_CONFIRM_LOGIN"] === "true") return Promise.resolve(true);` at top of `doesUserConfirmLogin()` | Bypasses the VS Code modal confirmation dialog |
| `packages/vscode-extension/src/commonlib/codeFlowLogin.ts` | Write auth URL to `$TEMP/atk-auth-url.txt` in `getAuthCodeUrl().then()` | Makes auth URL accessible to test automation |

**Required env vars** when launching VS Code:
```
TEAMSFX_AUTO_CONFIRM_LOGIN=true
TEAMSFX_BROKER_AUTH=false   # forces browser code flow (MSAL), not WAM
```

**Auth flow in test code:**

```typescript
// 1. Trigger sign-in via ATK accounts tree item click
//    (ATK must be activated first — workspaceContains:m365agents*.yml fires activation)
await sendSignal("waitForTextThenScreenshot:Sign in to Microsoft 365:20000:auth-01-signin-item", 28000);
await takeElementScreenshot("auth-01-accounts-pane",
  "[id='workbench.view.extension.teamsfx'] .pane-body");
await sendSignal("clickText:Sign in to Microsoft 365", 10000);

// 2. Poll for auth URL written by codeFlowLogin.ts patch
const AUTH_URL_FILE = path.join(os.tmpdir(), "atk-auth-url.txt");
let authUrl = "";
for (let i = 0; i < 30 && !authUrl; i++) {
  await wait(1000);
  if (fs.existsSync(AUTH_URL_FILE)) {
    authUrl = fs.readFileSync(AUTH_URL_FILE, "utf-8").trim();
  }
}
step("Auth URL captured", !!authUrl, authUrl ? authUrl.substring(0, 60) + "..." : "not found");

// 3. Open browser for OAuth (test infra handles the actual login)
//    The runner environment must have credentials pre-seeded.

// 4. After OAuth, verify sign-in by checking account view tree
await sendSignal("waitForTextThenScreenshot:M365 account:60000:auth-02-signed-in", 70000);
await takeElementScreenshot("auth-02-account-view",
  "[id='workbench.view.extension.teamsfx'] .pane-body");
```

---

## Test Quality Rules — Simulate Real User Behavior

Tests must reflect how a real user interacts with the extension. The following rules are **mandatory**:

1. **Every Step in the test plan must be implemented.** Each `Steps:` entry in the matching `test-plans/*.md` file must have a corresponding action in the test script. Do NOT skip steps or replace behavioral actions with static assertions.

2. **UI interactions must use real actions, not static checks.**
   - If the test plan says "Click List view button" → the test must call `btn.click()` (via `sendEvalSignal`) or `sendSignal("clickText:...")`.
   - If the test plan says "Toggle Gallery/List" → the test must click the button AND verify the resulting state change.
   - Reading an attribute without triggering any interaction is only acceptable for *precondition* verification, not the main assertion.

3. **CSS-existence checks are not behavioral tests.**
   - Checking whether a CSS rule exists in a stylesheet does NOT verify the user experience.
   - Use CSS-rule checks only as a *supplementary* signal alongside behavioral verification (e.g., focus the element, take screenshot, then also confirm the rule exists).
   - Never use a CSS-rule check as the sole pass/fail criterion unless the test plan explicitly calls for it.

4. **Color contrast must be computed, never hardcoded or blacklisted.**
   - When a test plan requires contrast ratio verification (e.g., "contrast >= 4.5:1"), use the **WCAG relative luminance formula** directly in the injected JS. Do NOT use a hardcoded blacklist of "bad" colors.
   - Inject this helper into every contrast test via `sendEvalSignal`:
     ```js
     function relativeLuminance(r, g, b) {
       var srgb = [r, g, b].map(function(c) {
         c = c / 255;
         return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
       });
       return 0.2126 * srgb[0] + 0.7152 * srgb[1] + 0.0722 * srgb[2];
     }
     function contrastRatio(L1, L2) {
       var lighter = Math.max(L1, L2);
       var darker  = Math.min(L1, L2);
       return (lighter + 0.05) / (darker + 0.05);
     }
     function effectiveBg(el, theme) {
       var node = el;
       while (node) {
         var cs = getComputedStyle(node);
         var bg = cs.backgroundColor;
         if (bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)') {
           var m = bg.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
           if (m) return [parseInt(m[1]), parseInt(m[2]), parseInt(m[3])];
         }
         node = node.parentElement;
       }
       // Read VS Code CSS variable as fallback before hardcoding
       var bodyBg = getComputedStyle(document.body).backgroundColor;
       var bm = bodyBg && bodyBg.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
       if (bm) return [parseInt(bm[1]), parseInt(bm[2]), parseInt(bm[3])];
       // Last resort: use theme-appropriate fallback
       return theme === 'dark' ? [30, 30, 30] : [255, 255, 255];
     }
     ```
   - Pass `theme` ('light' or 'dark') to `effectiveBg()` so the fallback is correct. **Never hardcode `[255,255,255]` as the only fallback** — in dark theme the background is near-black, not white.
   - Parse `getComputedStyle(el).color` (format: `rgb(r, g, b)`) for foreground color.
   - The pass criterion is `contrastRatio(...) >= threshold` where threshold is the value stated in the test plan (4.5 for normal text, 3.0 for large text / non-text UI).
   - Return the computed ratio in the step `detail` field so it is visible in `results.json`.

4. **State change must be verified after interaction.**
   - After clicking a toggle, re-query the DOM to confirm the `aria-pressed`, class, or visible state changed.
   - Include before-click and after-click screenshots.

5. **Test plan is the spec.** If a TC does not have an entry in `test-plans/`, create one before writing the test. No TC may exist in the `.test.ts` file without a corresponding plan entry.

---

## Step-by-Step: Generate Test Code

### Step 1 — Find the test plan

```bash
git log --oneline -5
# Find plan added/modified by the label agent:
git show --name-only HEAD -- packages/tests/copilot-test/test-plans/ | grep '\.md$' || \
git diff HEAD~1 --name-only -- packages/tests/copilot-test/test-plans/ | grep '\.md$'
# Read it:
cat <test-plan-path>
```

### Step 2 — Read infrastructure docs

```bash
cat packages/tests/copilot-test/README.md
ls packages/tests/copilot-test/src/
```

### Step 3 — Write the test file

Naming: `packages/tests/copilot-test/src/<feature>-<task>.test.ts`

Follow the Test Quality Rules above (every test plan step = code, real UI interactions, state verification).

Reference implementation: `packages/tests/copilot-test/src/simple-bot-create.test.ts`

### Step 4 — Save the script basename

```bash
# Write just the basename WITHOUT .test.ts extension:
echo "sample-app-a11y" > /tmp/script-name.txt
```

### Step 5 — Commit and push

```bash
git add packages/tests/copilot-test/src/
git commit -m "test(generated): <feature> test script (issue #$ISSUE)"
git push origin $(git branch --show-current)
```

---

## Test Code Structure

```typescript
import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

const OUTPUT_DIR     = process.env.TEST_OUTPUT_DIR          || path.join(os.tmpdir(), "atk-test-output");
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR           || path.join(OUTPUT_DIR, "screenshots");
const SIGNAL_DIR     = process.env.SCREENSHOT_SIGNAL_DIR    || path.join(OUTPUT_DIR, ".screenshot-signals");

function ensureDirs() {
  [OUTPUT_DIR, SCREENSHOT_DIR, SIGNAL_DIR].forEach(d => fs.mkdirSync(d, { recursive: true }));
}

/** Full-page screenshot (fallback only — prefer takeElementScreenshot) */
async function takeScreenshot(name: string): Promise<void> {
  const dest   = path.join(SCREENSHOT_DIR, `${name}.png`);
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
  fs.writeFileSync(signal, `screenshot:${dest}`, "utf8");
  await _waitForSignal(signal, 8000);
}

/**
 * Element-level screenshot: Playwright gets the element's bounding box (x,y,w,h)
 * and clips the screenshot to that region. Always prefer this over takeScreenshot().
 *
 * @param name    Screenshot basename (no extension)
 * @param selector  CSS selector for the element to capture
 *
 * Common selectors:
 *   Wizard QuickPick panel:   ".quick-input-widget"
 *   QuickPick options list:   ".quick-input-list"
 *   ATK sidebar (all panes):  "[id='workbench.view.extension.teamsfx'] .pane-body"
 *   Accounts pane:            ".view-pane[aria-label*='Accounts']"
 *   File explorer:            ".explorer-viewlet"
 *   Debug toolbar:            ".debug-toolbar"
 *   Notification toasts:      ".notifications-toasts"
 *   Status bar:               "#workbench\\.parts\\.statusbar"
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

async function sendSignal(content: string, timeoutMs = 15000): Promise<void> {
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-action.signal`);
  fs.writeFileSync(signal, content, "utf8");
  await _waitForSignal(signal, timeoutMs);
}

function wait(ms: number): Promise<void> {
  return new Promise(r => setTimeout(r, ms));
}

/** Step 0 (mandatory in every test): verify ATK extension is active */
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

suite("Your Suite Name", function () {
  this.timeout(10 * 60 * 1000);

  suiteSetup(() => ensureDirs());

  test("Step 0: ATK extension is active", async () => {
    const active = await waitForATKActivation();
    // Screenshot the whole workbench to confirm extension loaded
    await takeElementScreenshot("00-atk-active",
      "[id='workbench.view.extension.teamsfx'] .pane-body");
    assert.ok(active, "ATK extension must be active before any test step runs");
  });

  // ... route-specific tests
});
```

### Key rules

- Fire UI-blocking commands WITHOUT `await` (wizard blocks until user action)
- **Always use `takeElementScreenshot()`** — full-page `takeScreenshot()` is a fallback only
- Call `takeElementScreenshot()` after every meaningful UI state change
- Write results to `${TEST_OUTPUT_DIR}/results.json`:
  `{ passed: N, failed: N, steps: [{ name, status, detail }] }`

### Signal file types (write to `.signal` file; Playwright reads and acts)

| Signal content | Action |
|----------------|--------|
| `screenshot:<dest>` | Full-page screenshot |
| `screenshotElement:<cssSelector>:<dest>` | **Element screenshot** — Playwright calls `element.boundingBox()` and clips to that region |
| `clickText:<text>` | Click element containing text |
| `click:<selector>` | Click CSS selector |
| `type:<text>` | Type text |
| `pressKey:<key>` | Press keyboard key |
| `waitForTextThenScreenshot:<text>:<timeoutMs>:<screenshotName>` | Wait for text to appear, then full-page screenshot |
| `eval:<resultFile>:<script>` | Evaluate JS in webview, write result to resultFile |

