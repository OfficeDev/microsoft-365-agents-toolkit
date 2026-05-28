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

function takeScreenshot(name: string): void {
  const dest   = path.join(SCREENSHOT_DIR, `${name}.png`);
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
  fs.writeFileSync(signal, `screenshot:${dest}`, "utf8");
  const deadline = Date.now() + 8000;
  while (fs.existsSync(signal) && Date.now() < deadline) {
    const end = Date.now() + 100; while (Date.now() < end) {}
  }
}

suite("Your Suite Name", function () {
  this.timeout(5 * 60 * 1000);
  // ... tests using vscode.* API + takeScreenshot()
});
```

### Key rules

- Fire UI-blocking commands WITHOUT `await` (wizard blocks until user action)
- Call `takeScreenshot()` after every meaningful UI state change — before AND after interactions
- Write results to `${TEST_OUTPUT_DIR}/results.json`:
  `{ passed: N, failed: N, steps: [{ name, status, detail }] }`

### Signal file types (write to `.signal` file; Playwright reads and acts)

| Signal content | Action |
|----------------|--------|
| `screenshot:<dest>` | Take screenshot |
| `clickText:<text>` | Click element containing text |
| `click:<selector>` | Click CSS selector |
| `type:<text>` | Type text |
| `pressKey:<key>` | Press keyboard key |
| `eval:<resultFile>:<script>` | Evaluate JS in webview, write result to resultFile |

