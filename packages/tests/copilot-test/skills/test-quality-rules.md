# Test Quality Rules

Tests must reflect how a real user interacts with the extension. All rules are **mandatory**.

---

## 1 — Every step in the test plan must be implemented

Each `Steps:` entry in the matching `test-plans/*.md` file must have a corresponding action
in the test script. Do NOT skip steps or replace behavioral actions with static assertions.

---

## 2 — UI interactions must use real actions, not static checks

- "Click List view button" → the test must call `btn.click()` (via `sendEvalSignal`) or `sendSignal("clickText:...")`.
- "Toggle Gallery/List" → the test must click the button AND verify the resulting state change.
- Reading an attribute without triggering any interaction is only acceptable for *precondition* verification, not the main assertion.

---

## 3 — CSS-existence checks are not behavioral tests

Checking whether a CSS rule exists in a stylesheet does NOT verify the user experience.
Use CSS-rule checks only as a *supplementary* signal alongside behavioral verification
(e.g., focus the element, take screenshot, then also confirm the rule exists).
Never use a CSS-rule check as the sole pass/fail criterion unless the test plan explicitly calls for it.

---

## 4 — Color contrast must be computed, never hardcoded or blacklisted

When a test plan requires contrast ratio verification (e.g., "contrast >= 4.5:1"), use the
**WCAG relative luminance formula** directly in the injected JS. Do NOT use a hardcoded
blacklist of "bad" colors.

Inject this helper into every contrast test via `sendEvalSignal`:

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
  // Last resort: theme-appropriate fallback
  return theme === 'dark' ? [30, 30, 30] : [255, 255, 255];
}
```

- Pass `theme` (`'light'` or `'dark'`) to `effectiveBg()` so the fallback is correct.
  **Never hardcode `[255,255,255]` as the only fallback** — dark theme background is near-black.
- Parse `getComputedStyle(el).color` (format: `rgb(r, g, b)`) for foreground color.
- Pass criterion: `contrastRatio(...) >= threshold` (4.5 for normal text, 3.0 for large text / non-text UI).
- Return the computed ratio in the step `detail` field so it appears in `results.json`.

---

## 5 — Verify state change after interaction

After clicking a toggle, re-query the DOM to confirm the `aria-pressed`, class, or visible
state changed. Include **before-click** and **after-click** screenshots.

---

## 6 — Test plan is the spec

If a TC does not have an entry in `test-plans/`, create one before writing the test.
No TC may exist in the `.test.ts` file without a corresponding plan entry.
