# Bug Report — Issue #15916

## Test Case
**TC-006b: Focus ring contrast >= 3:1 in List view (Light theme)**
Test file: `packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Expected Behaviour
Per WCAG 2.1 SC 1.4.11 (Non-text Contrast), the focus ring rendered on a focused
`.sample-list-item` must achieve a contrast ratio of **at least 3:1** against its adjacent
background colour.

In List view the first focusable item sits inside the "Featured" section whose background is
`rgb(60,139,197)`. The focus ring must therefore have a contrast ratio >= 3:1 against that
blue background.

Gallery view already passes: the same focus-ring colour (`rgb(0,95,184)`) measured against a
white card background (`rgb(255,255,255)`) achieves 6.31:1.

## Actual Behaviour
When keyboard focus moves to the first `.sample-list-item` in List view the computed outline
colour is `rgb(0,95,184)` (solid). The effective background resolved by walking up the DOM is
`rgb(60,139,197)` — the featured section's blue fill.

| Element | Colour | WCAG Luminance |
|---------|--------|----------------|
| Focus ring (outline) | `rgb(0, 95, 184)` | ~0.127 |
| Featured section background | `rgb(60, 139, 197)` | ~0.248 |

**Computed contrast ratio: 1.71:1** — well below the required 3:1 minimum.

## CI Evidence
```
Eval result: {"outlineRaw":"rgb(0, 95, 184)","bgRgb":"rgb(60,139,197)",
              "outlineStyle":"solid","ratio":1.71,"passes":false}
FAIL TC-006b List focus ring contrast >= 3:1:
     Computed ratio=1.71:1; outline=rgb(0, 95, 184) (solid); bg=rgb(60,139,197)
```

## Root Cause
The ATK extension's sample-gallery webview uses the same focus-ring colour
(`rgb(0,95,184)`, VS Code's default blue) in both Gallery and List views, but does not
account for the featured section's opaque blue background present only in List view. No CSS
override provides a higher-contrast outline for list items that sit on the coloured section
background.

## Impact
WCAG 2.1 SC 1.4.11 failure. Keyboard-only and low-vision users cannot reliably see which
list item has focus when it is located in the "Featured" section, because the focus ring is
nearly invisible against the blue background.

## Suggested Fix
Add a CSS rule in the gallery webview styles that increases the focus-ring contrast for
`.sample-list-item:focus-visible` elements inside the featured section, for example:

```css
/* Ensure focus ring is visible against the featured-section blue background */
.featured-section .sample-list-item:focus-visible {
  outline: 2px solid #ffffff;   /* white — contrast vs rgb(60,139,197) ≈ 4.2:1 */
  outline-offset: 2px;
}
```

Alternatively, use a double-ring (white inner + dark outer via `box-shadow`) that achieves
sufficient contrast against any background without requiring a section-specific override.
