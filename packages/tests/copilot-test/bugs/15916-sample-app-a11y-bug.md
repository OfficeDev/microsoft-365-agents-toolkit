# Bug Report — Issue #15916

## Test Case
**TC-004: Featured badge non-text contrast >= 3:1 (WCAG 1.4.11)**
Test file: `packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Expected Behaviour
Per WCAG 1.4.11 (Non-text Contrast), the "Featured" badge background colour must achieve
a contrast ratio of **at least 3:1** against the adjacent card background colour in all
supported VS Code colour themes (including Dark theme).

## Actual Behaviour
In VS Code Dark theme the badge and card rendered with the following colours:

| Element | Colour | WCAG Luminance |
|---------|--------|----------------|
| `.featured-badge` background | `rgb(122, 92, 0)` | ~0.134 |
| `.sample-card` background | `rgb(37, 37, 38)` | ~0.025 |

**Computed contrast ratio: 2.45:1** — below the required 3:1 minimum.

The badge colour `rgb(122, 92, 0)` (dark amber/gold) was likely chosen to contrast against
a light card background but is insufficient when rendered on the VS Code dark-theme card
surface `rgb(37, 37, 38)`.

## Steps to Reproduce
1. Open VS Code with the default Dark theme.
2. Open the ATK extension's Sample Gallery (command `fx-extension.openSamples`).
3. Inspect the background colour of any `.featured-badge` element via DevTools /
   `getComputedStyle`.
4. Inspect the background colour of its parent `.sample-card` element.
5. Compute WCAG contrast ratio — result is 2.45:1 (fails >= 3:1).

## Impact
WCAG 2.1 SC 1.4.11 failure. Users who rely on colour to identify "Featured" samples (e.g.
low-vision users) cannot reliably distinguish the badge from the card background in dark
theme.

## Suggested Fix
Update the `.featured-badge` token/CSS variable so the background colour has a contrast
ratio >= 3:1 against the dark-theme card surface `rgb(37, 37, 38)`.  For example, a
saturated amber such as `rgb(255, 185, 0)` (~`#FFB900`) yields a ratio of ~7.5:1 against
that background, which satisfies both dark and light themes.

The fix should be applied in the Sample Gallery webview source (likely in the
`packages/vscode-extension` or `packages/fx-core` package under the sample-gallery UI
assets).
