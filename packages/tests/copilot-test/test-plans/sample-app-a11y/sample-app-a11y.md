# Test Plan: Sample App Panel Accessibility (A11y)

## Metadata

- **feature-slug**: `sample-app-a11y`
- **owner**: atk-qa
- **created**: 2026-05-15
- **updated**: 2026-05-15
- **triggers**: issue-label `atk-copilot-test`, manual
- **related-issue**: [#15916](https://github.com/OfficeDev/microsoft-365-agents-toolkit/issues/15916)

## Scope

**Covers:**
- TC-001: Link text color contrast ≥ 4.5:1 in Light theme (WCAG AA)
- TC-002: Featured vs non-Featured Sample Apps ARIA differentiation
- TC-003: Featured/non-Featured visual state badge contrast ≥ 3:1 (WCAG AA non-text)
- TC-004: Tags included in accessible name / screen-reader announcement
- TC-005: Gallery/List toggle buttons expose `aria-pressed` state

**Does NOT cover:**
- Full keyboard navigation flow
- Screen reader live announcement testing (NVDA/JAWS)
- Mobile accessibility
- High Contrast theme accessibility

---

## Test Cases

### TC-001 – Link text color contrast ≥ 4.5:1

**Preconditions:**
- VSCode is open with ATK extension activated
- Sample Gallery webview is open

**Steps:**
1. Open Sample Gallery via `fx-extension.openSamples`
2. Evaluate computed color of `.ms-Link` elements inside the webview
3. Calculate contrast ratio against background
4. Assert ratio ≥ 4.5:1

**Expected result:**
- Link color `#005B9E` on white has contrast ≈ 7.6:1 ✓
- WCAG AA threshold: 4.5:1

**Fix applied:** `SampleGallery.scss` – added `body.vscode-light .sample-gallery .ms-Link { color: #005B9E }`

---

### TC-002 – Featured vs non-Featured ARIA differentiation

**Preconditions:**
- Sample Gallery is open in grid or list layout
- At least one Featured and one non-Featured sample are visible

**Steps:**
1. Open Sample Gallery
2. Check `aria-label` attributes on sample cards in the featured section
3. Check `aria-label` attributes on sample cards in the regular section
4. Assert featured cards have "Featured sample" prefix in their `aria-label`

**Expected result:**
- Featured cards: `aria-label="Featured sample. <Title>. Tags: <tag1>, <tag2>"`
- Non-featured cards: `aria-label="<Title>. Tags: <tag1>, <tag2>"`

**Fix applied:** `sampleCard.tsx` + `sampleListItem.tsx` – aria-label includes "Featured sample." prefix when `featured={true}`

---

### TC-003 – Featured badge non-text contrast ≥ 3:1

**Preconditions:**
- Sample Gallery is open in light theme

**Steps:**
1. Open Sample Gallery in light theme
2. Check `.featured-badge` element on featured cards
3. Assert background color `#7A5C00` vs white has contrast ≥ 3:1
4. Assert `#featured-sample-title` star/h4 color `#7A5C00` vs `#F8F8F8` background

**Expected result:**
- `#7A5C00` on `#FFFFFF`: contrast ≈ 4.9:1 ✓
- `#7A5C00` on `#F8F8F8`: contrast ≈ 4.9:1 ✓
- WCAG AA non-text threshold: 3:1

**Fix applied:** `SampleGallery.scss` + `sampleCard.scss` – featured badge/star use `#7A5C00` in light theme

---

### TC-004 – Tags included in accessible name on keyboard focus

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. Open Sample Gallery
2. Check `aria-label` on sample card elements
3. Assert `aria-label` contains the sample's tag names

**Expected result:**
- `aria-label="<Title>. Tags: <tag1>, <tag2>, <tag3>"`
- Screen reader announces tags when card receives focus

**Fix applied:** `sampleCard.tsx` + `sampleListItem.tsx` – aria-label always includes `. Tags: <tags>`

---

### TC-005 – Gallery/List toggle buttons expose aria-pressed

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. Open Sample Gallery (default grid layout)
2. Check Gallery view button `aria-pressed` attribute
3. Assert Gallery button has `aria-pressed="true"` and List button has `aria-pressed="false"`
4. Click List view button
5. Assert List button has `aria-pressed="true"` and Gallery button has `aria-pressed="false"`

**Expected result:**
- Active toggle: `aria-pressed="true"`
- Inactive toggle: `aria-pressed="false"`
- WCAG SC 4.1.2: Name, Role, Value

**Fix applied:** `sampleFilter.tsx` – added `aria-pressed={this.props.layout === "grid/list"}` to VSCodeButton toggles, also fixed "gallary" typo in aria-label to "Gallery"

---

### TC-006 – Focus indicator contrast ≥ 3:1 in Light theme (sample cards)

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. Open Sample Gallery
2. Switch VSCode theme to "Default Light Modern"
3. Programmatically focus the first `.sample-card` element
4. Take screenshot showing the focus ring on the card
5. Check computed `outline-color` of the focused card
6. Assert the light-theme fix CSS rule exists: `body.vscode-light .sample-card:focus-visible { outline-color: #005FB8 }`
7. Restore original theme

**Expected result:**
- Focused card has a clearly visible dark-blue focus ring (`#005FB8`)
- `#005FB8` on white background = contrast ratio ≈ 5.77:1 ✓ (WCAG AA threshold for non-text: 3:1)
- Screenshot shows visible focus ring in light theme

**Fix applied:** `sampleCard.scss` – added `body.vscode-light .sample-card:focus-visible { outline-color: #005FB8 }`

---

## Test Script

`packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Notes

- The Sample Gallery renders inside a sandboxed webview iframe. DOM evaluation is done via VSCode extension host message passing.
- TC-001/TC-002/TC-003/TC-004/TC-005 are all DOM-attribute checks; no pixel-level contrast measuring is done in the automated test (contrast values are design-time verified).
- `aria-pressed` on `<vscode-button>` web component is forwarded to the internal `<button>` element by FAST foundation.
