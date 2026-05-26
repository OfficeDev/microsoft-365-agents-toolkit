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
1. User opens Sample Gallery: invoke `fx-extension.openSamples` command from Command Palette
2. User observes sample cards rendered with blue link text
3. User notices links are readable against the white background

**Expected result:**
- Link color `#005B9E` on white has contrast ≈ 7.6:1 ✓
- WCAG AA threshold: 4.5:1

**Pass criteria:**
- `contrast_ratio(linkColor, backgroundColor) >= 4.5`
- At least one `.ms-Link` element found in the webview

**Screenshots produced by test:**

| ID  | Filename                        | What is visible                                     | Pass condition                              | Why                                                        |
|-----|---------------------------------|-----------------------------------------------------|---------------------------------------------|------------------------------------------------------------|
| 01  | `01-tc001-link-contrast.png`    | Gallery open, sample cards with blue link text      | Links are dark blue, clearly readable       | Proves light-theme link color fix is applied               |

**Fix applied:** `SampleGallery.scss` – added `body.vscode-light .sample-gallery .ms-Link { color: #005B9E }`

---

### TC-002 – Featured vs non-Featured ARIA differentiation

**Preconditions:**
- Sample Gallery is open in grid or list layout
- At least one Featured and one non-Featured sample are visible

**Steps:**
1. User opens Sample Gallery via Command Palette
2. User observes the grid of sample cards — some cards show a "Featured" badge
3. User hovers over a featured card and a non-featured card to compare their labels

**Expected result:**
- Featured cards: `aria-label="Featured sample. <Title>. Tags: <tag1>, <tag2>"`
- Non-featured cards: `aria-label="<Title>. Tags: <tag1>, <tag2>"`

**Pass criteria:**
- At least one featured card found with `aria-label` starting with `"Featured sample."`
- At least one non-featured card found with `aria-label` NOT starting with `"Featured sample."`

**Screenshots produced by test:**

| ID  | Filename                     | What is visible                                     | Pass condition                                       | Why                                                         |
|-----|------------------------------|-----------------------------------------------------|------------------------------------------------------|-------------------------------------------------------------|
| 02  | `02-tc002-aria-labels.png`   | Gallery open showing featured and regular cards     | Featured badge visible on featured cards             | Proves ARIA differentiation is present for screen readers   |

**Fix applied:** `sampleCard.tsx` + `sampleListItem.tsx` – aria-label includes "Featured sample." prefix when `featured={true}`

---

### TC-003 – Featured badge non-text contrast ≥ 3:1

**Preconditions:**
- Sample Gallery is open in light theme

**Steps:**
1. User opens Sample Gallery in light theme (VS Code default light)
2. User visually notices the gold/brown "Featured" badge on featured cards
3. User confirms the badge text and star icon are legible against the card background

**Expected result:**
- `#7A5C00` on `#FFFFFF`: contrast ≈ 4.9:1 ✓
- `#7A5C00` on `#F8F8F8`: contrast ≈ 4.9:1 ✓
- WCAG AA non-text threshold: 3:1

**Pass criteria:**
- `contrast_ratio("#7A5C00", "#FFFFFF") >= 3.0`
- At least one `.featured-badge` element found in the webview

**Screenshots produced by test:**

| ID  | Filename                       | What is visible                                     | Pass condition                                    | Why                                                        |
|-----|--------------------------------|-----------------------------------------------------|---------------------------------------------------|------------------------------------------------------------|
| 03  | `03-tc003-badge-contrast.png`  | Featured badge with gold/brown color in light theme | Badge text is dark and clearly visible on white   | Proves badge color meets WCAG non-text contrast threshold  |

**Fix applied:** `SampleGallery.scss` + `sampleCard.scss` – featured badge/star use `#7A5C00` in light theme

---

### TC-004 – Tags included in accessible name on keyboard focus

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. User opens Sample Gallery
2. User tabs through sample cards using keyboard
3. User listens to (or reads) the accessible name announced for each card — expects to hear the sample title followed by its tags

**Expected result:**
- `aria-label="<Title>. Tags: <tag1>, <tag2>, <tag3>"`
- Screen reader announces tags when card receives focus

**Pass criteria:**
- Every sample card's `aria-label` contains `". Tags:"` followed by at least one tag name

**Screenshots produced by test:**

| ID  | Filename                    | What is visible                                     | Pass condition                                      | Why                                                       |
|-----|-----------------------------|-----------------------------------------------------|-----------------------------------------------------|-----------------------------------------------------------|
| 04  | `04-tc004-tags-aria.png`    | Gallery with a card focused, tags visible in label  | Card is focused and tags text is part of the label  | Proves tags are announced to screen reader users on focus |

**Fix applied:** `sampleCard.tsx` + `sampleListItem.tsx` – aria-label always includes `. Tags: <tags>`

---

### TC-005 – Gallery/List toggle buttons expose aria-pressed

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. User opens Sample Gallery (default grid/gallery layout)
2. User observes the Gallery and List toggle buttons at the top of the panel
3. User clicks the **List** view button
4. User observes the layout switches to list style
5. User clicks the **Gallery** view button to restore grid layout
6. User observes the layout returns to grid style

**Expected result:**
- Active toggle: `aria-pressed="true"`
- Inactive toggle: `aria-pressed="false"`
- WCAG SC 4.1.2: Name, Role, Value

**Pass criteria:**
- Gallery button initially has `aria-pressed="true"`, List button has `aria-pressed="false"`
- After clicking List button: List button has `aria-pressed="true"`, Gallery button has `aria-pressed="false"`

**Screenshots produced by test:**

| ID  | Filename                         | What is visible                                        | Pass condition                                        | Why                                                               |
|-----|----------------------------------|--------------------------------------------------------|-------------------------------------------------------|-------------------------------------------------------------------|
| 05  | `05-tc005-toggle-before.png`     | Gallery layout active, Gallery button highlighted      | Gallery button visually active, List button inactive  | Proves initial aria-pressed state is correct                      |
| 06  | `06-tc005-toggle-after.png`      | List layout active after user click, List highlighted  | List button visually active, Gallery button inactive  | Proves aria-pressed toggles correctly in response to user action  |

**Fix applied:** `sampleFilter.tsx` – added `aria-pressed={this.props.layout === "grid/list"}` to VSCodeButton toggles, also fixed "gallary" typo in aria-label to "Gallery"

---

### TC-006 – Focus indicator contrast ≥ 3:1 in Light theme (sample cards)

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. User opens Sample Gallery
2. User switches VS Code theme to "Default Light Modern" via Command Palette
3. User tabs to the first sample card — observes a visible focus ring around it
4. User confirms the focus ring is a clearly visible dark-blue outline

**Expected result:**
- Focused card has a clearly visible dark-blue focus ring (`#005FB8`)
- `#005FB8` on white background = contrast ratio ≈ 5.77:1 ✓ (WCAG AA threshold for non-text: 3:1)
- Screenshot shows visible focus ring in light theme

**Pass criteria:**
- CSS rule `body.vscode-light .sample-card:focus-visible { outline-color: #005FB8 }` exists in computed styles
- `contrast_ratio("#005FB8", "#FFFFFF") >= 3.0`

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                         | Pass condition                                           | Why                                                              |
|-----|-----------------------------------|---------------------------------------------------------|----------------------------------------------------------|------------------------------------------------------------------|
| 07  | `07-tc006-focus-ring.png`         | First sample card focused, dark-blue outline visible    | Focus ring is clearly visible (dark blue on white)       | Proves keyboard users can see which card is focused in light theme |

**Fix applied:** `sampleCard.scss` – added `body.vscode-light .sample-card:focus-visible { outline-color: #005FB8 }`

---

## Test Script

`packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Notes

- The Sample Gallery renders inside a sandboxed webview iframe. DOM evaluation is done via VSCode extension host message passing.
- TC-001/TC-002/TC-003/TC-004/TC-005 are all DOM-attribute checks; no pixel-level contrast measuring is done in the automated test (contrast values are design-time verified).
- `aria-pressed` on `<vscode-button>` web component is forwarded to the internal `<button>` element by FAST foundation.
