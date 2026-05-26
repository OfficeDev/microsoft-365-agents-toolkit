# Test Plan: Sample App Panel Accessibility (A11y)

## Metadata

- **feature-slug**: `sample-app-a11y`
- **owner**: atk-qa
- **created**: 2026-05-15
- **updated**: 2026-05-26
- **triggers**: issue-label `atk-copilot-test`, manual
- **related-issue**: [#15916](https://github.com/OfficeDev/microsoft-365-agents-toolkit/issues/15916)

## Core Principle

Tests simulate **real user behavior**: every step describes what a user clicks, types, tabs to, or
observes — not internal DOM attributes or CSS values. Screenshots capture what the user actually sees.

All contrast checks use the **WCAG relative luminance formula** (IEC 61966-2-1 sRGB).

---

## Scope

**Covers:**
- TC-001: Link text color contrast ≥ 4.5:1 when the link is **focused** in Light theme
- TC-002: Gallery/List toggle buttons expose correct `aria-pressed` state before and after click
- TC-003: Sample card accessible names include `. Tags: <tag1>, <tag2>` on keyboard focus
- TC-004: Featured state visual indicator (badge) has non-text contrast ≥ 3:1 against card background (WCAG 1.4.11)
- TC-005: Featured cards have aria-label starting with `"Featured sample."`, non-featured cards do not
- TC-006a: Focus ring contrast ≥ 3:1 in **Gallery view** (Light theme)
- TC-006b: Focus ring contrast ≥ 3:1 in **List view** (Light theme)

**Does NOT cover:**
- Full keyboard Tab-order traversal across the entire panel
- Live screen reader announcement testing (NVDA/JAWS)
- Mobile or touch accessibility
- High Contrast theme
- Dark theme variants

---

## Test Cases

### TC-001 – Link text color contrast ≥ 4.5:1 when focused (Light theme)

**Preconditions:**
- VS Code is open, ATK extension activated
- VS Code theme is set to a Light variant (e.g. Default Light Modern)

**Steps:**
1. User opens Command Palette (`Ctrl+Shift+P`) and runs `Microsoft 365 Agents Toolkit: View Samples`
2. User observes the Sample Gallery panel opens
3. Test focuses a `.ms-Link` element via `element.focus()`
4. Test reads `getComputedStyle(link).color` from the focused link
5. Test computes `contrast_ratio(link_color, background_color)` using WCAG relative luminance
6. Take screenshot showing the link in its **focused state** with a visible focus indicator

**Expected result:**
- Link color in focused state is visibly dark on the white/light background
- Computed contrast ratio ≥ 4.5

**Pass criteria:**
- `contrast_ratio(getComputedStyle(link).color, background) >= 4.5`
- The `detail` field in the result includes the actual computed ratio value

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                       | Pass condition                                        | Why                                                         |
|-----|-----------------------------------|-------------------------------------------------------|-------------------------------------------------------|-------------------------------------------------------------|
| 01  | `01-extension-active.png`         | VS Code with ATK extension activated                  | ATK sidebar icon visible                              | Baseline: extension loaded before gallery opens             |
| 02  | `02-gallery-open.png`             | Sample Gallery panel with sample cards                | Gallery renders with link text on light background    | Proves gallery opens after user command                     |
| 03  | `03-tc001-link-focused.png`       | Focused link element in gallery (focus ring visible)  | Link text is dark and clearly readable                | Proves light-theme focused link color meets WCAG 4.5:1      |

---

### TC-002 – Gallery/List toggle buttons expose `aria-pressed` before and after click

**Preconditions:**
- Sample Gallery is open (default Gallery/Grid layout active)

**Steps:**
1. User opens Sample Gallery
2. Test reads `aria-pressed` attribute on the **Gallery** button — expects `"true"`
3. Test reads `aria-pressed` attribute on the **List** button — expects `"false"`
4. Take screenshot showing Gallery layout is active
5. User clicks the **List** view toggle button
6. Test reads `aria-pressed` on **Gallery** button — expects `"false"`
7. Test reads `aria-pressed` on **List** button — expects `"true"`
8. Take screenshot showing List layout is active

**Expected result:**
- Before click: Gallery `aria-pressed="true"`, List `aria-pressed="false"`
- After clicking List: List `aria-pressed="true"`, Gallery `aria-pressed="false"`

**Pass criteria:**
- Before click: `galleryButton.ariaPressed === "true"` AND `listButton.ariaPressed === "false"`
- After click: `listButton.ariaPressed === "true"` AND `galleryButton.ariaPressed === "false"`
- `detail` field must include button labels and their pressed values

**Screenshots produced by test:**

| ID  | Filename                        | What is visible                                         | Pass condition                                 | Why                                                              |
|-----|---------------------------------|---------------------------------------------------------|------------------------------------------------|------------------------------------------------------------------|
| 04  | `04-tc002-gallery-active.png`   | Gallery layout, Gallery toggle button visually selected | Gallery button highlighted                     | Proves initial aria-pressed state matches visual selected state  |
| 05  | `05-tc002-list-active.png`      | List layout, List toggle button visually selected       | List button highlighted, Gallery button not    | Proves aria-pressed toggles correctly in response to click       |

---

### TC-003 – Sample card accessible names include tags on keyboard focus

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. User opens Sample Gallery
2. User presses **Tab** to move keyboard focus onto the first sample card
3. Test reads `aria-label` attribute from the focused card element
4. Test verifies the `aria-label` contains `". Tags:"` followed by at least one tag name
5. Take screenshot showing the **focused card** with focus ring visible

**Expected result:**
- Accessible name: `"<Title>. Tags: <tag1>, <tag2>, <tag3>"`
- Screen reader would announce title then tags when card receives focus

**Pass criteria:**
- Focused card `aria-label` matches pattern `/.+\. Tags: .+/`
- At least one tag name appears after `". Tags:"`

**Screenshots produced by test:**

| ID  | Filename                      | What is visible                                          | Pass condition                                        | Why                                                            |
|-----|-------------------------------|----------------------------------------------------------|-------------------------------------------------------|----------------------------------------------------------------|
| 06  | `06-tc003-focused-card.png`   | Gallery with first card focused, focus ring visible      | Card has visible focus ring; aria-label includes tags | Proves tags are in accessible name for screen reader users     |

---

### TC-004 – Featured state visual indicator contrast ≥ 3:1 against card background (WCAG 1.4.11)

**Preconditions:**
- Sample Gallery is open in Light theme
- At least one featured card is visible with a `.featured-badge` element

**Steps:**
1. User opens Sample Gallery in Light theme
2. User observes the featured badge (dark gold label) on featured cards
3. Test reads `getComputedStyle(badge).backgroundColor` from a `.featured-badge` element
4. Test reads the card background color from the parent `.sample-card`
5. Test computes `contrast_ratio(badge_background, card_background)` using WCAG relative luminance
6. Take screenshot showing **both featured and non-featured cards** together

**Expected result:**
- Featured badge background color (dark gold `#7A5C00`) contrasts strongly with card background (white)
- Computed contrast ratio ≥ 3.0 (WCAG 1.4.11 non-text contrast)

**Pass criteria:**
- `contrast_ratio(badge_bg_color, card_bg_color) >= 3.0`
- `detail` field includes computed ratio plus both color values

**Screenshots produced by test:**

| ID  | Filename                        | What is visible                                          | Pass condition                                            | Why                                                                    |
|-----|---------------------------------|----------------------------------------------------------|-----------------------------------------------------------|------------------------------------------------------------------------|
| 07  | `07-tc004-featured-badge.png`   | Featured and non-featured cards side by side in gallery  | Featured badge (gold) visually distinct from card bg      | Proves non-text contrast of featured indicator meets WCAG AA (≥3:1)   |

---

### TC-005 – Screen reader differentiates Featured from non-Featured cards

**Preconditions:**
- Sample Gallery is open in either Grid or List view
- At least one Featured and one non-Featured card are visible

**Steps:**
1. User opens Sample Gallery
2. Test finds all card elements and reads their `aria-label` attributes
3. Test verifies at least one card has `aria-label` starting with `"Featured sample."`
4. Test verifies at least one card does NOT have `aria-label` starting with `"Featured sample."`
5. Take screenshot

**Expected result:**
- Featured card: `aria-label="Featured sample. <Title>. Tags: <tags>"`
- Non-featured card: `aria-label="<Title>. Tags: <tags>"` (no "Featured sample." prefix)

**Pass criteria:**
- At least 1 card element with `aria-label` starting with `"Featured sample."`
- At least 1 card element with `aria-label` NOT starting with `"Featured sample."`

**Screenshots produced by test:**

| ID  | Filename                      | What is visible                                           | Pass condition                               | Why                                                          |
|-----|-------------------------------|-----------------------------------------------------------|----------------------------------------------|--------------------------------------------------------------|
| 08  | `08-tc005-aria-labels.png`    | Gallery showing featured and non-featured cards           | Featured badge visible on featured cards     | Proves ARIA differentiation matches visual differentiation   |

---

### TC-006a – Focus ring contrast ≥ 3:1 in Gallery view (Light theme)

**Preconditions:**
- Sample Gallery is open in **Gallery (Grid) view**
- VS Code theme is set to a Light variant

**Steps:**
1. User opens Sample Gallery in Gallery view
2. User presses **Tab** to move keyboard focus to the first `.sample-card`
3. Test reads `getComputedStyle(card).outlineColor` from the focused card
4. Test reads the card background color
5. Test computes `contrast_ratio(outline_color, card_background)` using WCAG relative luminance
6. Take screenshot **after Tab** so the focus ring is clearly visible

**Expected result:**
- Focused card has a dark-blue focus ring (`#005FB8`)
- Computed contrast ratio ≥ 3.0 (WCAG AA non-text threshold)

**Pass criteria:**
- `contrast_ratio(outline_color, card_bg) >= 3.0`
- `detail` field includes the computed ratio

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                        | Pass condition                                       | Why                                                                    |
|-----|-----------------------------------|--------------------------------------------------------|------------------------------------------------------|------------------------------------------------------------------------|
| 09  | `09-tc006a-gallery-focus.png`     | Gallery view with first card focused, focus ring shown | Dark-blue outline clearly surrounds the focused card | Proves keyboard users can identify focused card in gallery Light theme |

---

### TC-006b – Focus ring contrast ≥ 3:1 in List view (Light theme)

**Preconditions:**
- Sample Gallery is open and user has switched to **List view**
- VS Code theme is set to a Light variant

**Steps:**
1. User opens Sample Gallery (default Gallery view)
2. User clicks the **List** view toggle button to switch to List layout
3. User presses **Tab** to move keyboard focus to the first `.sample-list-item`
4. Test reads `getComputedStyle(listItem).outlineColor` from the focused list item
5. Test reads the list item background color
6. Test computes `contrast_ratio(outline_color, list_item_background)` using WCAG relative luminance
7. Take screenshot **after Tab** so the focus ring is clearly visible

**Expected result:**
- Focused list item has a dark-blue focus ring (`#005FB8`)
- Computed contrast ratio ≥ 3.0

**Pass criteria:**
- `contrast_ratio(outline_color, list_item_bg) >= 3.0`
- `detail` field includes the computed ratio

**Screenshots produced by test:**

| ID  | Filename                        | What is visible                                          | Pass condition                                         | Why                                                                 |
|-----|---------------------------------|----------------------------------------------------------|--------------------------------------------------------|---------------------------------------------------------------------|
| 10  | `10-tc006b-list-focus.png`      | List view with first list item focused, focus ring shown | Dark-blue outline clearly surrounds the focused item   | Proves keyboard users can identify focused item in list Light theme |

---

## Test Script

`packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Notes

- The Sample Gallery renders inside a VS Code webview. DOM evaluation is performed via the Playwright CDP session targeting the webview frame.
- All contrast calculations use the WCAG relative luminance formula (IEC 61966-2-1 sRGB, not a simple color blacklist).
- For TC-001: the link must be **focused** (`element.focus()`) before reading computed color — the `:focus` CSS rule changes the link color from `#005B9E` to `#004480`.
- For TC-004: WCAG 1.4.11 (Non-text Contrast) applies to UI components; the featured badge background is the visual indicator that differentiates featured items. The badge uses `background-color: #7A5C00` which contrasts ~5:1 against the white card background.
- For TC-006: `element.focus()` triggers `:focus-visible` styles in VSCode webview. The light-theme override sets `outline-color: #005FB8` (~10:1 against white) on both `.sample-card` and `.sample-list-item`.
- `aria-pressed` on `<vscode-button>` is forwarded to the inner `<button>` element by the FAST foundation runtime.
- TC-002, TC-003, TC-005: these attributes are set inline in the React render path and are accessible without CDP frame evaluation.
