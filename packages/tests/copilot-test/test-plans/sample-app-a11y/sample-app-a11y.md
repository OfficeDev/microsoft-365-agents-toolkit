# Test Plan: Sample App Panel Accessibility (A11y)

## Metadata

- **feature-slug**: `sample-app-a11y`
- **owner**: atk-qa
- **created**: 2026-05-15
- **updated**: 2026-06-01
- **triggers**: issue-label `atk-copilot-test`, manual
- **related-issue**: [#15916](https://github.com/OfficeDev/microsoft-365-agents-toolkit/issues/15916)

## Core Principle

Tests simulate **real user behavior**: every step describes what a user clicks, types, tabs to, or
observes — not internal DOM attributes or CSS values. Screenshots capture what the user actually sees.

All contrast checks use the **WCAG relative luminance formula** (IEC 61966-2-1 sRGB).

---

## Scope

**Covers:**
- TC-001a: Link text color contrast ≥ 4.5:1 when the link is **focused** in **Light** theme (Default Light Modern)
- TC-001b: Link text color contrast ≥ 4.5:1 when the link is **focused** in **Dark** theme (Default Dark Modern)
- TC-002: Gallery/List toggle buttons expose correct `aria-pressed` state before and after click
- TC-003: Sample card accessible names include `. Tags: <tag1>, <tag2>` on keyboard focus
- TC-004a: Featured cards in **Gallery (grid) view** display a blue corner triangle badge with white star (`codicon-star-full`) on the top-left of the thumbnail image; non-featured cards do not
- TC-004b: Featured items in **List view** display a star (`codicon-star-full`) icon before the sample title text in the **same color as the title text** (`currentColor`); non-featured items do not
- TC-005: Featured cards have aria-label starting with `"Featured sample."`, non-featured cards do not
- TC-006a: Focus ring contrast ≥ 3:1 in **Gallery view** (Light theme)
- TC-006b: Focus ring contrast ≥ 3:1 in **List view** (Light theme)

**Does NOT cover:**
- Full keyboard Tab-order traversal across the entire panel
- Live screen reader announcement testing (NVDA/JAWS)
- Mobile or touch accessibility
- High Contrast theme

---

## Test Cases

### TC-001a – Link text color contrast ≥ 4.5:1 when focused (Light theme)

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon (M365 Agents Toolkit) in the VS Code Activity Bar to open the ATK extension panel. Observe the extension sidebar appears.
3. Take screenshot showing the ATK extension panel open in Light theme.
4. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery panel opens showing sample cards with link text.
5. Take screenshot showing the Sample Gallery open with visible link text on the light background.
6. Focus a `.ms-Link` element inside the gallery using `element.focus()` programmatically. Observe the focus ring appears around the link.
7. Read `getComputedStyle(link).color` from the focused link and compute `contrast_ratio(link_color, background_color)` using the WCAG relative luminance formula.
8. Take screenshot showing the link in its **focused state** with the focus indicator visible.

**Expected result:**
- ATK extension panel opens in Light theme showing the M365 Agents Toolkit UI.
- Sample Gallery panel renders with sample cards and visible link text.
- Link color in focused state is visibly dark (≥ `#005B9E`) on the white/light background.
- Computed contrast ratio ≥ 4.5:1.

**Pass criteria:**
- `contrast_ratio(getComputedStyle(link).color, background) >= 4.5`
- The `detail` field in the result includes the actual computed ratio value.

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                             | Pass condition                                          | Why                                                              |
|-----|-----------------------------------|-------------------------------------------------------------|---------------------------------------------------------|------------------------------------------------------------------|
| 01  | `01-tc001a-extension-active.png`  | VS Code with ATK extension panel open in Light theme        | ATK sidebar panel visible in Default Light Modern theme | Baseline: extension loaded and light theme active before test   |
| 02  | `02-tc001a-gallery-open.png`      | Sample Gallery panel with sample cards and link text        | Gallery renders with dark link text on light background | Proves gallery opens correctly after user command                |
| 03  | `03-tc001a-link-focused.png`      | Focused `.ms-Link` element with focus ring visible          | Link text is dark and clearly readable with focus ring  | Proves fixed link color (#005B9E, ≥4.5:1) is applied in light theme |

---

### TC-001b – Link text color contrast ≥ 4.5:1 when focused (Dark theme)

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Dark Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Dark Modern").
2. Click the ATK icon (M365 Agents Toolkit) in the VS Code Activity Bar to open the ATK extension panel. Observe the extension sidebar appears with dark background.
3. Take screenshot showing the ATK extension panel open in Dark theme.
4. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery panel opens showing sample cards with link text on a dark background.
5. Take screenshot showing the Sample Gallery open with visible link text on the dark background.
6. Focus a `.ms-Link` element inside the gallery using `element.focus()` programmatically. Observe the focus ring appears around the link.
7. Read `getComputedStyle(link).color` from the focused link and the background color, then compute `contrast_ratio(link_color, background_color)` using the WCAG relative luminance formula.
8. Take screenshot showing the link in its **focused state** with the focus indicator visible.

**Expected result:**
- ATK extension panel opens in Dark theme.
- Sample Gallery panel renders with sample cards and visible light-colored link text on the dark background.
- Computed contrast ratio ≥ 4.5:1.

**Pass criteria:**
- `contrast_ratio(getComputedStyle(link).color, background) >= 4.5`
- The `detail` field includes the actual computed ratio value.

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                              | Pass condition                                           | Why                                                               |
|-----|-----------------------------------|--------------------------------------------------------------|----------------------------------------------------------|-------------------------------------------------------------------|
| 04  | `04-tc001b-extension-active.png`  | VS Code with ATK extension panel open in Dark theme          | ATK sidebar panel visible in Default Dark Modern theme   | Baseline: extension loaded and dark theme active before test      |
| 05  | `05-tc001b-gallery-open.png`      | Sample Gallery panel with sample cards and link text (dark)  | Gallery renders with light link text on dark background  | Proves gallery opens correctly in dark theme                      |
| 06  | `06-tc001b-link-focused.png`      | Focused `.ms-Link` element with focus ring visible (dark)    | Link text is light and clearly readable with focus ring  | Proves link color meets WCAG AA ≥4.5:1 in dark theme             |

---

### TC-002 – Gallery/List toggle buttons expose `aria-pressed` before and after click

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel. Observe the ATK sidebar opens.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery opens in Gallery (Grid) layout by default.
4. Read `aria-pressed` attribute on the **Gallery** view toggle button — expects `"true"`.
5. Read `aria-pressed` attribute on the **List** view toggle button — expects `"false"`.
6. Take screenshot showing Gallery layout is active with Gallery toggle visually selected.
7. Click the **List** view toggle button. Observe the layout switches to a vertical list.
8. Read `aria-pressed` on **Gallery** button — expects `"false"`.
9. Read `aria-pressed` on **List** button — expects `"true"`.
10. Take screenshot showing List layout is active with List toggle visually selected.

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
| 04  | `04-tc002-gallery-active.png`   | Gallery layout, Gallery toggle button visually selected | Gallery button highlighted, list layout absent | Proves initial aria-pressed state matches visual selected state  |
| 05  | `05-tc002-list-active.png`      | List layout, List toggle button visually selected       | List button highlighted, Gallery button not    | Proves aria-pressed toggles correctly in response to user click  |

---

### TC-003 – Sample card accessible names include tags on keyboard focus

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery panel opens.
4. Click inside the Sample Gallery webview area to give it focus.
5. Press **Tab** to move keyboard focus onto the first sample card. Observe the focus ring appears around a card.
6. Read the `aria-label` attribute from the focused card element.
7. Verify the `aria-label` contains `". Tags:"` followed by at least one tag name.
8. Take screenshot showing the **focused card** with the focus ring clearly visible.

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

### TC-004a – Featured cards display corner triangle badge on thumbnail in Gallery view

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery opens in Gallery (Grid) view. Observe the "Featured samples" section with sample cards.
4. Take screenshot showing the featured samples section — a blue corner triangle badge with a white star should be visible on the top-left corner of featured card thumbnails.
5. Find all `.sample-card` elements inside `.featured-sample-section`. For each, verify the card contains a `.featured-corner-badge` element and inside it a `.featured-corner-star` element (the `codicon-star-full` span).
6. Find all `.sample-card` elements **outside** `.featured-sample-section` (i.e., in `.sample-section`). Verify none contain a `.featured-corner-badge` element.
7. Verify the `.featured-sample-section` element has no background-color set (i.e., `getComputedStyle(section).backgroundColor` equals `"rgba(0, 0, 0, 0)"` or inherits from the page — confirming the background was removed per the new design).

**Expected result:**
- Each featured card in gallery view has a `.featured-corner-badge` span containing a `.featured-corner-star` span (white star on blue corner triangle) overlaid on the top-left of the thumbnail.
- Non-featured cards do not have this element.
- The featured section has no distinct background color (it matches the page background).

**Pass criteria:**
- `featuredCards.every(card => card.querySelector('.featured-corner-badge') !== null)`
- `featuredCards.every(card => card.querySelector('.featured-corner-badge .featured-corner-star') !== null)`
- `nonFeaturedCards.every(card => card.querySelector('.featured-corner-badge') === null)`
- `getComputedStyle(featuredSection).backgroundColor` is transparent or matches page bg

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                                        | Pass condition                                                        | Why                                                                       |
|-----|-----------------------------------|------------------------------------------------------------------------|-----------------------------------------------------------------------|---------------------------------------------------------------------------|
| 13  | `13-tc004a-gallery-star.png`      | Gallery view with featured section, blue corner badge visible on thumbnails   | Blue triangle badge with white star visible on top-left corner of featured card thumbnails      | Proves star-on-badge indicator distinguishes featured items visually in gallery view |

---

### TC-004b – Featured items display star icon before title in List view (color matches text)

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery opens.
4. Click the **List** view toggle button to switch to List layout. Observe the layout changes to a vertical list.
5. Take screenshot showing the list view with the featured section at the top — the star icons should appear before the title text of featured items.
6. Find all `.sample-list-item` elements inside `.featured-sample-section`. For each, verify the item's `h3` contains a `.featured-star` span (the `codicon-star-full` icon).
7. Find all `.sample-list-item` elements **outside** `.featured-sample-section` (i.e., in `.sample-section`). Verify none have a `.featured-star` span in their `h3`.
8. Read `getComputedStyle(starSpan).color` for a `.featured-star` element. Read `getComputedStyle(h3).color` for its parent `h3`. Verify the two colors are equal (star inherits text color via `currentColor`).

**Expected result:**
- Each featured list item has a `codicon-star-full` icon as the first child of its `h3` element, before the title text.
- The star icon uses `currentColor`, so it renders in the same color as the surrounding title text.
- Non-featured list items do not have this icon in their `h3`.

**Pass criteria:**
- `featuredListItems.every(item => item.querySelector('h3 .featured-star') !== null)`
- `nonFeaturedListItems.every(item => item.querySelector('h3 .featured-star') === null)`
- `getComputedStyle(starSpan).color === getComputedStyle(h3).color` (star color matches title color)

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                                        | Pass condition                                                        | Why                                                                       |
|-----|-----------------------------------|------------------------------------------------------------------------|-----------------------------------------------------------------------|---------------------------------------------------------------------------|
| 14  | `14-tc004b-list-star.png`         | List view with featured section, star icons visible before title text  | Star icon visible before title of featured list items; star color matches title color  | Proves star indicator uses text color (not yellow) in list view  |

---

### TC-005 – Screen reader differentiates Featured from non-Featured cards

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery opens showing a mix of featured and non-featured cards.
4. Take screenshot showing the gallery with both featured and non-featured cards visible.
5. Find all card elements (`.sample-card`) and read their `aria-label` attributes.
6. Verify at least one card has `aria-label` starting with `"Featured sample."`.
7. Verify at least one card does NOT have `aria-label` starting with `"Featured sample."`.

**Expected result:**
- Featured card: `aria-label="Featured sample. <Title>. Tags: <tags>"`
- Non-featured card: `aria-label="<Title>. Tags: <tags>"` (no "Featured sample." prefix)

**Pass criteria:**
- At least 1 card element with `aria-label` starting with `"Featured sample."`
- At least 1 card element with `aria-label` NOT starting with `"Featured sample."`

**Screenshots produced by test:**

| ID  | Filename                      | What is visible                                              | Pass condition                                        | Why                                                              |
|-----|-------------------------------|--------------------------------------------------------------|-------------------------------------------------------|------------------------------------------------------------------|
| 08  | `08-tc005-aria-labels.png`    | Gallery showing both featured and non-featured cards         | Featured and non-featured cards both visible          | Proves ARIA "Featured sample." prefix is present on featured items |

---

### TC-006a – Focus ring contrast ≥ 3:1 in Gallery view (Light theme)

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery opens in Gallery (Grid) view.
4. Click inside the Sample Gallery webview area to give it keyboard focus.
5. Press **Tab** to move keyboard focus to the first `.sample-card`. Observe the dark-blue focus ring appears around the card.
6. Take screenshot **immediately after Tab** so the focus ring is clearly visible.
7. Read `getComputedStyle(card).outlineColor` from the focused card.
8. Read the card background color from the focused card element.
9. Compute `contrast_ratio(outline_color, card_background)` using the WCAG relative luminance formula.

**Expected result:**
- Focused card has a dark-blue focus ring (`#005FB8`)
- Computed contrast ratio ≥ 3.0 (WCAG AA non-text threshold)

**Pass criteria:**
- `contrast_ratio(outline_color, card_bg) >= 3.0`
- `detail` field includes the computed ratio

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                                          | Pass condition                                        | Why                                                                    |
|-----|-----------------------------------|----------------------------------------------------------|-------------------------------------------------------|------------------------------------------------------------------------|
| 09  | `09-tc006a-gallery-focus.png`     | Gallery view with first card focused, focus ring visible | Dark-blue outline clearly surrounds the focused card  | Proves keyboard focus indicator meets WCAG non-text contrast (≥3:1)   |

---

### TC-006b – Focus ring contrast ≥ 3:1 in List view (Light theme)

**Preconditions:**
- VS Code is open with ATK extension installed from the `fix/issue-15916-copilot` branch build

**Steps:**
1. Set VS Code color theme to "Default Light Modern" via Command Palette (`Ctrl+Shift+P` → `Preferences: Color Theme` → select "Default Light Modern").
2. Click the ATK icon in the VS Code Activity Bar to activate the extension panel.
3. Open Command Palette (`Ctrl+Shift+P`) and run `Microsoft 365 Agents Toolkit: View Samples`. Observe the Sample Gallery opens in Gallery (Grid) view.
4. Click the **List** view toggle button to switch to List layout. Observe the layout changes to a vertical list of items.
5. Click inside the Sample Gallery webview area to give it keyboard focus.
6. Press **Tab** to move keyboard focus to the first `.sample-list-item`. Observe the focus ring appears around the list item.
7. Take screenshot **immediately after Tab** so the focus ring is clearly visible.
8. Read `getComputedStyle(listItem).outlineColor` from the focused list item.
9. Find the effective background color of the list item by reading `getComputedStyle` up the DOM until a non-transparent `backgroundColor` is found (expected: page background white `#FFFFFF` since featured section no longer has a grey background).
10. Compute `contrast_ratio(outline_color, effective_background)` using the WCAG relative luminance formula.

**Expected result:**
- All list items (featured and non-featured) share the same page background since the featured section background was removed.
- Focused list item has a dark-blue focus ring (`#005FB8`, ~10:1 against white background).
- Computed contrast ratio ≥ 3.0.

**Pass criteria:**
- `contrast_ratio(outline_color, effective_background) >= 3.0`
- `detail` field includes the computed ratio, the outline color, and the resolved background color

**Screenshots produced by test:**

| ID  | Filename                        | What is visible                                                         | Pass condition                                              | Why                                                                              |
|-----|---------------------------------|-------------------------------------------------------------------------|-------------------------------------------------------------|----------------------------------------------------------------------------------|
| 10  | `10-tc006b-list-focus.png`      | List view with first list item focused, focus ring visible              | Dark-blue outline clearly surrounds the focused list item   | Proves focus ring is visible against white page background (≥3:1)               |

---

## Test Script

`packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Notes

- **Extension activation**: Each TC begins with clicking the ATK Activity Bar icon to guarantee the extension is activated before any command is run. The test harness should install the extension VSIX built from the `fix/issue-15916-copilot` branch before running.
- The Sample Gallery renders inside a VS Code webview. DOM evaluation is performed via the Playwright CDP session targeting the webview frame.
- All contrast calculations use the WCAG relative luminance formula (IEC 61966-2-1 sRGB, not a simple color blacklist).
- For TC-001a/b: the link must be **focused** (`element.focus()`) before reading computed color — the `:focus` CSS rule changes the link color in Light theme (≥4.5:1 contrast). Dark theme uses `#4FC3F7` (≈8.32:1 on `#1E1E1E`), applied with `!important` to override Fluent UI inline styles.
- For TC-004a/b: The new design removes the featured section background color. Featured items are now distinguished by a `codicon-star-full` icon — overlaid on the top-left of the card thumbnail in gallery view, and prepended to the title text in list view. The icon color uses `var(--vscode-editorLightBulb-foreground, #DDB100)` which adapts to the active VS Code theme. TC-004a checks gallery star presence; TC-004b checks list star presence.
- For TC-006a/b: `element.focus()` triggers `:focus-visible` styles in VSCode webview. In the light theme:
  - `.sample-card:focus-visible` uses `outline-color: #005FB8` (~10:1 against white card background).
  - `.sample-list-item:focus-visible` uses `outline: 2px solid #005FB8` (~10:1 against white background).
  - Since the featured section no longer has a grey background, a single focus-ring rule applies to all list items.
  - The screenshot step must immediately follow focus so the ring is still visible.
- `aria-pressed` on `<vscode-button>` is forwarded to the inner `<button>` element by the FAST foundation runtime.
- TC-002, TC-003, TC-005: these attributes are set inline in the React render path and are accessible without CDP frame evaluation.
