# Test Plan: Sample App Panel Accessibility (A11y)

## Metadata

- **feature-slug**: `sample-app-a11y`
- **owner**: atk-qa
- **created**: 2026-05-15
- **updated**: 2026-05-26
- **triggers**: issue-label `atk-copilot-test`, manual
- **related-issue**: [#15916](https://github.com/OfficeDev/microsoft-365-agents-toolkit/issues/15916)

## Core Principle

Tests simulate **real user behavior**: every step describes what a user clicks, types, tabs to, or observes — not internal DOM attributes or CSS values. Screenshots capture what the user actually sees on screen.

---

## Scope

**Covers:**
- TC-001: User can read sample card links clearly in Light theme (WCAG AA contrast >= 4.5:1)
- TC-001b: Same contrast check in List view layout
- TC-002: User relying on a screen reader hears "Featured sample" prefix for featured cards
- TC-003: User can distinguish the Featured badge visually in Light theme (non-text contrast >= 3:1)
- TC-004: User tabbing through cards hears the sample''s tags announced
- TC-005: User can tell which view toggle (Gallery/List) is currently active via aria-pressed
- TC-006: User navigating by keyboard sees a clearly visible focus ring in Light theme

**Does NOT cover:**
- Full keyboard navigation flow (Tab order across entire panel)
- Live screen reader announcement testing (NVDA/JAWS)
- Mobile or touch accessibility
- High Contrast theme

---

## Test Cases

### TC-001 - Link text is readable in Light theme (Gallery view)

**Preconditions:**
- VS Code is open, ATK extension v6.8.0+ activated
- VS Code theme is set to a Light variant (e.g. Default Light Modern)

**Steps:**
1. User opens Command Palette and runs `Microsoft 365 Agents Toolkit: View Samples`
2. User observes the Sample Gallery panel opens with a grid of sample cards
3. User reads the blue link text on sample cards and confirms it is clearly legible against the white background
4. Take screenshot

**Expected result:**
- Sample card links appear in dark blue, readable on white

**Pass criteria:**
- `contrast_ratio(computed link color, background color) >= 4.5`
- At least one `.ms-Link` element present in the webview

**Screenshots produced by test:**

| ID  | Filename                     | What is visible                                 | Pass condition                               | Why                                                      |
|-----|------------------------------|-------------------------------------------------|----------------------------------------------|----------------------------------------------------------|
| 01  | `01-extension-active.png`    | VS Code with ATK extension activated            | ATK sidebar icon visible, extension loaded   | Baseline: proves extension is ready before gallery opens |
| 02  | `02-gallery-open.png`        | Sample Gallery panel with grid of sample cards  | Gallery renders with blue link text on white | Proves gallery opens successfully after user command     |
| 03  | `03-tc001-link-contrast.png` | Close-up of sample card link text in light theme | Link text is dark blue, clearly readable    | Proves light-theme link color fix is applied             |

---

### TC-001b - Link text is readable in Light theme (List view)

**Preconditions:**
- Same as TC-001; List view is active

**Steps:**
1. User opens Sample Gallery (Gallery view active by default)
2. User clicks the **List** view toggle button at the top of the panel
3. User observes the layout switches to a list of sample rows
4. User reads the blue link text in the list and confirms legibility
5. Take screenshot

**Expected result:**
- Sample card links in list view appear in dark blue, readable on white

**Pass criteria:**
- `contrast_ratio(computed link color, background color) >= 4.5` in list layout
- At least one `.ms-Link` element present in list view

**Screenshots produced by test:**

| ID  | Filename                          | What is visible                               | Pass condition                               | Why                                                          |
|-----|-----------------------------------|-----------------------------------------------|----------------------------------------------|--------------------------------------------------------------|
| 04  | `04-tc001b-link-contrast-list.png`| Sample Gallery in List view, link text visible | Link text is dark blue in list layout       | Proves color fix applies to list view as well as gallery view |

---

### TC-002 - Screen reader users hear "Featured sample" prefix on featured cards

**Preconditions:**
- Sample Gallery is open in either Grid or List view
- At least one Featured and one non-Featured card are visible

**Steps:**
1. User opens Sample Gallery
2. User (simulating a screen reader) inspects the accessible name of a card in the "Featured" section
3. User confirms the name begins with "Featured sample." so they know it is featured
4. User inspects a non-featured card and confirms no "Featured sample." prefix
5. Take screenshot

**Expected result:**
- Featured card accessible name: `"Featured sample. <Title>. Tags: <tag1>, <tag2>"`
- Non-featured card accessible name: `"<Title>. Tags: <tag1>, <tag2>"`

**Pass criteria:**
- >= 1 card with `aria-label` starting with `"Featured sample."`
- >= 1 card with `aria-label` NOT starting with `"Featured sample."`

**Screenshots produced by test:**

| ID  | Filename                   | What is visible                                    | Pass condition                              | Why                                                        |
|-----|----------------------------|----------------------------------------------------|---------------------------------------------|------------------------------------------------------------|
| 05  | `05-tc002-aria-labels.png` | Gallery showing featured badge on featured cards   | "Featured" badge is visually present        | Proves ARIA differentiation matches visual differentiation |

---

### TC-003 - Featured badge is distinguishable in Light theme

**Preconditions:**
- Sample Gallery open in Light theme

**Steps:**
1. User opens Sample Gallery in Light theme
2. User looks at the gold/brown "Featured" badge on a featured card
3. User confirms the badge text and star icon are legible against the card''s white/light background
4. Take screenshot

**Expected result:**
- Featured badge uses `#7A5C00` (dark gold) on white -- contrast ~4.9:1

**Pass criteria:**
- `contrast_ratio("#7A5C00", "#FFFFFF") >= 3.0`
- At least one `.featured-badge` element found in the webview

**Screenshots produced by test:**

| ID  | Filename                       | What is visible                                     | Pass condition                                  | Why                                                           |
|-----|--------------------------------|-----------------------------------------------------|-------------------------------------------------|---------------------------------------------------------------|
| 06  | `06-tc003-badge-contrast.png`  | Featured card with gold/brown badge in light theme  | Badge text is visually distinct from background | Proves featured badge meets WCAG non-text contrast threshold  |

---

### TC-004 - Keyboard users hear sample tags when navigating cards

**Preconditions:**
- Sample Gallery is open

**Steps:**
1. User opens Sample Gallery
2. User presses **Tab** to move keyboard focus onto the first sample card
3. User (or screen reader) reads the full accessible name of the focused card -- expects to hear the title followed by its tags
4. Take screenshot of the focused card

**Expected result:**
- Accessible name: `"<Title>. Tags: <tag1>, <tag2>, <tag3>"`

**Pass criteria:**
- Every sample card `aria-label` contains `". Tags:"` followed by at least one tag name

**Screenshots produced by test:**

| ID  | Filename                  | What is visible                                         | Pass condition                                        | Why                                                            |
|-----|---------------------------|---------------------------------------------------------|-------------------------------------------------------|----------------------------------------------------------------|
| 07  | `07-tc004-tags-aria.png`  | Gallery with first card focused, tags visible in label  | Focused card has visible focus ring; tags in a11y name | Proves tags are included in accessible name for screen readers |

---

### TC-005 - Gallery/List toggle buttons communicate their active state

**Preconditions:**
- Sample Gallery is open (Gallery view active by default)

**Steps:**
1. User opens Sample Gallery -- Gallery toggle button appears highlighted/active
2. User clicks the **List** view button
3. User observes layout switches to list style and List button appears highlighted/active
4. User clicks the **Gallery** view button to switch back
5. User observes Gallery button is highlighted/active again
6. Take screenshots before and after click

**Expected result:**
- Active toggle: `aria-pressed="true"` (highlighted)
- Inactive toggle: `aria-pressed="false"` (not highlighted)

**Pass criteria:**
- Before click: Gallery button `aria-pressed="true"`, List button `aria-pressed="false"`
- After clicking List: List button `aria-pressed="true"`, Gallery button `aria-pressed="false"`

**Screenshots produced by test:**

| ID  | Filename                      | What is visible                                        | Pass condition                                  | Why                                                             |
|-----|-------------------------------|--------------------------------------------------------|-------------------------------------------------|-----------------------------------------------------------------|
| 08  | `08-tc005-toggle-before.png`  | Gallery layout active, Gallery button visually selected | Gallery button highlighted, List button not   | Proves initial aria-pressed state matches visual active state   |
| 09  | `09-tc005-toggle-after.png`   | List layout active, List button visually selected      | List button highlighted, Gallery button not     | Proves aria-pressed toggles correctly in response to user click |

---

### TC-006 - Keyboard focus ring is visible on sample cards in Light theme

**Preconditions:**
- Sample Gallery is open
- VS Code theme is set to a Light variant

**Steps:**
1. User opens Sample Gallery in Light theme
2. User presses **Tab** to move keyboard focus to the first sample card
3. User observes a dark-blue focus ring appears around the focused card
4. User confirms the focus ring is clearly visible against the white card background
5. Take screenshot

**Expected result:**
- Focused card has a clearly visible dark-blue focus ring (`#005FB8`)
- Contrast of `#005FB8` on white ~5.77:1 (WCAG AA non-text threshold: 3:1)

**Pass criteria:**
- Computed `outline-color` of focused `.sample-card` is `#005FB8` in light theme
- `contrast_ratio("#005FB8", "#FFFFFF") >= 3.0`

**Screenshots produced by test:**

| ID  | Filename                   | What is visible                                        | Pass condition                                       | Why                                                                |
|-----|----------------------------|--------------------------------------------------------|------------------------------------------------------|--------------------------------------------------------------------|
| 10  | `10-tc006-focus-ring.png`  | First sample card focused, dark-blue outline shown     | Focus ring is dark blue and clearly distinguishable  | Proves keyboard users can identify focused card in light theme     |

---

## Test Script

`packages/tests/copilot-test/src/sample-app-a11y.test.ts`

## Notes

- The Sample Gallery renders inside a sandboxed VS Code webview iframe. DOM evaluation uses the Playwright CDP session.
- Contrast ratios in pass criteria use the WCAG relative luminance formula.
- `aria-pressed` on `<vscode-button>` is forwarded to the inner `<button>` by the FAST foundation runtime.
- TC-006: VS Code applies `:focus-visible` styles; `element.focus()` from test code triggers the ring.