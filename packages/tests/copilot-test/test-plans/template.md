# Test Plan: <Feature Name>

<!--
  HOW TO USE THIS TEMPLATE
  ========================
  1. Copy this file to:  test-plans/<feature-slug>/<feature-slug>.md
  2. Fill in every <placeholder>.
  3. Remove all comment blocks (<!-- ... -->) before committing.
  4. The Copilot CLI agent reads this file at runtime to generate & run the
     corresponding TypeScript test in:
       packages/tests/src/<feature-slug>-<task>.test.ts

  CORE PRINCIPLE: Tests must simulate REAL USER BEHAVIOR
  ======================================================
  Every test case must describe what a real user does — not what the code does internally.

  ✓ GOOD (user action):  "Click the List view button, observe layout changes to list style"
  ✗ BAD  (code check):   "Assert CSS class 'list-view' is present on the container"

  ✓ GOOD (user action):  "Tab to the first sample card, press Enter to open it"
  ✗ BAD  (code check):   "Check that aria-label attribute contains the word 'Featured'"

  Steps MUST be written from the user's perspective:
  - Use "click", "type", "press", "navigate", "observe", "verify visually"
  - Every step must state what the user does AND what they expect to see immediately after
  - Screenshots must capture the UI state a user would see — not internal DOM state
-->

## Metadata

- **feature-slug**: `<feature-slug>`        <!-- lowercase-hyphen, matches folder name -->
- **owner**: `<alias>`
- **created**: YYYY-MM-DD
- **updated**: YYYY-MM-DD
- **triggers**: issue-label `atk-copilot-test`, manual

## Scope

**Covers:**
- <What the test exercises, one bullet per concern>

**Does NOT cover:**
- <Explicitly excluded scenarios>

---

## Test Cases

### TC-001 – <Short description of the happy path>

**Preconditions:**
- VSCode is open with no project loaded
- ATK extension v6.8.0+ is installed and activated
- <Any other pre-condition>

<!--
  WIZARD FLOW TABLE (for UI wizard tests)
  Fill in the QuickPick step order as verified against the real extension.
  If the test is CLI-only, replace this table with the CLI command(s).
-->
**Wizard flow:**

| Step | QuickPick / InputBox      | Value to select / type |
|------|---------------------------|------------------------|
| 1    | <First QuickPick label>   | <Option to click>      |
| 2    | <Second QuickPick label>  | <Option to click>      |
| N    | Application Name          | `<app-name>` (type + Enter) |

**Steps:**
1. <First action>
2. <Second action>
3. …

**Expected result:**
- <What the test asserts>
- <File(s) or UI state(s) that must be present>

**Pass criteria:**
- <Measurable assertion 1, e.g. `contrast_ratio >= 4.5`>
- <Measurable assertion 2, e.g. `aria-label contains "Featured sample"`>

**Test script:**
`packages/tests/src/<feature-slug>-<task>.test.ts`

**Screenshots produced by test:**

<!--
  One row per named screenshot. The test calls `screenshot("id")` before each
  meaningful user action so the UI state is captured while still visible.
  IDs must be sequential two-digit strings (01, 02, …).
  Pass condition: what must be TRUE in this screenshot for the TC to pass.
  Why: why this screenshot proves the behavior (what regression would make it fail).
-->

| ID  | Filename                  | What is visible                   | Pass condition                              | Why                                              |
|-----|---------------------------|-----------------------------------|---------------------------------------------|--------------------------------------------------|
| 01  | `01-<name>.png`           | <UI state after user action>      | <e.g. button is highlighted / text appears> | <e.g. proves feature activated after user click> |
| 02  | `02-<name>.png`           | <UI state after user action>      | <e.g. error message is absent>              | <e.g. proves no regression in happy path>        |

---

### TC-002 – <Short description of an error / edge case> _(optional)_

**Preconditions:** <Same as TC-001 unless noted.>

**Steps:**
1. …

**Expected result:**
- …

**Test script:**
`packages/tests/src/<feature-slug>-<task2>.test.ts`

---

## Notes

<!--
  Add any implementation notes that help the Copilot agent understand quirks,
  e.g. timing issues, known flakiness, platform differences.
-->

- <Note 1>
