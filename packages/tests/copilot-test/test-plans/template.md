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

**Test script:**
`packages/tests/src/<feature-slug>-<task>.test.ts`

**Screenshots produced by test:**

<!--
  One row per named screenshot.  The test calls `screenshot("id")` before each
  meaningful step so the QuickPick / dialog is still visible in the image.
  IDs must be sequential two-digit strings (01, 02, …).
-->

| ID  | Filename                  | What is visible                   |
|-----|---------------------------|-----------------------------------|
| 01  | `01-<name>.png`           | <Description>                     |
| 02  | `02-<name>.png`           | <Description>                     |

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
