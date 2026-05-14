# Test Plans

This directory contains test plans for Copilot-driven ATK (Microsoft 365 Agents Toolkit) tests.
Each subdirectory corresponds to a feature area and contains one or more test case definitions.

## Adding a new test plan

1. Copy `template.md` to `test-plans/<feature-slug>/<feature-slug>.md`
2. Fill in all placeholders and remove the comment blocks
3. The Copilot CLI agent reads this file at runtime to understand what to test
4. Create the corresponding TypeScript test in:
   `packages/tests/src/<feature-slug>-<task>.test.ts`

## Directory structure

```
test-plans/
  README.md                     ← this file
  template.md                   ← copy this for new test plans
  simple-bot/
    simple-bot.md               ← TC-001: create Simple Bot via wizard
  <feature-slug>/
    <feature-slug>.md           ← one file per feature area, multiple TCs inside
```

## Feature slug naming

Use lowercase-hyphen names that match the ATK feature being tested:

| Slug                   | Feature area                                     |
|------------------------|--------------------------------------------------|
| `simple-bot`           | Simple Bot template creation via wizard          |
| `notification-bot`     | Notification Bot template                        |
| `custom-engine-agent`  | Custom Engine Agent (CEA) wizard                 |
| `message-extension`    | Message Extension template                       |
| `atk-cli`              | ATK CLI (`atktk new`, `atktk provision`, …)      |
| `treeview`             | ATK sidebar tree-view (Lifecycle, Accounts, …)   |

## Test plan format

Each `.md` file follows the structure defined in `template.md`:

```
# Test Plan: <Feature Name>
## Metadata        – feature-slug, owner, dates, triggers
## Scope           – what is and is not covered
## Test Cases      – one ### block per TC, each with:
   - Preconditions
   - Wizard flow table (for UI tests) or CLI commands (for CLI tests)
   - Steps
   - Expected result
   - Test script path
   - Screenshots table
## Notes           – timing quirks, known flakiness, etc.
```

## How Copilot uses these plans

When a GitHub issue is labelled `atk-copilot-test`, the Copilot CLI agent:

1. Reads the issue body to identify the feature area and relevant TC(s)
2. Locates (or creates) the matching test plan in this directory
3. Generates or updates the TypeScript test in `packages/tests/src/`
4. Runs the test inside Docker and collects screenshots + `results.json`
5. Posts a summary comment to the issue with pass/fail table and inline GIF
6. Removes the `atk-copilot-test` label and adds `atk-copilot-test:done`
