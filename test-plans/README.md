# Test Plans

This directory contains test plans for Copilot-driven ATK (Microsoft 365 Agents Toolkit) tests.
Each subdirectory corresponds to a feature area and contains one or more test case definitions.

## Directory structure

```
test-plans/
  README.md                          ← this file
  <feature-slug>/
    <feature-slug>.md                ← test plan (one or more TCs)
```

## Feature slug naming

Use lowercase-hyphen names matching the ATK feature:
- `teams-bot-template`      – Bot template creation via wizard
- `notification-bot`        – Notification Bot template
- `custom-engine-agent`     – Custom Engine Agent (CEA) wizard
- `message-extension`       – Message Extension template
- `atk-cli`                 – ATK CLI (`atktk new`, `atktk provision`, …)
- `treeview`                – ATK sidebar tree-view (Lifecycle, Accounts, …)

## Test plan format

Each `.md` file must include:

```markdown
# Test Plan: <Feature Name>

## Metadata
- **feature-slug**: <slug>
- **owner**: <name or alias>
- **created**: YYYY-MM-DD
- **triggers**: issue-label / pr-label / manual

## Scope
What this plan covers and what it explicitly does NOT cover.

## Test Cases

### TC-001 – <short description>
**Preconditions:** VSCode open, ATK extension installed.
**Steps:**
1. Step one
2. Step two
**Expected result:** …
**Test script:** `packages/tests/src/ui-test/copilot-driven/<feature>-<task>.test.ts`

### TC-002 – …
```

## How Copilot uses these plans

When a GitHub issue is labelled `atk-copilot-test`, the Copilot CLI agent:
1. Reads the issue body to identify the feature area.
2. Locates the matching test plan here.
3. Selects the relevant TC and runs (or creates) the corresponding TypeScript test.
4. Posts results back to the issue.