---
name: atk-copilot-test-generator
description: >
  ATK generator agent: reads a test plan and generates .test.ts code for the
  @vscode/test-electron + Playwright CDP test framework. Used by atk-copilot-test-generator.yml.
---

# ATK Copilot Test Generator Skill

## Role

You are the **test code generator**. The label agent has already:
1. Fixed any product code needed.
2. Created/updated the test plan.
3. Committed and pushed to a branch.

Your job: read that test plan → write `.test.ts` code that faithfully implements it →
commit and push → write the script basename to `/tmp/script-name.txt`.

You do NOT run the test. The runner workflow handles execution.

---

## Workflow Overview

Every test follows this top-level sequence:

```
Open VS Code  →  Step 0: Verify ATK activated  →  Route-specific steps  →  Write results.json
```

**Success condition per route:**

| Route | Auth needed | Primary success signal | Required screenshots | Guide |
|---|---|---|---|---|
| **Launch** | No | ATK extension active | ATK sidebar pane (element) | [routes/launch.md](routes/launch.md) |
| Scaffold | No | All expected files exist under project dir | wizard QuickPick (element), file-check in explorer (element) | [routes/scaffold.md](routes/scaffold.md) |
| Local Debug | M365 | Debug session attaches | accounts pane (element), debug toolbar (element) | [routes/local-debug.md](routes/local-debug.md) |
| Remote Provision | M365 + Azure | Provision completes notification | accounts pane (element), notification toast (element) | [routes/provision-deploy.md](routes/provision-deploy.md) |
| Remote Deploy | M365 + Azure | Deploy completes notification | accounts pane (element), notification toast (element) | [routes/provision-deploy.md](routes/provision-deploy.md) |

**Always take element-level screenshots** (not full-page) — see [test-code-helpers.md](test-code-helpers.md).

---

## Step-by-Step: Generate Test Code

### Step 1 — Find the test plan

```bash
git log --oneline -5
# Find plan added/modified by the label agent:
git show --name-only HEAD -- packages/tests/copilot-test/test-plans/ | grep '\.md$' || \
git diff HEAD~1 --name-only -- packages/tests/copilot-test/test-plans/ | grep '\.md$'
cat <test-plan-path>
```

### Step 2 — Read infrastructure and route docs

```bash
cat packages/tests/copilot-test/README.md
cat packages/tests/copilot-test/skills/test-code-helpers.md   # boilerplate + signal types
cat packages/tests/copilot-test/skills/routes/launch.md       # VS Code launch
cat packages/tests/copilot-test/skills/routes/scaffold.md     # scaffold route
cat packages/tests/copilot-test/skills/routes/auth.md         # auth flow
cat packages/tests/copilot-test/skills/routes/local-debug.md  # debug route
cat packages/tests/copilot-test/skills/routes/provision-deploy.md
```

### Step 3 — Write the test file

- Naming: `packages/tests/copilot-test/src/<feature>-<task>.test.ts`
- Reference implementation: `packages/tests/copilot-test/src/simple-bot-create.test.ts`
- **Every test plan step must map to code.** No skipping steps, no static-only assertions.
- Step 0 (`waitForATKActivation`) is **mandatory** in every suite — see [test-code-helpers.md](test-code-helpers.md).

### Step 4 — Save the script basename

```bash
# Write just the basename WITHOUT .test.ts extension:
echo "sample-app-create" > /tmp/script-name.txt
```

### Step 5 — Commit and push

```bash
git add packages/tests/copilot-test/src/
git commit -m "test(generated): <feature> test script (issue #$ISSUE)"
git push origin $(git branch --show-current)
```

---

## Reference documents

| Document | Contents |
|---|---|
| [test-code-helpers.md](test-code-helpers.md) | Imports, dir constants, `takeElementScreenshot`, `sendSignal`, `waitForATKActivation`, signal type table |
| [test-quality-rules.md](test-quality-rules.md) | Mandatory quality rules: no skipped steps, real UI actions, contrast formula, state-change verification |
| [routes/launch.md](routes/launch.md) | Launch VS Code: local dev (`--extensionDevelopmentPath`), VSIX/remote CI, Windows detach, workspace prereq |
| [routes/scaffold.md](routes/scaffold.md) | Scaffold route: wizard pattern, `templates/src/ui/` labels, `findScaffoldedDir`, file verification |
| [routes/auth.md](routes/auth.md) | Auth: source patches, env vars, step-by-step auth flow code |
| [routes/local-debug.md](routes/local-debug.md) | Local Debug route: preconditions, debug picker, toolbar screenshot |
| [routes/provision-deploy.md](routes/provision-deploy.md) | Remote Provision and Deploy routes |

