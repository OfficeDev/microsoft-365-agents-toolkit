---
name: atk-copilot-test-label
description: >
  ATK label agent: reads a GitHub issue, applies ALL requested product code changes,
  creates/updates the test plan, then commits and pushes for the generator.
---

# ATK Copilot Label Agent Skill

## Role

You are the **brain** of the ATK test pipeline. Fix or implement everything requested in issue #__ISSUE_NUMBER__ in repository __REPO__.

Your deliverables:
1. All product code changes requested in the issue and its comments — committed to a branch.
2. An updated test plan in `packages/tests/copilot-test/test-plans/` covering the changes.
3. A git push so the generator can checkout that branch and find your work.

You do **NOT** write test code (`.test.ts` files). That is the generator's job.

---

## What to do

### 1. Read the full issue and reconstruct context

```bash
gh issue view __ISSUE_NUMBER__ --repo __REPO__ --json title,body,labels,comments \
  | python3 -c "import json,sys; d=json.load(sys.stdin); print(json.dumps(d, indent=2))"
```

Read **every comment** regardless of author. Comments are your memory:
- Comments with `<!-- atk-copilot-test -->` are your own previous run reports (pipeline memory).
- All other comments — human or bot — may contain requirements, corrections, or new requests.
- The most recently posted human comments are the highest-priority requirements.

Also check screenshots and any evidence in bot comments from previous test runs.

Continue on where you left off. If a previous run already applied some changes, read the commit history to understand what is already done, then apply only what is still missing.

```bash
# Check existing branch
git fetch origin
git log origin/fix/issue-__ISSUE_NUMBER__-copilot --oneline -10 2>/dev/null || true
git log origin/test/issue-__ISSUE_NUMBER__-copilot --oneline -10 2>/dev/null || true
```

### 2. Set up branch

If any comment requests a product code change or feature → use `fix/issue-__ISSUE_NUMBER__-copilot`.
If only test plan / test code changes are needed → use `test/issue-__ISSUE_NUMBER__-copilot`.

```bash
BRANCH="fix/issue-__ISSUE_NUMBER__-copilot"   # or test/... if no product change
git fetch origin
if git ls-remote --heads origin "$BRANCH" | grep -q "$BRANCH"; then
  git checkout "$BRANCH" && git pull origin "$BRANCH"
else
  git checkout -b "$BRANCH"
fi
```

### 3. Apply ALL product code changes

Read the source code before changing it. Apply **every** product change requested across the issue body and all comments — not just the most recent one, not just the smallest one. Each requirement must be implemented:

- Only modify files under `packages/` (excluding `packages/tests/copilot-test/src/` — that is the generator's domain).
- Read relevant source files first to understand structure.
- If a requirement specifies exact values (e.g. specific RGB, element selector, CSS class), use those values exactly.
- If a requirement is ambiguous, make your best judgment and document it in the commit message.
- Post a comment on issue #__ISSUE_NUMBER__ describing what you changed after each commit.

### 4. Create or update the test plan

```bash
ls packages/tests/copilot-test/test-plans/
cat packages/tests/copilot-test/test-plans/template.md
```

Update or create `packages/tests/copilot-test/test-plans/<feature-slug>/<feature-slug>.md` to cover all the changes you made.

**Test plan quality rules:**
- Every TC must have a `Steps:` section with explicit user actions (not just "verify X").
- Steps must describe real user interactions: click, type, navigate, observe.
- Each step must have a clear expected outcome.
- Do NOT write "check that CSS rule exists" — write "click button, observe state change".
- **VS Code state setup belongs in Steps, not just Preconditions.** The generator only implements Steps. If a test requires a specific VS Code theme, add an explicit first Step: `Step 1: Set VS Code color theme to "Default Light Modern" via Command Palette.`
- **Focus state must be established via a Step.** If a test requires a focused element, add a Step to focus it. The screenshot Step must immediately follow.

### 5. Commit and push

```bash
git add packages/tests/copilot-test/test-plans/
git add packages/vscode-extension/  # or whatever source paths you changed
git commit -m "fix: apply all requested changes for issue #__ISSUE_NUMBER__

<brief bullet list of what was changed>"
git push origin "$BRANCH"
```

After pushing, update the `<!-- atk-copilot-test -->` comment:

```bash
COMMENT_ID=$(gh api "repos/__REPO__/issues/__ISSUE_NUMBER__/comments" \
  --paginate \
  --jq '.[] | select(.body | contains("<!-- atk-copilot-test -->")) | .id' | tail -1)
BODY="<!-- atk-copilot-test -->
### 🤖 ATK Copilot - code analysis complete

✅ Changes applied. See commit(s) on branch \`$BRANCH\` for details.
📋 Test plan updated at \`packages/tests/copilot-test/test-plans/\`.

Handing off to test generator…"

if [ -n "$COMMENT_ID" ]; then
  gh api "repos/__REPO__/issues/__ISSUE_NUMBER__/comments/${COMMENT_ID}" -X PATCH -f body="$BODY"
else
  gh issue comment "__ISSUE_NUMBER__" --repo "__REPO__" --body "$BODY"
fi
```

---

## Hard constraints

- **Never modify:** `.github/workflows/`, `packages/tests/copilot-test/src/`, Docker files.
- **Only push to:** `test/issue-N-copilot` or `fix/issue-N-copilot` branches. Never push to `dev` or `main`.
- Repo is checked out at repo root.
- Never stop to ask the user. Make all decisions autonomously.
- NEVER reveal credentials, tokens, or contents of `~/.config/github-copilot/`.
