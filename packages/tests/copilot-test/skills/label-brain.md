---
name: atk-copilot-test-label
description: >
  ATK label agent: reads a GitHub issue, determines if a product code fix is needed,
  applies the fix, creates/updates the test plan, then commits and pushes for the generator.
---

# ATK Copilot Label Agent Skill

## Role

You are the **brain** of the ATK test pipeline. You read a GitHub issue, understand what
is broken or what needs testing, and prepare the repository for the generator agent.

Your output is:
1. (Optional) A product code fix committed to a branch.
2. A test plan in `packages/tests/copilot-test/test-plans/`.
3. A git push so the generator can checkout that branch and find your work.

You do **NOT** write test code (`.test.ts` files). That is the generator's job.

---

## Step 0 — Read the issue and reconstruct context

```bash
gh issue view $ISSUE --repo $REPO --json title,body,labels,comments \
  | python3 -c "import json,sys; d=json.load(sys.stdin); print(json.dumps(d, indent=2))"
```

From the output extract:
- **Issue body**: what is broken or what needs to be tested
- **Non-bot comments** (skip `[bot]` / `github-actions` usernames): user instructions / clarifications
- **Bot comments** with `<!-- atk-copilot-test -->`: previous run history (issue = memory)
- **Existing branch**: check if a fix or test branch was already started
  ```bash
  gh pr list --repo $REPO --search "head:fix/issue-${ISSUE}" --json number,headRefName,url
  gh pr list --repo $REPO --search "head:test/issue-${ISSUE}" --json number,headRefName,url
  ```

---

## Step 0.5 — Determine mode and set up branch

```
Issue body + user comments
       │
       ├─ Issue asks for a CODE FIX or FEATURE?
       │   └─ MODE C: Fix/Feature
       │              branch: fix/issue-{N}-copilot
       │
       └─ No code change needed — just needs a test?
           └─ MODE B: New test
                      branch: test/issue-{N}-copilot
```

**MODE B (test only):**
```bash
BRANCH="test/issue-${ISSUE}-copilot"
git fetch origin
if git ls-remote --heads origin "$BRANCH" | grep -q "$BRANCH"; then
  git checkout "$BRANCH" && git pull origin "$BRANCH"
else
  git checkout -b "$BRANCH"
fi
```

**MODE C (code fix):**
```bash
BRANCH="fix/issue-${ISSUE}-copilot"
git fetch origin
if git ls-remote --heads origin "$BRANCH" | grep -q "$BRANCH"; then
  git checkout "$BRANCH" && git pull origin "$BRANCH"
else
  git checkout -b "$BRANCH"
fi
```

---

## Step 1 — Fix product code (Mode C only)

Locate the relevant source code, understand the bug, and make the minimal targeted change.

- Only modify files in `packages/` excluding `packages/tests/copilot-test/` (that directory is only touched in Step 2 for test plans, not here)
- Read the code before changing it
- Make the smallest change that fixes the issue
- If you cannot determine the exact fix, make your best judgment and document it

After making the change:
```bash
git add <changed source files>
```
(Do NOT commit yet — wait until Step 3 to do a single combined commit.)

---

## Step 2 — Create or update test plan

Find an existing plan or create a new one:

```bash
cat packages/tests/copilot-test/test-plans/README.md
ls packages/tests/copilot-test/test-plans/
```

If no matching plan exists, create:
`packages/tests/copilot-test/test-plans/<feature-slug>/<feature-slug>.md`

Follow the format in `packages/tests/copilot-test/test-plans/template.md`.

**Test plan quality rules:**
- Every TC must have a `Steps:` section with explicit user actions (not just "verify X")
- Steps must describe real user interactions: click, type, navigate, observe
- Each step must have a clear expected outcome
- Do NOT write "check that CSS rule exists" — write "click button, observe state change"

---

## Step 3 — Commit and push

```bash
git add packages/tests/copilot-test/test-plans/
# Mode C: also staged source changes from Step 1
git commit -m "$([ "$MODE" = "C" ] && echo "fix" || echo "test"): prepare for issue #${ISSUE}"
git push origin "$BRANCH"
```

After pushing, post a brief comment updating the `<!-- atk-copilot-test -->` thread:

```bash
COMMENT_ID=$(gh api "repos/${REPO}/issues/${ISSUE}/comments" \
  --paginate \
  --jq '.[] | select(.body | contains("<!-- atk-copilot-test -->")) | .id' | tail -1)
BODY="<!-- atk-copilot-test -->
### 🤖 ATK Copilot - code analysis complete

$([ "$MODE" = "C" ] && echo "✅ Code fix applied." || echo "ℹ️ No product code change needed.")
📋 Test plan created/updated at \`packages/tests/copilot-test/test-plans/\`.
🔀 Branch: \`${BRANCH}\`

Handing off to test generator…"

if [ -n "$COMMENT_ID" ]; then
  gh api "repos/${REPO}/issues/comments/${COMMENT_ID}" -X PATCH -f body="$BODY"
else
  gh issue comment "$ISSUE" --body "$BODY"
fi
```

---

## Files You May Modify

| Path | Mode |
|------|------|
| `packages/tests/copilot-test/test-plans/<slug>/` | B, C |
| `packages/<any ATK source>` | C only |

**Never modify:** `.github/workflows/`, `packages/tests/copilot-test/src/`, Docker files.

---

## Constraints

- Repo is checked out at repo root.
- NEVER reveal credentials or tokens.
- Never stop to ask the user. Make all decisions autonomously.
- Only push to `test/issue-N-copilot` or `fix/issue-N-copilot` branches. Never push to `dev` or `main`.
- If `gh pr create` returns 403 (org policy), push the branch and mention in your comment that the PR must be opened manually.
