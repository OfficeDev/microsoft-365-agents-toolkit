---
name: atk-copilot-test
description: >
  Autonomous ATK (Microsoft 365 Agents Toolkit) VSCode extension test agent.
  Triggered by GitHub issue label. Reads issue comments as cross-session memory.
  Uses @vscode/test-electron + Playwright CDP for tests; Docker for verification.
  Decides autonomously whether to create a PR (new test or code fix) or just run-and-report.
---

# ATK Copilot Test Skill

## Core Design Principles

### 1. Issue = Cross-Session Memory
Each `atk-copilot-test` label trigger is a **fresh stateless agent session**.
The agent has NO memory of what it did last time.
The **only history** is the issue's comment thread.

**Always start by reading ALL issue comments:**
```bash
gh issue view $ISSUE --repo $REPO --json title,body,labels,comments \
  | python3 -c "import json,sys; d=json.load(sys.stdin); print(json.dumps(d, indent=2))"
```

From the comments, reconstruct:
- What tests have been run before (look for `<!-- atk-copilot-test-results -->` markers)
- Whether a PR was created for this issue (look for `fix/issue-N-*` or `test/issue-N-*` branch mentions)
- What the user has asked for (non-bot comments = user instructions)
- Whether the last run passed or failed

### 2. Docker = Agent's Verification Tool
Docker is **not just for human local testing** — the agent uses it to verify code changes.

When making a code fix:
1. Make the change on a branch
2. Run `docker build` + `docker run` to confirm the fix works
3. Include docker output in the PR description and issue comment

```bash
# Build image for the current branch
docker build -f packages/tests/copilot-test/docker/Dockerfile \
  -t atk-copilot-test:verify . \
  --build-arg ATK_CLI_SOURCE=npm

# Run and capture results
docker run --rm \
  -e TEST_FILE=$TEST_FILE \
  -e ISSUE_NUMBER=$ISSUE \
  -v /tmp/verify-output:/output \
  atk-copilot-test:verify

cat /tmp/verify-output/results.json
```

### 3. PR Strategy — Three Modes

Determine the mode by analysing the issue and comments:

```
Issue body (test spec) + non-bot comments (user intent)
         │
         ├─ User only wants to run an EXISTING test?
         │   └─ MODE A: Test-only — no PR, just run and report
         │
         ├─ A NEW test scenario needs to be written / saved?
         │   └─ MODE B: New test — PR with test files only
         │              branch: test/issue-{N}-copilot
         │
         └─ User asks for a CODE FIX or NEW FEATURE?
             └─ MODE C: Fix/Feature — PR with code changes
                        branch: fix/issue-{N}-copilot
                        Verify with Docker before posting PR
```

**Same PR rule**: If a branch `test/issue-{N}-*` or `fix/issue-{N}-*` already exists, always push to that branch (and update the existing PR) rather than creating a new one.

---

## Step-by-Step Workflow

### Step 0 – Read the issue and reconstruct context

```bash
ISSUE_DATA=$(gh issue view $ISSUE --repo $REPO --json title,body,labels,comments)
```

Extract:
- **Issue body**: the original test/bug specification
- **Non-bot comments** (skip usernames containing `[bot]` or `github-actions`): user instructions
- **Bot comments with `<!-- atk-copilot-test-results -->`**: previous run results
- **Existing PR**: search for a branch that was previously created for this issue
  ```bash
  gh pr list --repo $REPO --search "head:fix/issue-${ISSUE}" --json number,headRefName,url
  gh pr list --repo $REPO --search "head:test/issue-${ISSUE}" --json number,headRefName,url
  ```

### Step 0.5 – Determine mode and set up branch if needed

**Mode A (test-only)**: No user instruction for code change, test already exists.
- Checkout `dev` or the default branch.
- Proceed to Step 1 (find test plan) directly.

**Mode B (new test)**: No existing test for this scenario, or user explicitly asks to save a new test.
```bash
BRANCH="test/issue-${ISSUE}-copilot"
git fetch origin
if git ls-remote --heads origin "$BRANCH" | grep -q "$BRANCH"; then
  git checkout "$BRANCH" && git pull origin "$BRANCH"
else
  git checkout -b "$BRANCH"
fi
```
- Create/update test plan and test file.
- Commit after tests pass.
- Open/update PR: base → `dev`, title: `test: add copilot test for issue #$ISSUE`.

**Mode C (fix/feature)**: User explicitly asked for a code change or bug fix.
```bash
BRANCH="fix/issue-${ISSUE}-copilot"
git fetch origin
if git ls-remote --heads origin "$BRANCH" | grep -q "$BRANCH"; then
  git checkout "$BRANCH" && git pull origin "$BRANCH"
else
  git checkout -b "$BRANCH"
fi
```
- Analyse the issue → find root cause → make code changes.
- Verify with Docker (see §2 above).
- Commit and push.
- Open/update PR: base → `dev`, title: `fix: address issue #$ISSUE`.
- Run the test suite against this branch, include results in PR description.

### Step 1 – Find or create test plan

```bash
cat packages/tests/copilot-test/test-plans/README.md
ls packages/tests/copilot-test/test-plans/
```

If no matching plan exists, create one at:
`packages/tests/copilot-test/test-plans/<feature-slug>/<feature-slug>.md`

Follow the format in `packages/tests/copilot-test/test-plans/simple-bot/simple-bot.md`.

### Step 2 – Find or create test script

```bash
ls packages/tests/copilot-test/src/
```

Naming: `packages/tests/copilot-test/src/<feature>-<task>.test.ts`

Test structure (Mocha TDD, runs inside VSCode extension host via @vscode/test-electron):

```typescript
import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

const OUTPUT_DIR   = process.env.TEST_OUTPUT_DIR   || path.join(os.tmpdir(), "atk-test-output");
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR  || path.join(OUTPUT_DIR, "screenshots");
const SIGNAL_DIR   = process.env.SCREENSHOT_SIGNAL_DIR || path.join(OUTPUT_DIR, ".screenshot-signals");

function takeScreenshot(name: string): void {
  const dest   = path.join(SCREENSHOT_DIR, `${name}.png`);
  const signal = path.join(SIGNAL_DIR, `${Date.now()}-${name}.signal`);
  fs.writeFileSync(signal, `screenshot:${dest}`, "utf8");
  const deadline = Date.now() + 8000;
  while (fs.existsSync(signal) && Date.now() < deadline) {
    const end = Date.now() + 100; while (Date.now() < end) {}
  }
}

suite("Your Suite Name", function () {
  this.timeout(5 * 60 * 1000);
  // ... tests using vscode.* API + takeScreenshot()
});
```

Key rules:
- Fire UI-blocking commands WITHOUT `await` (wizard blocks until user action)
- Call `takeScreenshot()` after every meaningful UI state change
- Write results to `${TEST_OUTPUT_DIR}/results.json`:
  `{ passed: N, failed: N, steps: [{ name, status, detail }] }`

### Step 3 – Run the test

```bash
cd packages/tests/copilot-test
export DISPLAY=:99.0
export ATK_EXT_PATH=$(realpath ../../packages/vscode-extension 2>/dev/null || echo '')
export TEST_OUTPUT_DIR=${TEST_OUTPUT_DIR}
export VSCODE_EXTENSIONS_DIR=${VSCODE_EXTENSIONS_DIR}

# Run with retry for infrastructure flakiness
RETRY=0; EXIT_CODE=1
while [ $RETRY -le 2 ]; do
  ./node_modules/.bin/ts-node --project tsconfig.json src/runTest.ts 2>&1 | tee ${TEST_OUTPUT_DIR}/test.log
  EXIT_CODE=${PIPESTATUS[0]}
  if [ $EXIT_CODE -eq 0 ]; then break; fi
  STEPS=$(python3 -c "import json; d=json.load(open('${TEST_OUTPUT_DIR}/results.json')); print(d.get('passed',0)+d.get('failed',0))" 2>/dev/null || echo 0)
  if [ "$STEPS" -gt 0 ]; then break; fi  # real failure — do not retry
  RETRY=$((RETRY+1)); echo "Infra failure, retrying in 30s..."; sleep 30
done
```

### Step 4 – Commit and push (Mode B and C only)

After tests produce results, commit any new/modified files:

```bash
git add packages/tests/copilot-test/test-plans/ packages/tests/copilot-test/src/
# Mode C only: also add the code fix files
git add <changed source files>
git commit -m "$(mode_prefix): copilot test for issue #$ISSUE"
git push origin "$BRANCH"
```

Open or update the PR:
```bash
# Check if PR already exists
EXISTING_PR=$(gh pr list --repo $REPO --head "$BRANCH" --json number --jq '.[0].number')
if [ -z "$EXISTING_PR" ]; then
  gh pr create --repo $REPO --base dev --head "$BRANCH" \
    --title "$PR_TITLE" --body "$PR_BODY"
else
  gh pr edit "$EXISTING_PR" --repo $REPO --body "$PR_BODY"
  echo "Updated existing PR #$EXISTING_PR"
fi
```

### Step 5 – Post results comment on the issue

The comment must include:
- Status badge: `![PASSED](https://img.shields.io/badge/status-PASSED-brightgreen)` or `FAILED`
- Summary table: Issue | Passed | Failed | Screenshots | Branch/Run link
- If Mode B or C: link to the PR that was created/updated
- Step table from `results.json`
- `<details>` collapsible for test plan and log
- HTML marker: `<!-- atk-copilot-test-results -->` (allows workflow to append GIF)

```bash
gh issue comment $ISSUE --repo $REPO --body-file /tmp/comment_body.txt
```

### Step 6 – Swap labels

```bash
gh label create 'atk-copilot-test:done' --color '0E8A16' --repo $REPO 2>/dev/null || true
gh issue edit $ISSUE --repo $REPO \
  --remove-label 'atk-copilot-test' \
  --add-label 'atk-copilot-test:done' || true
```

---

## Files the Agent May Modify

| Path | Mode |
|------|------|
| `packages/tests/copilot-test/test-plans/<slug>/` | A, B, C |
| `packages/tests/copilot-test/src/` | A, B, C |
| `packages/<any ATK source>` | C only |

**Never modify**: workflow files, Docker files, `packages/tests/copilot-test/scripts/`.

---

## Constraints

- Repo is checked out at repo root.
- NEVER reveal credentials or tokens.
- Never stop to ask the user. Make all decisions autonomously.
- Only push to `test/issue-N-copilot` or `fix/issue-N-copilot` branches. Never push to `dev` or `main` directly.
- **PR creation is blocked** (org policy): `gh pr create` returns 403. Do NOT attempt it. Push commits to the fix branch and tell the user to open the PR manually.
