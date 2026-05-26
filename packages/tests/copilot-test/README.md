# packages/tests/copilot-test

Copilot-driven automated UI tests for the ATK (Microsoft 365 Agents Toolkit) VSCode extension.

Tests run headlessly using `@vscode/test-electron` + Playwright CDP hybrid architecture.

## Pipeline Architecture

The CI pipeline is split into three layers:

```
GitHub Issue labeled `atk-copilot-test`
        |
        v
[Layer 1] atk-copilot-test-label.yml
  Label trigger -- routes to the coverage agent
        |
        v
[Layer 2] atk-copilot-test-coverage.yml
  Copilot CLI agent -- reads the issue, finds/creates a test plan and
  test script, runs the runner, analyzes results, iterates (max 3x),
  posts a structured results comment, swaps label
        |
        v
[Layer 3] atk-copilot-test-runner.yml
  Pure test executor -- given a script name, runs it headlessly,
  uploads results.json + screenshots artifact
        |
        v
  Results posted to GitHub issue
```

### Layer 1 - Label Trigger (atk-copilot-test-label.yml)

Listens for the `atk-copilot-test` label on issues and routes to the coverage
workflow. No business logic.

### Layer 2 - Coverage Agent (atk-copilot-test-coverage.yml)

Runs `copilot --yolo` as the orchestrating brain. The agent:

1. Reads the issue to understand what to test
2. Finds or creates a **test plan** in `test-plans/<feature>/`
3. Finds or creates a **test script** in `src/<feature>.test.ts`
4. Commits + pushes, then triggers the runner via `gh workflow run`
5. Downloads `results.json` and analyzes:
   - All steps pass -> post success comment, swap label to `atk-copilot-test:done`
   - **Test script issue** (wrong selector, wrong API call, timing) -> fix the
     script, re-run (max 3 iterations)
   - **Product bug** (ATK does not behave as the test plan describes) -> write
     `bugs/<feature>-bug.md`, post failure comment

### Layer 3 - Test Runner (atk-copilot-test-runner.yml)

Pure executor. Given a `script` input (test file basename, e.g.
`simple-bot-create`), it:

- Installs all dependencies and system libs
- Starts Xvfb for headless rendering
- Runs `ts-node src/runTest.ts` with `TEST_FILE=<script>`
- Parses `results.json`, generates a GIF, uploads artifacts
- Outputs `passed`/`failed` counts

The runner has no knowledge of issues, labels, or Copilot CLI.

## User Experience

1. **Developer labels** a GitHub issue with `atk-copilot-test`
2. **Layer 1** wakes up and routes to the coverage agent
3. **Coverage agent** reads the issue and finds or creates a matching test plan
4. Agent writes or updates the test script, commits, and pushes
5. Agent triggers the **runner** with the test file name and current branch ref
6. Runner executes headlessly and uploads `results.json` + screenshots
7. Agent downloads and analyzes the results:
   - Iterates up to 3 times if the test script has bugs
   - Files a bug report if the issue is a product defect
8. Structured results comment posted to the issue with pass/fail table and GIF link
9. Label swapped: `atk-copilot-test` -> `atk-copilot-test:done`

## Directory layout

```
copilot-test/
  src/                      Test source files
    runTest.ts              Hybrid orchestrator (@vscode/test-electron + Playwright CDP)
    suite/index.ts          Mocha suite entry point
    simple-bot-create.test.ts  TC-001: create Simple Bot via wizard
  bugs/                     Product bug reports filed by the coverage agent
  docker/                   Docker image for local + CI runs
    Dockerfile
    docker-compose.yml
    run-test.sh             Container entry point
    README.md               Docker quick-start guide
  test-plans/               Test plans (read by Copilot CLI agent)
    README.md
    template.md             Template for new test plans
    simple-bot/
      simple-bot.md         TC-001 test plan (Simple Bot wizard flow)
  scripts/
    post_issue_comment.py   Post test results to GitHub issue
    build-docker-local-cli.ps1  Build Docker image with local ATK CLI
    local-test.js           Run tests locally without Docker
```

## Adding a new test plan

1. Copy `test-plans/template.md` -> `test-plans/<feature-slug>/<feature-slug>.md`
   and fill in the test plan.
2. Create `src/<feature-slug>.test.ts` following `src/simple-bot-create.test.ts`
   as a reference.

Or let Copilot write the test for you -- just label any issue `atk-copilot-test`!

## Writing a test script

Test files use Mocha TDD inside VSCode's extension host via `@vscode/test-electron`.
Key rules:

- **Do NOT** use `vscode-extension-tester` (wrong library)
- Call `takeScreenshot(name)` after every meaningful UI state change
- Fire UI-blocking VSCode commands **without** `await` (wizards block until user
  action -- see `simple-bot-create.test.ts` for the pattern)
- Write `results.json` to `$TEST_OUTPUT_DIR`:
  `{ passed, failed, steps: [{name, status, detail}] }`

See `src/simple-bot-create.test.ts` for a complete working example.

## Running locally (Docker)

```bash
# From repo root -- build once
docker build -t atk-copilot-test -f packages/tests/copilot-test/docker/Dockerfile .

# Run TC-001 (Simple Bot creation)
docker run --rm --shm-size=512m \
  -v "$HOME/.vscode/extensions/teamsdevapp.ms-teams-vscode-extension-6.8.0:/atk-ext:ro" \
  -v "$(pwd)/test-output:/output" \
  atk-copilot-test

# Check results
cat test-output/results.json
ls  test-output/screenshots/
```