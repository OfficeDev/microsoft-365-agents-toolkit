# packages/tests/copilot-test

Copilot-driven automated UI tests for the ATK (Microsoft 365 Agents Toolkit) VSCode extension.

Tests run headlessly using `@vscode/test-electron` + Playwright CDP hybrid architecture.
The same Docker image runs both locally and in CI.

## Directory layout

```
copilot-test/
  src/                      Test source files
    runTest.ts              Hybrid orchestrator (@vscode/test-electron + Playwright CDP)
    suite/index.ts          Mocha suite entry point
    simple-bot-create.test.ts   Example test: create Simple Bot via wizard
    sample-app-a11y.test.ts     Example test: sample app accessibility checks
  docker/                   Docker image for local + CI runs
    Dockerfile
    docker-compose.yml
    run-test.sh             Container entry point
    package.json            Standalone deps (no pnpm workspace protocol)
    tsconfig.json
    README.md               Docker quick-start guide
  skills/                   Copilot CLI skill documents (prompts for CI agents)
    label-brain.md          Label agent: reads issue, fixes code, creates test plan
    test-generator.md       Generator agent: generates .test.ts from test plan, runs & fixes
  test-plans/               Test plans (human-authored, read by Copilot CLI agents)
    README.md
    template.md             Template for new test plans
    sample-app-a11y/
      sample-app-a11y.md    Accessibility test plan (TC-001 through TC-006)
    simple-bot/
      simple-bot.md         Simple Bot wizard test plan
  scripts/
    post_issue_comment.py   Post test results to GitHub issue
    build-docker-local-cli.ps1  Build Docker image with local ATK CLI
```

## Quick start (Docker)

The ATK extension is baked into the image at build time — no bind-mount needed for normal runs.

```bash
# From repo root — build once (includes ATK extension compiled output)
docker build -t atk-copilot-test -f packages/tests/copilot-test/docker/Dockerfile .

# Run the default test
docker run --rm --shm-size=512m \
  -v "$(pwd)/packages/tests/copilot-test/test-output:/output" \
  atk-copilot-test

# Run a specific test file
docker run --rm --shm-size=512m \
  -e TEST_FILE=sample-app-a11y \
  -v "$(pwd)/packages/tests/copilot-test/test-output:/output" \
  atk-copilot-test

# Check results
cat packages/tests/copilot-test/test-output/results.json
ls  packages/tests/copilot-test/test-output/screenshots/
```

Or use Docker Compose (handles volumes automatically):

```bash
cd packages/tests/copilot-test/docker
TEST_FILE=sample-app-a11y docker compose run --rm test
```

## Adding a new test

1. Copy `test-plans/template.md` → `test-plans/<feature-slug>/<feature-slug>.md` and fill in the test plan.
2. Create `src/<feature-slug>.test.ts` following the existing tests as a reference.
3. Build the Docker image and run: `TEST_FILE=<feature-slug> docker compose run --rm test`

## CI Pipeline (3-layer)

Triggered when a GitHub issue is labelled `atk-copilot-test`:

```
[Label Brain]        atk-copilot-test-label.yml
  ↓ reads issue, optionally fixes product code, creates test plan
[Generator]          atk-copilot-test-generator.yml
  ↓ generates .test.ts from test plan, runs test, analyzes & retries on failure
[Runner]             atk-copilot-test-runner.yml
    pure executor: builds VSIX → runs headless VSCode → uploads artifacts
```

Skill documents used by the Copilot CLI agents live in `skills/`.