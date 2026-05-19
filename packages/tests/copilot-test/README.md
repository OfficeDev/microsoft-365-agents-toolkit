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
    simple-bot-create.test.ts  TC-001: create Simple Bot via wizard
  docker/                   Docker image for local + CI runs
    Dockerfile
    docker-compose.yml
    run-test.sh             Container entry point
    package.json            Standalone deps (no pnpm workspace protocol)
    tsconfig.json
    cli-local.tgz           Placeholder for local ATK CLI build
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

## Quick start (Docker)

```bash
# From repo root — build once
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

## Adding a new test

1. Copy `test-plans/template.md` → `test-plans/<feature-slug>/<feature-slug>.md` and fill in the test plan.
2. Create `src/<feature-slug>-<task>.test.ts` following `src/simple-bot-create.test.ts` as a reference.
3. Run via Docker or with `node scripts/local-test.js <feature-slug>-<task>`.

## Pipeline

The CI pipeline (`atk-copilot-test-runner.yml`) triggers when a GitHub issue is labelled `atk-copilot-test`.
The Copilot CLI agent reads the issue, selects the test plan, runs the test, and posts results back.