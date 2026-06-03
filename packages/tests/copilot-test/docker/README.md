# ATK Copilot Test – Docker Setup

Self-contained Docker image that runs `@vscode/test-electron` + Playwright CDP UI tests against the
Microsoft 365 Agents Toolkit (ATK) VSCode extension — the same tests that run in the
GitHub Actions pipeline, runnable locally without any extra setup.

## Quick start

```bash
# From the repository root:

# 1. Build the image (first time or after code changes)
docker build -t atk-copilot-test -f packages/tests/copilot-test/docker/Dockerfile .

# 2. Run the default test
docker run --rm \
  -v $(pwd)/packages/tests/copilot-test/test-output:/output \
  --shm-size=512m \
  atk-copilot-test

# 3. Check results
ls packages/tests/copilot-test/test-output/screenshots/     # PNG screenshots for every step
cat packages/tests/copilot-test/test-output/results.json    # { passed, failed, steps: [...] }
cat packages/tests/copilot-test/test-output/test.log        # full console output
```

## Run a specific test

```bash
docker run --rm \
  -v $(pwd)/packages/tests/copilot-test/test-output:/output \
  -e TEST_FILE=sample-app-a11y \
  --shm-size=512m \
  atk-copilot-test
```

`TEST_FILE` is the stem of a `.test.ts` file under `packages/tests/copilot-test/src/`.

## Run with docker compose (easier)

```bash
cd packages/tests/copilot-test/docker

# Default test
docker compose run --rm test

# Specific test
TEST_FILE=sample-app-a11y docker compose run --rm test
```

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `TEST_FILE` | `simple-bot-create` | Test file stem (without `.test.ts`) |
| `VSCODE_VERSION` | `stable` | VSCode version |
| `TEST_OUTPUT_DIR` | `/output` | Where results are written (mount a volume here) |

## Using a locally built ATK extension

By default the image uses the ATK extension baked in at build time.
To use a custom local build, mount it to `/atk-ext`:

```bash
docker run --rm \
  -v $(pwd)/packages/tests/copilot-test/test-output:/output \
  -v $(pwd)/packages/vscode-extension:/atk-ext:ro \
  --shm-size=512m \
  atk-copilot-test
```

The container auto-detects a valid mounted extension (`/atk-ext/out/src/extension.js`) and prefers it over the baked copy.

## Output files

| File | Description |
|---|---|
| `screenshots/<id>.png` | Screenshot after each named step |
| `results.json` | `{ passed, failed, steps: [...] }` |
| `test.log` | Full mocha + runner console output |
| `projects/<app-name>/` | Scaffolded project (if test creates one) |