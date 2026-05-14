# ATK Copilot Test – Docker Setup

Self-contained Docker image that runs `vscode-extension-tester` UI tests against the
Microsoft 365 Agents Toolkit (ATK) VSCode extension — the same tests that run in the
GitHub Actions pipeline, runnable locally without any extra setup.

## Quick start

```bash
# From the repository root:

# 1. Build the image (first time or after code changes)
docker build -t atk-copilot-test -f packages/tests/copilot-test/docker/Dockerfile .

# 2. Run the default test (TC-001: create Teams Bot template)
docker run --rm \
  -v $(pwd)/test-output:/output \
  --shm-size=512m \
  atk-copilot-test

# 3. Check results
ls test-output/screenshots/     # PNG screenshots for every step
cat test-output/results.json    # { passed, failed, steps: [...] }
cat test-output/test.log        # full console output
```

## Run a specific test

```bash
docker run --rm \
  -v $(pwd)/test-output:/output \
  -e TEST_FILE=teams-bot-name-validation \
  --shm-size=512m \
  atk-copilot-test
```

`TEST_FILE` is the stem of a file under `packages/tests/src/`.

## Run with docker compose (easier)

```bash
cd packages/tests/copilot-test/docker

# Default test
docker compose run --rm test

# Specific test
TEST_FILE=teams-bot-name-validation docker compose run --rm test
```

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `TEST_FILE` | `simple-bot-create` | Test file stem |
| `VSIX_PATH` | _(auto-detect)_ | Path to ATK `.vsix` inside container |
| `VSCODE_VERSION` | `stable` | VSCode version for ExTester |
| `TEST_OUTPUT_DIR` | `/output` | Where results are written (mount a volume here) |

## Passing a custom VSIX

```bash
docker run --rm \
  -v $(pwd)/test-output:/output \
  -v $(pwd)/my-extension.vsix:/ext.vsix \
  -e VSIX_PATH=/ext.vsix \
  --shm-size=512m \
  atk-copilot-test
```

## Output files

| File | Description |
|---|---|
| `screenshots/<id>.png` | Screenshot after each named step |
| `results.json` | `{ passed, failed, steps: [...] }` |
| `test.log` | Full mocha + ExTester console output |
| `projects/<app-name>/` | Scaffolded project (if test creates one) |