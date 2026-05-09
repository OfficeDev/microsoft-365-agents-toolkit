#!/usr/bin/env bash
# ============================================================================
# ATK Copilot Test – Container Entry Point
# ============================================================================
# Environment variables:
#   TEST_FILE       Test filename stem (e.g. "teams-bot-create-template")
#                   Resolves to: packages/tests/src/ui-test/copilot-driven/<TEST_FILE>.test.ts
#   ATK_EXT_PATH    (optional) Path to a mounted ATK extension directory.
#                   Mount your locally built extension:
#                     -v /path/to/vscode-extension:/atk-ext:ro
#                   then set -e ATK_EXT_PATH=/atk-ext
#   TEST_OUTPUT_DIR (optional) Where to write results (default: /output)
#   VSCODE_VERSION  (optional) VSCode version (default: latest)
# ============================================================================
set -euo pipefail

TEST_FILE="${TEST_FILE:-teams-bot-create-template}"
ATK_EXT_PATH="${ATK_EXT_PATH:-}"
VSCODE_VERSION="${VSCODE_VERSION:-latest}"
TEST_OUTPUT_DIR="${TEST_OUTPUT_DIR:-/output}"
VSCODE_STORAGE="${VSCODE_STORAGE:-/tmp/vscode-test-storage}"
REPO_ROOT="/app"

echo "======================================================"
echo "  ATK Copilot Test Runner"
echo "  Test:    ${TEST_FILE}"
echo "  Output:  ${TEST_OUTPUT_DIR}"
echo "  ExtPath: ${ATK_EXT_PATH:-<none – test only>}"
echo "======================================================"

mkdir -p "${TEST_OUTPUT_DIR}/screenshots" "${TEST_OUTPUT_DIR}/projects"

# ── Start virtual display ─────────────────────────────────────────────────────
echo "[1/4] Starting Xvfb virtual display..."
Xvfb :99 -ac -screen 0 1920x1080x24 &
XVFB_PID=$!
sleep 2
export DISPLAY=:99.0

cleanup() {
  kill "${XVFB_PID}" 2>/dev/null || true
}
trap cleanup EXIT

# ── Resolve test spec ─────────────────────────────────────────────────────────
echo "[2/4] Resolving test spec..."
SPEC="src/ui-test/copilot-driven/${TEST_FILE}.test.ts"
FULL_SPEC="${REPO_ROOT}/packages/tests/${SPEC}"
if [ ! -f "${FULL_SPEC}" ]; then
  echo "ERROR: Test file not found: ${FULL_SPEC}"
  exit 1
fi
echo "      Spec: ${SPEC}"

# ── Run test ──────────────────────────────────────────────────────────────────
echo "[3/4] Running test suite..."
cd "${REPO_ROOT}/packages/tests"

# ExTester v8 strategy:
#   - coverage=true skips installVsix (no extension to install from marketplace)
#   - coverage=true also sets EXTENSION_DEV_PATH = process.cwd() in runTests()
#   - We temporarily override process.cwd() to return ATK_EXT_PATH so VSCode
#     launches with --extensionDevelopmentPath pointing to the real extension.
#   - If ATK_EXT_PATH is empty, EXTENSION_DEV_PATH remains unset and VSCode
#     starts without ATK loaded (useful for API/smoke tests).

ATK_EXT_PATH_ESCAPED="${ATK_EXT_PATH//\'/\'}"

node -e "
'use strict';
const path = require('path');
const fs = require('fs');
const { ExTester, ReleaseQuality } = require('vscode-extension-tester');

const OUTPUT    = '${TEST_OUTPUT_DIR}';
const STORAGE   = '${VSCODE_STORAGE}';
const SPEC      = '${SPEC}';
const EXT_PATH  = '${ATK_EXT_PATH_ESCAPED}';

process.env.TEST_OUTPUT_DIR = OUTPUT;

const tester = new ExTester(STORAGE, ReleaseQuality.Stable, undefined, true /* coverage */);

(async () => {
  // Download VSCode + ChromeDriver (cached on subsequent runs)
  await tester.setupRequirements(
    { vscodeVersion: '${VSCODE_VERSION}', installDependencies: false },
    false, false
  );

  // If extension path provided, monkey-patch process.cwd() so ExTester picks it up
  if (EXT_PATH) {
    const realCwd = process.cwd.bind(process);
    process.cwd = () => EXT_PATH;
    const code = await tester.runTests(SPEC, { vscodeVersion: '${VSCODE_VERSION}' });
    process.cwd = realCwd;
    process.exit(code);
  } else {
    // No extension – run test without --extensionDevelopmentPath
    const code = await tester.runTests(SPEC, { vscodeVersion: '${VSCODE_VERSION}' });
    process.exit(code);
  }
})().catch(e => { console.error('Error:', e.message); process.exit(1); });
" 2>&1 | tee "${TEST_OUTPUT_DIR}/test.log"
TEST_EXIT=${PIPESTATUS[0]}

# ── Summarise results ─────────────────────────────────────────────────────────
echo "[4/4] Summarising results..."
SCREENSHOTS=$(ls "${TEST_OUTPUT_DIR}/screenshots/"*.png 2>/dev/null | wc -l || echo 0)

RESULTS_FILE="${TEST_OUTPUT_DIR}/results.json"
if [ ! -f "${RESULTS_FILE}" ]; then
  PASSED=$(grep -oP '\d+(?= passing)' "${TEST_OUTPUT_DIR}/test.log" | tail -1 || echo 0)
  FAILED=$(grep -oP '\d+(?= failing)' "${TEST_OUTPUT_DIR}/test.log" | tail -1 || echo 0)
  python3 -c "
import json
with open('${RESULTS_FILE}', 'w') as f:
    json.dump({'passed': int('${PASSED}' or 0), 'failed': int('${FAILED}' or 0), 'steps': []}, f, indent=2)
"
else
  PASSED=$(python3 -c "import json; d=json.load(open('${RESULTS_FILE}')); print(d.get('passed',0))")
  FAILED=$(python3 -c "import json; d=json.load(open('${RESULTS_FILE}')); print(d.get('failed',0))")
fi

echo ""
echo "======================================================"
echo "  Results"
echo "  Passed:      ${PASSED}"
echo "  Failed:      ${FAILED}"
echo "  Screenshots: ${SCREENSHOTS} captured"
echo "  Output:      ${TEST_OUTPUT_DIR}/"
echo "======================================================"

if [ "${TEST_EXIT}" -ne 0 ] || [ "${FAILED}" -gt 0 ]; then
  echo "RESULT: FAILED"; exit 1
else
  echo "RESULT: PASSED"; exit 0
fi