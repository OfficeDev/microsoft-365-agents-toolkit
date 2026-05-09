#!/usr/bin/env bash
# ============================================================================
# ATK Copilot Test – Container Entry Point
# ============================================================================
# Usage (via docker run or docker-compose):
#
#   docker run --rm \
#     -v $(pwd)/test-output:/output \
#     -e TEST_FILE=teams-bot-create-template \
#     atk-copilot-test
#
# Environment variables:
#   TEST_FILE          (required) filename stem of the test to run,
#                      e.g. "teams-bot-create-template"
#                      resolves to: packages/tests/src/ui-test/copilot-driven/<TEST_FILE>.test.ts
#   VSIX_PATH          (optional) path to ATK .vsix file; auto-detected if not set
#   VSCODE_VERSION     (optional) VSCode version, default: stable
#   TEST_OUTPUT_DIR    (optional) where to write results; default: /output
# ============================================================================

set -euo pipefail

TEST_FILE="${TEST_FILE:-teams-bot-create-template}"
VSIX_PATH="${VSIX_PATH:-}"
VSCODE_VERSION="${VSCODE_VERSION:-stable}"
TEST_OUTPUT_DIR="${TEST_OUTPUT_DIR:-/output}"
VSCODE_STORAGE="${VSCODE_STORAGE:-/tmp/vscode-test-storage}"
REPO_ROOT="/app"

echo "======================================================"
echo "  ATK Copilot Test Runner"
echo "  Test:    ${TEST_FILE}"
echo "  Output:  ${TEST_OUTPUT_DIR}"
echo "======================================================"

mkdir -p "${TEST_OUTPUT_DIR}/screenshots" "${TEST_OUTPUT_DIR}/projects"

# ── Start virtual display ─────────────────────────────────────────────────────
echo "[1/5] Starting Xvfb virtual display..."
Xvfb :99 -ac -screen 0 1920x1080x24 &
XVFB_PID=$!
sleep 2
export DISPLAY=:99.0
echo "      DISPLAY=${DISPLAY}"

cleanup() {
  echo ""
  echo "[cleanup] Stopping Xvfb..."
  kill "${XVFB_PID}" 2>/dev/null || true
}
trap cleanup EXIT

# ── Resolve ATK VSIX ─────────────────────────────────────────────────────────
echo "[2/5] Resolving ATK extension..."
if [ -z "${VSIX_PATH}" ]; then
  VSIX_PATH=$(find "${REPO_ROOT}" -name "*.vsix" 2>/dev/null | head -1 || true)
fi

if [ -n "${VSIX_PATH}" ] && [ -f "${VSIX_PATH}" ]; then
  echo "      Using VSIX: ${VSIX_PATH}"
  EXT_ARG="--extensionPath ${VSIX_PATH}"
else
  echo "      No VSIX found – using extensionDevelopmentPath (built extension)"
  EXT_DIR="${REPO_ROOT}/packages/vscode-extension"
  if [ -d "${EXT_DIR}" ]; then
    EXT_ARG="--extensionDevelopmentPath ${EXT_DIR}"
  else
    echo "      WARNING: No extension found. Test may fail without ATK installed."
    EXT_ARG=""
  fi
fi

# ── Resolve test spec ─────────────────────────────────────────────────────────
echo "[3/5] Resolving test spec..."
SPEC="src/ui-test/copilot-driven/${TEST_FILE}.test.ts"
FULL_SPEC="${REPO_ROOT}/packages/tests/${SPEC}"
if [ ! -f "${FULL_SPEC}" ]; then
  echo "ERROR: Test file not found: ${FULL_SPEC}"
  exit 1
fi
echo "      Spec: ${SPEC}"

# ── Run the test via vscode-extension-tester ─────────────────────────────────
echo "[4/5] Running test..."
cd "${REPO_ROOT}/packages/tests"

DISPLAY=:99.0 node -e "
  const { ExTester } = require('vscode-extension-tester');
  process.env.TEST_OUTPUT_DIR = '${TEST_OUTPUT_DIR}';

  const extArgs = '${EXT_ARG}';
  const t = new ExTester('${VSCODE_STORAGE}', undefined, '${TEST_OUTPUT_DIR}/screenshots');

  const runOpts = {
    config: '.mocharc.json',
    logLevel: 'info',
  };

  t.runTests('${SPEC}', runOpts)
    .then(code => {
      console.log('ExTester exit code:', code);
      process.exit(code);
    })
    .catch(e => {
      console.error('ExTester error:', e.message);
      process.exit(1);
    });
" 2>&1 | tee "${TEST_OUTPUT_DIR}/test.log"
TEST_EXIT=${PIPESTATUS[0]}

# ── Summarise results ─────────────────────────────────────────────────────────
echo "[5/5] Summarising results..."
RESULTS_FILE="${TEST_OUTPUT_DIR}/results.json"
SCREENSHOTS=$(ls "${TEST_OUTPUT_DIR}/screenshots/"*.png 2>/dev/null | wc -l || echo 0)

if [ -f "${RESULTS_FILE}" ]; then
  PASSED=$(python3 -c "import json; d=json.load(open('${RESULTS_FILE}')); print(d.get('passed',0))")
  FAILED=$(python3 -c "import json; d=json.load(open('${RESULTS_FILE}')); print(d.get('failed',0))")
else
  # Parse from mocha log
  PASSED=$(grep -oP '\d+(?= passing)' "${TEST_OUTPUT_DIR}/test.log" | tail -1 || echo 0)
  FAILED=$(grep -oP '\d+(?= failing)' "${TEST_OUTPUT_DIR}/test.log" | tail -1 || echo 0)
  # Write minimal results.json
  python3 -c "
import json
with open('${RESULTS_FILE}', 'w') as f:
    json.dump({'passed': int('${PASSED}' or 0), 'failed': int('${FAILED}' or 0), 'steps': []}, f, indent=2)
"
fi

echo ""
echo "======================================================"
echo "  Results"
echo "  Passed:      ${PASSED}"
echo "  Failed:      ${FAILED}"
echo "  Screenshots: ${SCREENSHOTS} captured"
echo "  Output dir:  ${TEST_OUTPUT_DIR}/"
echo "======================================================"

if [ "${TEST_EXIT}" -ne 0 ] || [ "${FAILED}" -gt 0 ]; then
  echo "RESULT: ❌ FAILED"
  exit 1
else
  echo "RESULT: ✅ PASSED"
  exit 0
fi