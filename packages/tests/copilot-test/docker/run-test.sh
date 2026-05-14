#!/usr/bin/env bash
# ============================================================================
# ATK Copilot Test – Container Entry Point (@vscode/test-electron)
# ============================================================================
set -euo pipefail

TEST_FILE="${TEST_FILE:-simple-bot-create}"
ATK_EXT_PATH="${ATK_EXT_PATH:-/atk-ext}"
TEST_OUTPUT_DIR="${TEST_OUTPUT_DIR:-/output}"
RUNNER_HOME="/runner"

echo "======================================================"
echo "  ATK Copilot Test Runner (@vscode/test-electron)"
echo "  Test    : ${TEST_FILE}"
echo "  Ext     : ${ATK_EXT_PATH}"
echo "  Output  : ${TEST_OUTPUT_DIR}"
echo "======================================================"

mkdir -p "${TEST_OUTPUT_DIR}/screenshots" "${TEST_OUTPUT_DIR}/projects"

[ -d "${ATK_EXT_PATH}" ] || { echo "ERROR: ATK extension not found at ${ATK_EXT_PATH}"; echo "Mount: -v /path/to/vscode-extension:/atk-ext:ro"; exit 1; }

# ── Virtual display ───────────────────────────────────────────────────────────
echo "[1/3] Starting Xvfb..."
Xvfb :99 -ac -screen 0 1920x1080x24 &
XVFB_PID=$!
sleep 2
export DISPLAY=:99.0
cleanup() { kill "${XVFB_PID}" 2>/dev/null || true; }
trap cleanup EXIT

# ── Verify test file exists ───────────────────────────────────────────────────
TEST_SPEC="${RUNNER_HOME}/src/${TEST_FILE}.test.ts"
RUNNER_TS="${RUNNER_HOME}/src/runTest.ts"
[ -f "${TEST_SPEC}" ]  || { echo "ERROR: Test spec not found: ${TEST_SPEC}"; exit 1; }
[ -f "${RUNNER_TS}" ]  || { echo "ERROR: runTest.ts not found: ${RUNNER_TS}"; exit 1; }

# ── Run @vscode/test-electron ─────────────────────────────────────────────────
echo "[2/3] Running @vscode/test-electron..."
cd "${RUNNER_HOME}"

DISPLAY=:99.0 \
ATK_EXT_PATH="${ATK_EXT_PATH}" \
TEST_OUTPUT_DIR="${TEST_OUTPUT_DIR}" \
TEST_FILE="${TEST_FILE}" \
  ./node_modules/.bin/ts-node \
    --project tsconfig.json \
    "src/runTest.ts" 2>&1 | tee "${TEST_OUTPUT_DIR}/test.log"

TEST_EXIT=${PIPESTATUS[0]}

# ── Summary ───────────────────────────────────────────────────────────────────
echo "[3/3] Summarising..."
SCREENSHOTS=$(ls "${TEST_OUTPUT_DIR}/screenshots/"*.png 2>/dev/null | wc -l || echo 0)
RESULTS_FILE="${TEST_OUTPUT_DIR}/results.json"
PASSED=0; FAILED=0
if [ -f "${RESULTS_FILE}" ]; then
  PASSED=$(python3 -c "import json; d=json.load(open('${RESULTS_FILE}')); print(d.get('passed',0))" 2>/dev/null || echo 0)
  FAILED=$(python3 -c "import json; d=json.load(open('${RESULTS_FILE}')); print(d.get('failed',0))" 2>/dev/null || echo 0)
else
  PASSED=$(grep -oP '\d+(?= passing)' "${TEST_OUTPUT_DIR}/test.log" | tail -1 || echo 0)
  FAILED=$(grep -oP '\d+(?= failing)' "${TEST_OUTPUT_DIR}/test.log" | tail -1 || echo 0)
fi

echo ""
echo "======================================================"
echo "  Results  |  Passed: ${PASSED}  Failed: ${FAILED}  Screenshots: ${SCREENSHOTS}"
echo "  Output   :  ${TEST_OUTPUT_DIR}/"
echo "======================================================"

if [ "${TEST_EXIT}" -ne 0 ] || [ "${FAILED:-0}" -gt 0 ]; then
  echo "RESULT: FAILED"; exit 1
else
  echo "RESULT: PASSED"; exit 0
fi