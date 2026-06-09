#!/usr/bin/env bash
# ============================================================================
# ATK Copilot Test – Container Entry Point (@vscode/test-electron)
# ============================================================================
set -euo pipefail

TEST_FILE="${TEST_FILE:-simple-bot-create}"
ATK_EXT_PATH="${ATK_EXT_PATH:-/atk-ext-baked}"
TEST_OUTPUT_DIR="${TEST_OUTPUT_DIR:-/output}"
# Use local bundled templates to avoid GitHub download rate-limits in CI/Docker
export TEMPLATE_VERSION="${TEMPLATE_VERSION:-local}"
RUNNER_HOME="/runner"

echo "======================================================"
echo "  ATK Copilot Test Runner (@vscode/test-electron)"
echo "  Test    : ${TEST_FILE}"
echo "  Ext     : ${ATK_EXT_PATH}"
echo "  Output  : ${TEST_OUTPUT_DIR}"
echo "======================================================"

mkdir -p "${TEST_OUTPUT_DIR}/screenshots" "${TEST_OUTPUT_DIR}/projects"

# Clean up artifacts from previous runs — Docker creates root-owned files that
# Windows cannot delete. We run as root inside the container so this always works.
rm -f "${TEST_OUTPUT_DIR}/screenshots/"*.png 2>/dev/null || true
rm -f "${TEST_OUTPUT_DIR}/test.log" 2>/dev/null || true
echo "[init] Cleared previous screenshots and logs"

# Pre-create ATK global state file to prevent proper-lockfile stale threshold crash.
# Without this, the extension's first write to ~/.fx/state.json acquires a lock during
# gallery operations, which can exceed the 10s stale threshold and crash the test.
mkdir -p "${HOME}/.fx"
[ -f "${HOME}/.fx/state.json" ] || echo '{}' > "${HOME}/.fx/state.json"
echo "[init] Pre-created ~/.fx/state.json"

# Auto-detect extension: prefer bind-mounted /atk-ext if it has compiled output,
# otherwise fall back to the baked version built into the image.
if [ -f "/atk-ext/out/src/extension.js" ]; then
  ATK_EXT_PATH="/atk-ext"
  echo "[ext] Using bind-mounted extension at /atk-ext"
elif [ -f "/atk-ext-baked/out/src/extension.js" ]; then
  ATK_EXT_PATH="/atk-ext-baked"
  echo "[ext] Using baked extension at /atk-ext-baked (fast native-FS path)"
fi
export ATK_EXT_PATH
[ -d "${ATK_EXT_PATH}" ] || { echo "ERROR: ATK extension not found. Provide -v /path/to/vscode-extension:/atk-ext:ro or rebuild image with compiled out/."; exit 1; }

# -- Build extension if compiled output is missing --------------------------------
# The extension main entry is out/src/extension.js (compiled TypeScript).
# It is NOT committed to git. If missing, copy the source to a writable dir and build.
if [ ! -f "${ATK_EXT_PATH}/out/src/extension.js" ]; then
  echo "[build] Compiled extension not found at ${ATK_EXT_PATH}/out/src/extension.js"
  echo "[build] Building extension from source (this takes ~2 min)..."
  BUILD_DIR="/tmp/atk-ext-build"
  rm -rf "${BUILD_DIR}"
  cp -r "${ATK_EXT_PATH}" "${BUILD_DIR}"
  cd "${BUILD_DIR}"
  # Extension uses workspace:* deps — pnpm is required.
  # Since we are outside the monorepo, replace workspace:* with "*" for standalone build.
  python3 -c "
import json, re
with open(\"package.json\") as f: p = json.load(f)
for section in [\"dependencies\", \"devDependencies\"]:
    for k in p.get(section, {}):
        if p[section][k] == \"workspace:*\":
            p[section][k] = \"*\"
with open(\"package.json\", \"w\") as f: json.dump(p, f, indent=2)
print(\"Replaced workspace:* deps\")
"
  npm install --legacy-peer-deps 2>&1 | tail -5
  npm run build 2>&1 | tail -20
  if [ ! -f "${BUILD_DIR}/out/src/extension.js" ]; then
    echo "[build] ERROR: Build failed -- out/src/extension.js still missing"
    exit 1
  fi
  export ATK_EXT_PATH="${BUILD_DIR}"
  echo "[build] Extension built at ${ATK_EXT_PATH}"
  cd "${RUNNER_HOME}"
else
  echo "[build] Extension already compiled at ${ATK_EXT_PATH}/out/src/extension.js"
fi

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
    "src/runTest.ts" 2>&1 | tee /tmp/run-test.log || true

TEST_EXIT=${PIPESTATUS[0]}
# Copy log to mounted output (may fail if permissions are wrong on Windows bind-mount)
cp /tmp/run-test.log "${TEST_OUTPUT_DIR}/test.log" 2>/dev/null || true

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