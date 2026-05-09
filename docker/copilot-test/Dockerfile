# ============================================================================
# ATK Copilot Test – Local Verification Docker Image
# Uses @vscode/test-electron (official Microsoft VSCode test runner).
# Tests run INSIDE the VSCode extension host, not via external browser automation.
#
# Quick start:
#   # Build (once):
#   docker build -t atk-copilot-test -f docker/copilot-test/Dockerfile .
#
#   # Run – mount your locally built ATK extension and output dir:
#   docker run --rm --shm-size=512m \
#     -v /c/Users/quke/source/atk/microsoft-365-agents-toolkit/packages/vscode-extension:/atk-ext:ro \
#     -v $(pwd)/test-output:/output \
#     -e TEST_FILE=teams-bot-create-template \
#     atk-copilot-test
# ============================================================================
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:99 \
    TEST_OUTPUT_DIR=/output \
    ATK_EXT_PATH=/atk-ext \
    RUNNER_HOME=/runner

# ── System packages ───────────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    xvfb x11-utils \
    libgtk-3-0 libxss1 libnss3 libxtst6 libxrandr2 libgbm1 libxkbcommon0 libdrm2 \
    libsecret-1-0 libglib2.0-0 \
    curl wget git jq python3 ca-certificates unzip scrot \
    && { apt-get install -y libasound2 2>/dev/null || apt-get install -y libasound2t64 2>/dev/null || true; } \
    && rm -rf /var/lib/apt/lists/*

# ── Node.js 22 ────────────────────────────────────────────────────────────────
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y nodejs \
    && node --version && npm --version

# ── ATK CLI (optional – for CLI-based scaffold tests) ────────────────────────
RUN npm install -g @microsoft/m365agentstoolkit-cli@latest --unsafe-perm true 2>/dev/null \
    || echo "ATK CLI install skipped (non-fatal)"

# ── Test runner workspace ─────────────────────────────────────────────────────
# We use a standalone package.json so workspace:* deps don't break the install.
WORKDIR /runner

COPY docker/copilot-test/package.json  ./package.json
COPY docker/copilot-test/tsconfig.json ./tsconfig.json

# Install: @vscode/test-electron, ts-node, mocha, glob
RUN npm install --no-fund --no-audit

# Copy our test source into the runner workspace
COPY packages/tests/src/ui-test/copilot-driven/ ./src/ui-test/copilot-driven/
COPY test-plans/                                 ./test-plans/

# ── Pre-download VSCode (speeds up docker run) ────────────────────────────────
RUN node -e "\
  const { downloadAndUnzipVSCode } = require('@vscode/test-electron'); \
  downloadAndUnzipVSCode('stable') \
    .then(p => console.log('VSCode ready:', p)) \
    .catch(e => console.warn('VSCode pre-download skipped:', e.message)) \
    .finally(() => process.exit(0));" || true

# ── Entry point ───────────────────────────────────────────────────────────────
COPY docker/copilot-test/run-test.sh /usr/local/bin/run-test.sh
RUN sed -i 's/\r//g' /usr/local/bin/run-test.sh \
    && chmod +x /usr/local/bin/run-test.sh \
    && mkdir -p /output/screenshots /output/projects /atk-ext

VOLUME ["/output", "/atk-ext"]
ENTRYPOINT ["/usr/local/bin/run-test.sh"]