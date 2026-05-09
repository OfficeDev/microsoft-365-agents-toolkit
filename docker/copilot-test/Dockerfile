# ============================================================================
# ATK Copilot Test – Local Verification Docker Image
# Uses @vscode/test-electron (activates ATK extension) + Playwright (CDP screenshots)
#
# Quick start:
#   # Build (once):
#   docker build -t atk-copilot-test -f docker/copilot-test/Dockerfile .
#
#   # Run – mount your locally built ATK extension:
#   docker run --rm --shm-size=512m \
#     -v /path/to/vscode-extension:/atk-ext:ro \
#     -v $(pwd)/test-output:/output \
#     -e TEST_FILE=teams-bot-create-template \
#     atk-copilot-test
#
# Results appear in ./test-output/
#   screenshots/  – PNG files (real Playwright captures via CDP)
#   results.json  – { passed, failed, steps }
#   test.log      – full console output
# ============================================================================
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:99 \
    TEST_OUTPUT_DIR=/output \
    ATK_EXT_PATH=/atk-ext \
    RUNNER_HOME=/runner \
    # Playwright uses its own Chromium; skip full browser download
    PLAYWRIGHT_BROWSERS_PATH=/playwright-browsers

# ── System packages ───────────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    xvfb x11-utils \
    libgtk-3-0 libxss1 libnss3 libxtst6 libxrandr2 libgbm1 libxkbcommon0 libdrm2 \
    libsecret-1-0 libglib2.0-0 libpulse0 \
    curl wget git jq python3 ca-certificates unzip \
    && { apt-get install -y libasound2 2>/dev/null || apt-get install -y libasound2t64 2>/dev/null || true; } \
    && rm -rf /var/lib/apt/lists/*

# ── Node.js 22 ────────────────────────────────────────────────────────────────
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y nodejs \
    && node --version && npm --version

# ── ATK CLI (optional) ────────────────────────────────────────────────────────
RUN npm install -g @microsoft/m365agentstoolkit-cli@latest --unsafe-perm true 2>/dev/null \
    || echo "ATK CLI install skipped (non-fatal)"

# ── Test runner workspace ─────────────────────────────────────────────────────
WORKDIR /runner

COPY docker/copilot-test/package.json  ./package.json
COPY docker/copilot-test/tsconfig.json ./tsconfig.json

# Install: @vscode/test-electron, playwright, ts-node, mocha, glob
RUN npm install --no-fund --no-audit

# Install Playwright Chromium browser (for CDP screenshots)
RUN npx playwright install chromium --with-deps 2>&1 || \
    npx playwright install chromium 2>&1 || \
    echo "Playwright browser install warning (non-fatal)"

# Copy test source
COPY packages/tests/src/ui-test/copilot-driven/ ./src/ui-test/copilot-driven/
COPY test-plans/                                 ./test-plans/

# ── Pre-install redhat.vscode-yaml (satisfies ATK extensionDependencies) ──────
# We install it to a known extensions dir so VSCode finds it at test time
RUN mkdir -p /vscode-extensions && \
    node -e "\
      const https = require('https'); \
      const fs = require('fs'); \
      const url = 'https://marketplace.visualstudio.com/_apis/public/gallery/publishers/redhat/vsextensions/vscode-yaml/latest/vspackage'; \
      const file = fs.createWriteStream('/tmp/vscode-yaml.vsix'); \
      https.get(url, {headers: {'Accept': 'application/octet-stream'}}, res => res.pipe(file).on('finish', () => process.exit(0))); \
    " && \
    cd /vscode-extensions && \
    unzip -q /tmp/vscode-yaml.vsix -d redhat.vscode-yaml && \
    rm /tmp/vscode-yaml.vsix && \
    echo "redhat.vscode-yaml installed" || \
    echo "yaml extension pre-install skipped (will use --install-extension at runtime)"

ENV VSCODE_EXTENSIONS_DIR=/vscode-extensions

# ── Pre-download VSCode ────────────────────────────────────────────────────────
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
