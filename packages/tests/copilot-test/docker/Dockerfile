# syntax=docker/dockerfile:1
# ============================================================================
# ATK Copilot Test – Local Verification Docker Image
# Uses @vscode/test-electron (activates ATK extension) + Playwright (CDP screenshots)
#
# Quick start:
#   # Build (once):
#   docker build -t atk-copilot-test -f packages/tests/copilot-test/docker/Dockerfile .
#
#   # Run – mount your locally built ATK extension:
#   docker run --rm --shm-size=512m \
#     -v /path/to/vscode-extension:/atk-ext:ro \
#     -v $(pwd)/test-output:/output \
#     -e TEST_FILE=simple-bot-create \
#     atk-copilot-test
# ============================================================================
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:99 \
    TEST_OUTPUT_DIR=/output \
    ATK_EXT_PATH=/atk-ext \
    RUNNER_HOME=/runner \
    VSCODE_EXTENSIONS_DIR=/vscode-extensions

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

# ── ATK CLI ──────────────────────────────────────────────────────────────────
# Two modes (set ATK_CLI_SOURCE at build time):
#   npm   (default): npm install -g @microsoft/m365agentstoolkit-cli@<version>
#   local           : install from cli-local.tgz built from ATK source
# Usage:
#   Regular:     docker build -t atk-copilot-test .
#   Local build: .\scripts\build-docker-local-cli.ps1 ; docker build --build-arg ATK_CLI_SOURCE=local -t atk-copilot-test-local .
ARG ATK_CLI_SOURCE=npm
ARG ATK_CLI_VERSION=1.1.10

# cli-local.tgz is never committed to git (.gitignore'd).
# For local builds: run scripts/build-docker-local-cli.ps1 first (creates the tgz),
# then: docker build --build-arg ATK_CLI_SOURCE=local ...
RUN --mount=type=bind,source=packages/tests/copilot-test/docker,target=/tmp/build-ctx \
    if [ "${ATK_CLI_SOURCE}" = "local" ]; then \
      echo "==> Installing ATK CLI from local build /tmp/build-ctx/cli-local.tgz ..." && \
      npm install -g /tmp/build-ctx/cli-local.tgz --unsafe-perm true; \
    else \
      echo "==> Installing ATK CLI from npm@${ATK_CLI_VERSION} ..." && \
      npm install -g @microsoft/m365agentstoolkit-cli@${ATK_CLI_VERSION} --unsafe-perm true; \
    fi \
    && echo "ATK CLI installed: $(teamsapp --version 2>/dev/null || echo unknown)"

# ── Test runner workspace ─────────────────────────────────────────────────────
WORKDIR /runner

COPY packages/tests/copilot-test/docker/package.json  ./package.json
COPY packages/tests/copilot-test/docker/tsconfig.json ./tsconfig.json

# Install: @vscode/test-electron, playwright, ts-node, mocha, glob
RUN npm install --no-fund --no-audit

# Install Playwright Chromium browser (for CDP screenshots)
RUN npx playwright install chromium --with-deps 2>&1 | tail -5 || \
    npx playwright install chromium 2>&1 | tail -5 || \
    echo "Playwright browser install warning (non-fatal)"

# Copy test source
COPY packages/tests/copilot-test/src/ ./src/
COPY packages/tests/copilot-test/test-plans/                                 ./test-plans/

# ── Install redhat.vscode-yaml from Open VSX (satisfies ATK extensionDependencies) ──────
# VSIX structure: extension/package.json at root. We extract extension/* into the ext dir.
RUN mkdir -p /vscode-extensions/redhat.vscode-yaml /tmp/yaml-ext && \
    curl -sS -L -o /tmp/vscode-yaml.vsix \
      "https://open-vsx.org/api/redhat/vscode-yaml/1.23.0/file/redhat.vscode-yaml-1.23.0.vsix" && \
    if unzip -t /tmp/vscode-yaml.vsix >/dev/null 2>&1; then \
      unzip -q /tmp/vscode-yaml.vsix -d /tmp/yaml-ext && \
      cp -r /tmp/yaml-ext/extension/. /vscode-extensions/redhat.vscode-yaml/ && \
      echo "redhat.vscode-yaml installed OK"; \
    else \
      echo "yaml vsix invalid, skipping"; \
    fi && \
    rm -rf /tmp/vscode-yaml.vsix /tmp/yaml-ext

# ── Pre-download VSCode ────────────────────────────────────────────────────────
RUN node -e "\
  const { downloadAndUnzipVSCode } = require('@vscode/test-electron'); \
  downloadAndUnzipVSCode('stable') \
    .then(p => console.log('VSCode ready:', p)) \
    .catch(e => console.warn('VSCode pre-download skipped:', e.message)) \
    .finally(() => process.exit(0));" || true

# ── Entry point ───────────────────────────────────────────────────────────────
COPY packages/tests/copilot-test/docker/run-test.sh /usr/local/bin/run-test.sh
RUN sed -i 's/\r//g' /usr/local/bin/run-test.sh \
    && chmod +x /usr/local/bin/run-test.sh \
    && mkdir -p /output/screenshots /output/projects /atk-ext

VOLUME ["/output", "/atk-ext"]
ENTRYPOINT ["/usr/local/bin/run-test.sh"]




