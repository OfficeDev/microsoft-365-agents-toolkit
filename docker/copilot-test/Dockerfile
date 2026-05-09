# ============================================================================
# ATK Copilot Test – Local Verification Docker Image
# ============================================================================
# Provides a self-contained environment to run vscode-extension-tester UI tests
# against the Microsoft 365 Agents Toolkit (ATK) VSCode extension locally,
# matching what the GitHub Actions pipeline does in CI.
# ============================================================================

FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:99 \
    NODE_VERSION=22 \
    VSCODE_STORAGE=/tmp/vscode-test-storage \
    TEST_OUTPUT_DIR=/output

# ── System packages ───────────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    # Virtual display
    xvfb \
    x11-utils \
    # VSCode / Electron runtime deps
    libgtk-3-0 \
    libxss1 \
    libnss3 \
    libasound2 \
    libxtst6 \
    libxrandr2 \
    libgbm1 \
    libxkbcommon0 \
    libdrm2 \
    libgles2 \
    # General tools
    curl \
    wget \
    git \
    jq \
    python3 \
    python3-pip \
    ca-certificates \
    unzip \
    && rm -rf /var/lib/apt/lists/*

# ── Node.js ───────────────────────────────────────────────────────────────────
RUN curl -fsSL https://deb.nodesource.com/setup_${NODE_VERSION}.x | bash - \
    && apt-get install -y nodejs \
    && npm install -g pnpm@9 \
    && node --version && npm --version

# ── ATK CLI (for CLI-mode tests) ─────────────────────────────────────────────
RUN npm install -g @microsoft/m365agentstoolkit-cli@latest --unsafe-perm true 2>/dev/null || \
    echo "ATK CLI install skipped (network or package not available)"

# ── App directory ─────────────────────────────────────────────────────────────
WORKDIR /app

# Copy the repo's tests package
COPY packages/tests/package.json packages/tests/pnpm-lock.yaml ./packages/tests/
COPY packages/tests/tsconfig.json ./packages/tests/ 2>/dev/null || true
COPY packages/tests/tsconfig.test.json ./packages/tests/ 2>/dev/null || true
COPY packages/tests/.mocharc.json ./packages/tests/ 2>/dev/null || true

# Install test dependencies
RUN cd packages/tests && \
    pnpm install --frozen-lockfile --ignore-scripts 2>/dev/null || \
    npm install --legacy-peer-deps

# Copy full test source
COPY packages/tests/src ./packages/tests/src
COPY packages/vscode-extension ./packages/vscode-extension 2>/dev/null || true

# Copy test plans
COPY test-plans ./test-plans

# ── Pre-download VSCode ───────────────────────────────────────────────────────
RUN cd packages/tests && \
    node -e " \
      const { ExTester } = require('vscode-extension-tester'); \
      const t = new ExTester('${VSCODE_STORAGE}'); \
      t.downloadCode('stable') \
        .then(() => { console.log('VSCode downloaded'); process.exit(0); }) \
        .catch(e => { console.warn('VSCode download failed:', e.message); process.exit(0); }); \
    " 2>&1 || true

# ── Docker entry point ────────────────────────────────────────────────────────
COPY docker/copilot-test/run-test.sh /usr/local/bin/run-test.sh
RUN chmod +x /usr/local/bin/run-test.sh

RUN mkdir -p /output/screenshots /output/projects

VOLUME ["/output"]

ENTRYPOINT ["/usr/local/bin/run-test.sh"]