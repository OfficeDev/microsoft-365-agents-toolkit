# ============================================================================
# ATK Copilot Test – Local Verification Docker Image
# ============================================================================
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:99 \
    NODE_VERSION=22 \
    VSCODE_STORAGE=/tmp/vscode-test-storage \
    TEST_OUTPUT_DIR=/output

# ── System packages ───────────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    xvfb x11-utils \
    libgtk-3-0 libxss1 libnss3 libasound2 libxtst6 libxrandr2 libgbm1 libxkbcommon0 libdrm2 \
    curl wget git jq python3 python3-pip ca-certificates unzip \
    && rm -rf /var/lib/apt/lists/*

# ── Node.js 22 ────────────────────────────────────────────────────────────────
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y nodejs \
    && npm install -g pnpm@9 \
    && node --version && npm --version

# ── ATK CLI (for CLI-mode tests) ─────────────────────────────────────────────
RUN npm install -g @microsoft/m365agentstoolkit-cli@latest --unsafe-perm true 2>/dev/null \
    || echo "ATK CLI install skipped"

# ── App directory ─────────────────────────────────────────────────────────────
WORKDIR /app

# Copy only what's needed for the tests package (layer-cache friendly)
COPY packages/tests/package.json      packages/tests/package.json
COPY packages/tests/pnpm-lock.yaml    packages/tests/pnpm-lock.yaml
COPY packages/tests/tsconfig.json     packages/tests/tsconfig.json
COPY packages/tests/.mocharc.js      packages/tests/.mocharc.js

RUN cd packages/tests \
    && pnpm install --frozen-lockfile --ignore-scripts \
    || npm install --legacy-peer-deps

# Copy test source (our copilot-driven tests + existing utils)
COPY packages/tests/src packages/tests/src

# Copy test plans (used by Copilot to understand what to test)
COPY test-plans test-plans

# ── Pre-download VSCode binary ────────────────────────────────────────────────
# This caches the ~100 MB VSCode download so docker run is fast
RUN cd packages/tests && node -e "\
  const { ExTester } = require('vscode-extension-tester'); \
  const t = new ExTester(process.env.VSCODE_STORAGE || '/tmp/vscode-test-storage'); \
  t.downloadCode('latest') \
    .then(() => { console.log('VSCode ready'); process.exit(0); }) \
    .catch(e => { console.warn('VSCode download skipped:', e.message); process.exit(0); });" \
    || true

# ── Entry point ───────────────────────────────────────────────────────────────
COPY docker/copilot-test/run-test.sh /usr/local/bin/run-test.sh
RUN sed -i 's/\r//g' /usr/local/bin/run-test.sh
RUN chmod +x /usr/local/bin/run-test.sh \
    && mkdir -p /output/screenshots /output/projects

VOLUME ["/output"]
ENTRYPOINT ["/usr/local/bin/run-test.sh"]