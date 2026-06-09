# ATK Copilot Test Docker image
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:99 \
    TEST_OUTPUT_DIR=/output \
    ATK_EXT_PATH=/atk-ext \
    RUNNER_HOME=/runner \
    VSCODE_EXTENSIONS_DIR=/vscode-extensions

RUN apt-get update && apt-get install -y --no-install-recommends \
    xvfb x11-utils \
    libgtk-3-0 libxss1 libnss3 libxtst6 libxrandr2 libgbm1 libxkbcommon0 libdrm2 \
    libsecret-1-0 libglib2.0-0 libpulse0 \
    curl wget git jq python3 ca-certificates unzip \
    && { apt-get install -y libasound2 2>/dev/null || apt-get install -y libasound2t64 2>/dev/null || true; } \
    && rm -rf /var/lib/apt/lists/*

RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y nodejs \
    && node --version && npm --version

ARG ATK_CLI_SOURCE=npm
ARG ATK_CLI_VERSION=1.1.10

RUN if [ "${ATK_CLI_SOURCE}" = "local" ]; then \
      echo "==> Installing ATK CLI from local build /tmp/build-ctx/cli-local.tgz ..." && \
      npm install -g /tmp/build-ctx/cli-local.tgz --unsafe-perm true; \
    else \
      echo "==> Installing ATK CLI from npm@${ATK_CLI_VERSION} ..." && \
      npm install -g @microsoft/m365agentstoolkit-cli@${ATK_CLI_VERSION} --unsafe-perm true; \
    fi \
    && echo "ATK CLI installed: $(teamsapp --version 2>/dev/null || echo unknown)"

WORKDIR /runner

COPY packages/tests/copilot-test/docker/package.json ./package.json
COPY packages/tests/copilot-test/docker/tsconfig.json ./tsconfig.json

RUN npm install --no-fund --no-audit

RUN npx playwright install chromium --with-deps 2>&1 | tail -5 || \
    npx playwright install chromium 2>&1 | tail -5 || \
    echo "Playwright browser install warning (non-fatal)"

COPY packages/tests/copilot-test/src/ ./src/
COPY packages/tests/copilot-test/test-plans/ ./test-plans/


RUN node -e " \
  const { downloadAndUnzipVSCode } = require('@vscode/test-electron'); \
  downloadAndUnzipVSCode('stable') \
    .then(p => console.log('VSCode ready:', p)) \
    .catch(e => console.warn('VSCode pre-download skipped:', e.message)) \
    .finally(() => process.exit(0));" || true

COPY packages/tests/copilot-test/docker/run-test.sh /usr/local/bin/run-test.sh
RUN sed -i 's/\r//g' /usr/local/bin/run-test.sh \
    && chmod +x /usr/local/bin/run-test.sh \
    && mkdir -p /output/screenshots /output/projects /atk-ext

# Bake pre-compiled ATK extension into the image.
# This avoids the Docker-on-Windows bind-mount cold-start (170 KB/s via virtio-fs),
# which caused 600+ second unresponsive periods while Node loaded 83 MB of extension JS.
# Users can still override by bind-mounting a custom build to /atk-ext at runtime.
COPY packages/vscode-extension/out/ /atk-ext-baked/out/
COPY packages/vscode-extension/package.json /atk-ext-baked/package.json
COPY packages/vscode-extension/package.nls.json /atk-ext-baked/package.nls.json
COPY packages/vscode-extension/media/ /atk-ext-baked/media/
COPY packages/vscode-extension/syntaxes/ /atk-ext-baked/syntaxes/

VOLUME ["/output", "/atk-ext"]
ENTRYPOINT ["/usr/local/bin/run-test.sh"]

