# ============================================================================
# VscUse ATK Profile - Microsoft 365 Agent Toolkit
# ============================================================================
# Extends vscuse-base with .NET, Azure Functions, M365 CLI, Dev Tunnel, Edge
# ============================================================================

ARG BASE_IMAGE=ghcr.io/microsoft/vscuse-base:latest
FROM ${BASE_IMAGE}

# Install .NET 8.0 SDK
RUN curl -fsSL https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb -o packages-microsoft-prod.deb \
    && dpkg -i packages-microsoft-prod.deb \
    && rm packages-microsoft-prod.deb \
    && apt-get update \
    && apt-get install -y dotnet-sdk-8.0 \
    && rm -rf /var/lib/apt/lists/* \
    && dotnet --version \
    && dotnet --list-sdks

# Install Azure Functions Core Tools v4.4.0
RUN npm install -g azure-functions-core-tools@4.4.0 --unsafe-perm true

# Install Microsoft 365 Agents Toolkit CLI
RUN npm install -g @microsoft/m365agentstoolkit-cli@alpha --unsafe-perm true

# Install Dev Tunnel CLI (architecture-aware)
RUN ARCH=$(dpkg --print-architecture) && \
    if [ "$ARCH" = "amd64" ]; then \
        DEVTUNNEL_ARCH="x64"; \
    elif [ "$ARCH" = "arm64" ]; then \
        DEVTUNNEL_ARCH="arm64"; \
    else \
        DEVTUNNEL_ARCH="x64"; \
    fi && \
    mkdir -p /tmp/devtunnel && \
    curl -sL "https://aka.ms/TunnelsCliDownload/linux-${DEVTUNNEL_ARCH}" -o /tmp/devtunnel/devtunnel && \
    chmod +x /tmp/devtunnel/devtunnel && \
    mv /tmp/devtunnel/devtunnel /usr/local/bin/devtunnel && \
    rm -rf /tmp/devtunnel && \
    devtunnel --version || echo "DevTunnel installed"

# Create M365 template devTools directory
RUN mkdir -p /opt/m365-template/devTools/func \
    && if [ -d "/usr/lib/node_modules/azure-functions-core-tools/bin" ]; then \
        cp -r /usr/lib/node_modules/azure-functions-core-tools/bin/* /opt/m365-template/devTools/func/; \
    fi \
    && chmod +x /opt/m365-template/devTools/func/* 2>/dev/null || true \
    && chown -R vscode:vscode /opt/m365-template

# Install Microsoft Edge (AMD64 only)
RUN --mount=type=cache,target=/var/cache/apt,sharing=locked \
    --mount=type=cache,target=/var/lib/apt,sharing=locked \
    ARCH=$(dpkg --print-architecture) && \
    if [ "$ARCH" = "amd64" ]; then \
        echo "Installing Microsoft Edge for AMD64..." && \
        curl -fsSL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor -o /usr/share/keyrings/microsoft-edge-keyring.gpg && \
        echo "deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft-edge-keyring.gpg] https://packages.microsoft.com/repos/edge stable main" > /etc/apt/sources.list.d/microsoft-edge.list && \
        apt-get update && \
        apt-get install -y microsoft-edge-stable && \
        rm -rf /var/lib/apt/lists/* && \
        microsoft-edge-stable --version; \
    else \
        echo "Microsoft Edge is not officially supported on ARM64, skipping..."; \
    fi

# Create Edge wrapper script with Docker-optimized flags
RUN cat > /usr/local/bin/edge-wrapper << 'EOF'
#!/bin/bash
export DISPLAY=${DISPLAY:-:99}
mkdir -p /tmp/.X11-unix
chmod 1777 /tmp/.X11-unix

if [ -f "/usr/bin/microsoft-edge-stable" ]; then
    BROWSER="/usr/bin/microsoft-edge-stable"
else
    echo "Error: Microsoft Edge not found"
    exit 1
fi

exec $BROWSER \
  --no-sandbox \
  --disable-dev-shm-usage \
  --disable-gpu \
  --disable-gpu-sandbox \
  --disable-software-rasterizer \
  --disable-background-timer-throttling \
  --disable-backgrounding-occluded-windows \
  --disable-renderer-backgrounding \
  --disable-features=TranslateUI,VizDisplayCompositor \
  --disable-hang-monitor \
  --disable-client-side-phishing-detection \
  --disable-popup-blocking \
  --disable-default-apps \
  --disable-extensions-file-access-check \
  --disable-component-update \
  --disable-background-networking \
  --disable-sync \
  --disable-translate \
  --disable-web-security \
  --disable-ipc-flooding-protection \
  --disable-prompt-on-repost \
  --disable-extensions \
  --no-crash-upload \
  --no-default-browser-check \
  --no-first-run \
  --no-pings \
  --disable-first-run-ui \
  --noerrdialogs \
  --disable-restore-session-state \
  --remote-debugging-port=0 \
  --password-store=basic \
  --disable-password-generation \
  --disable-save-password-bubble \
  --disable-password-manager \
  --disable-password-manager-reauthentication \
  --user-data-dir=/tmp/edge-user-data-$(whoami) \
  "$@"
EOF

# Setup Edge wrapper and symlinks
RUN chmod +x /usr/local/bin/edge-wrapper \
    && sed -i 's/\r$//' /usr/local/bin/edge-wrapper \
    && ln -sf /usr/local/bin/edge-wrapper /usr/local/bin/microsoft-edge \
    && ln -sf /usr/local/bin/edge-wrapper /usr/local/bin/edge

# Configure Edge policies
RUN mkdir -p /etc/opt/microsoft/edge/policies/managed && \
    echo '{ \
        "PasswordManagerEnabled": false, \
        "AutofillAddressEnabled": false, \
        "AutofillCreditCardEnabled": false, \
        "BuiltInDnsClientEnabled": false, \
        "MetricsReportingEnabled": false, \
        "SearchSuggestEnabled": false, \
        "AlternateErrorPagesEnabled": false, \
        "SafeBrowsingProtectionLevel": 0, \
        "SpellcheckEnabled": false, \
        "TranslateEnabled": false, \
        "PasswordProtectionWarningTrigger": 0, \
        "SyncDisabled": true \
    }' > /etc/opt/microsoft/edge/policies/managed/disable_features.json && \
    chmod 644 /etc/opt/microsoft/edge/policies/managed/disable_features.json

# Install ATK-specific VS Code extensions
RUN mkdir -p /app/build-extensions
COPY build-extensions/ /app/build-extensions/
RUN /usr/local/bin/install-build-extensions.sh && \
    mkdir -p /home/vscode/.vscode/extensions /home/vscode/.vscode-insiders/extensions && \
    cp -r /root/.vscode/extensions/* /home/vscode/.vscode/extensions/ 2>/dev/null || true && \
    cp -r /root/.vscode-insiders/extensions/* /home/vscode/.vscode-insiders/extensions/ 2>/dev/null || true && \
    chown -R vscode:vscode /home/vscode/.vscode /home/vscode/.vscode-insiders

LABEL org.opencontainers.image.description="VscUse ATK Profile - Microsoft 365 Agent Toolkit"
LABEL org.opencontainers.image.profile="atk"
