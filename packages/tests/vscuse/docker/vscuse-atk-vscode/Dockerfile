# ============================================================================
# VscUse VS Code Docker Container
# ============================================================================
# Base: vscuse-base (Ubuntu 22.04 LTS with Python 3.12, noVNC, X11)
# Features: VS Code (stable + insiders), Chrome, Edge, Node.js, .NET, Azure tools, ATK
# ============================================================================

# Use vscuse-base as the base image
FROM ghcr.io/microsoft/vscuse-base:latest AS base

# Avoid prompts from apt
ENV DEBIAN_FRONTEND=noninteractive

# ============================================================================
# Stage 1: Node.js and Development Tools
# ============================================================================
FROM base AS dev-tools

# Install Node.js 20 LTS from NodeSource
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && rm -rf /var/lib/apt/lists/*

# ============================================================================
# .NET SDK Installation
# ============================================================================
# Install .NET 8.0 SDK (LTS) from Microsoft
RUN curl -fsSL https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb -o packages-microsoft-prod.deb \
    && dpkg -i packages-microsoft-prod.deb \
    && rm packages-microsoft-prod.deb \
    && apt-get update \
    && apt-get install -y dotnet-sdk-8.0 \
    && rm -rf /var/lib/apt/lists/* \
    && dotnet --version \
    && dotnet --list-sdks

# Install Azure Functions Core Tools v4.4.0 (latest stable version)
RUN npm install -g azure-functions-core-tools@4.4.0 --unsafe-perm true

# Install Microsoft 365 Agents Toolkit CLI
RUN npm install -g @microsoft/m365agentstoolkit-cli@alpha --unsafe-perm true

# Install Dev Tunnel CLI - manual installation for better compatibility
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

# Create a template devTools directory structure for M365 projects
RUN mkdir -p /opt/m365-template/devTools/func \
    && if [ -d "/usr/lib/node_modules/azure-functions-core-tools/bin" ]; then \
        cp -r /usr/lib/node_modules/azure-functions-core-tools/bin/* /opt/m365-template/devTools/func/; \
    fi \
    && chmod +x /opt/m365-template/devTools/func/* 2>/dev/null || true

# ============================================================================
# Stage 2: VS Code Installation
# ============================================================================
FROM dev-tools AS vscode-install

# Install VS Code stable and insiders
RUN wget -qO- https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor > packages.microsoft.gpg \
    && install -o root -g root -m 644 packages.microsoft.gpg /etc/apt/trusted.gpg.d/ \
    && echo "deb [arch=amd64,arm64,armhf signed-by=/etc/apt/trusted.gpg.d/packages.microsoft.gpg] https://packages.microsoft.com/repos/code stable main" > /etc/apt/sources.list.d/vscode.list \
    && apt-get update \
    && echo "Installing VS Code stable and Insider..." \
    && apt-get install -y code code-insiders \
    && rm -rf /var/lib/apt/lists/* \
    && echo "VS Code versions installed:" \
    && code --version || echo "VS Code stable: command available" \
    && code-insiders --version || echo "VS Code Insider: command available"

# ============================================================================
# Stage 3: Browser Installation (Architecture-Aware)
# ============================================================================
FROM vscode-install AS browser-install

# Install Chrome/Chromium based on architecture
RUN ARCH=$(dpkg --print-architecture) && \
    echo "Installing browser for architecture: $ARCH" && \
    if [ "$ARCH" = "amd64" ]; then \
        # For AMD64, install Google Chrome (most stable and performant)
        echo "Installing Google Chrome for AMD64..." && \
        wget -qO- https://dl.google.com/linux/linux_signing_key.pub | \
            gpg --dearmor -o /usr/share/keyrings/google-chrome-keyring.gpg && \
        echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome-keyring.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list && \
        apt-get update && \
        apt-get install -y google-chrome-stable && \
        rm -rf /var/lib/apt/lists/* && \
        google-chrome-stable --version; \
    else \
        # For ARM64, download and install Chromium directly (avoid snap issues)
        echo "Installing Chromium for ARM64..." && \
        apt-get update && \
        apt-get install -y wget unzip libgtk-3-0 libgbm1 libnss3 libxss1 libasound2 && \
        # Download pre-built Chromium for ARM64
        CHROMIUM_URL="https://download-chromium.appspot.com/dl/Linux_x64" && \
        if [ "$ARCH" = "arm64" ]; then \
            CHROMIUM_URL="https://download-chromium.appspot.com/dl/Linux_ARM"; \
        fi && \
        wget -O /tmp/chromium.zip "$CHROMIUM_URL" && \
        unzip /tmp/chromium.zip -d /opt/ && \
        mv /opt/chrome-linux* /opt/chromium && \
        ln -s /opt/chromium/chrome /usr/local/bin/chromium && \
        rm -f /tmp/chromium.zip && \
        rm -rf /var/lib/apt/lists/* && \
        /usr/local/bin/chromium --version || echo "Chromium binary installed"; \
    fi && \
    # Create a universal wrapper script
    cat > /usr/local/bin/chrome-wrapper << 'EOF'
#!/bin/bash
# Set display if not already set
export DISPLAY=${DISPLAY:-:99}

# Ensure temp directories exist
mkdir -p /tmp/.X11-unix
chmod 1777 /tmp/.X11-unix

# Find the appropriate browser binary
if [ -f "/usr/bin/google-chrome-stable" ]; then
    BROWSER="/usr/bin/google-chrome-stable"
elif [ -f "/usr/local/bin/chromium" ]; then
    BROWSER="/usr/local/bin/chromium"
else
    echo "Error: No supported browser found"
    exit 1
fi

# Browser flags optimized for Docker containers
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
  --user-data-dir=/tmp/chrome-user-data-$(whoami) \
  "$@"
EOF

# Make wrapper executable, fix line endings, and create symlinks
RUN chmod +x /usr/local/bin/chrome-wrapper \
    && sed -i 's/\r$//' /usr/local/bin/chrome-wrapper \
    && ln -sf /usr/local/bin/chrome-wrapper /usr/local/bin/google-chrome \
    && ln -sf /usr/local/bin/chrome-wrapper /usr/local/bin/chrome \
    && ln -sf /usr/local/bin/chrome-wrapper /usr/local/bin/chromium \
    && update-alternatives --install /usr/bin/x-www-browser x-www-browser /usr/local/bin/chrome-wrapper 10 \
    && update-alternatives --install /usr/bin/gnome-www-browser gnome-www-browser /usr/local/bin/chrome-wrapper 10

# Create Chrome managed policies to enforce password manager disable
RUN mkdir -p /etc/opt/chrome/policies/managed && \
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
    }' > /etc/opt/chrome/policies/managed/disable_features.json && \
    chmod 644 /etc/opt/chrome/policies/managed/disable_features.json

# Install Microsoft Edge (AMD64 only - official support)
RUN ARCH=$(dpkg --print-architecture) && \
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

# Create Edge wrapper script
RUN cat > /usr/local/bin/edge-wrapper << 'EOF'
#!/bin/bash
# Set display if not already set
export DISPLAY=${DISPLAY:-:99}

# Ensure temp directories exist
mkdir -p /tmp/.X11-unix
chmod 1777 /tmp/.X11-unix

# Find Microsoft Edge binary
if [ -f "/usr/bin/microsoft-edge-stable" ]; then
    BROWSER="/usr/bin/microsoft-edge-stable"
else
    echo "Error: Microsoft Edge not found"
    exit 1
fi

# Browser flags optimized for Docker containers
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

# Make Edge wrapper executable and fix line endings
RUN chmod +x /usr/local/bin/edge-wrapper \
    && sed -i 's/\r$//' /usr/local/bin/edge-wrapper \
    && ln -sf /usr/local/bin/edge-wrapper /usr/local/bin/microsoft-edge \
    && ln -sf /usr/local/bin/edge-wrapper /usr/local/bin/edge

# Create Edge managed policies to enforce password manager disable
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

# ============================================================================
# Stage 4: Final Assembly
# ============================================================================
FROM browser-install AS final

# Set environment variables for VS Code
ENV DISPLAY=:99
ENV DONT_PROMPT_WSL_INSTALL=1
ENV BROWSER=/usr/local/bin/chrome-wrapper
ENV CHROME_EXECUTABLE=/usr/local/bin/chrome-wrapper

# Create vscode user with predictable UID/GID (replacing base appuser)
ARG USER_ID=1000
ARG GROUP_ID=1000
# First remove the appuser from base image to free up UID/GID 1000
# Note: userdel may fail if user doesn't exist, groupdel may fail if group is primary for another user
RUN (userdel -r appuser 2>/dev/null || true) && \
    (groupdel appuser 2>/dev/null || true) && \
    # Check if vscode user already exists, if not create it
    (getent group vscode >/dev/null 2>&1 || groupadd -g $GROUP_ID vscode) && \
    (id -u vscode >/dev/null 2>&1 || useradd -u $USER_ID -g $GROUP_ID -m -s /bin/bash vscode) && \
    usermod -aG video,audio,dialout vscode && \
    # Verify user was created
    id vscode

# Create directories with proper permissions
RUN mkdir -p /home/vscode/.vscode-server /home/vscode/.vscode /home/vscode/.config/Code \
    /workspace /app/data /app/data/recordings /tmp/vscode-temp /tmp/chrome-user-data-vscode \
    /opt/m365-template/devTools/func \
    && chmod -R 755 /workspace \
    && chmod -R 755 /app/data \
    && chmod 1777 /tmp \
    && chown -R vscode:vscode /home/vscode \
    && chown -R vscode:vscode /workspace \
    && chown vscode:vscode /tmp/chrome-user-data-vscode \
    && chown -R vscode:vscode /opt/m365-template

# Set up working directory
WORKDIR /app

# Create extensions directories for build-time installation
RUN mkdir -p /app/build-extensions /app/build-extensions-optional

# Copy build-time extension installation script
COPY install-build-extensions.sh /usr/local/bin/install-build-extensions.sh
RUN chmod +x /usr/local/bin/install-build-extensions.sh \
    && sed -i 's/\r$//' /usr/local/bin/install-build-extensions.sh

# Copy build-time extensions directories (copy the entire directories, including README files)
COPY build-extensions/ /app/build-extensions/
COPY build-extensions-optional/ /app/build-extensions-optional/

# Install build-time extensions (properly install for both VS Code versions)
RUN /usr/local/bin/install-build-extensions.sh && \
    mkdir -p /home/vscode/.vscode/extensions /home/vscode/.vscode-insiders/extensions && \
    cp -r /root/.vscode/extensions/* /home/vscode/.vscode/extensions/ 2>/dev/null || true && \
    cp -r /root/.vscode-insiders/extensions/* /home/vscode/.vscode-insiders/extensions/ 2>/dev/null || true && \
    chown -R vscode:vscode /home/vscode/.vscode /home/vscode/.vscode-insiders

# Set up supervisord configuration (override base config)
COPY supervisord.conf /etc/supervisor/conf.d/supervisord.conf

# Copy VS Code specific scripts
COPY start.sh /app/start.sh
COPY entrypoint.sh /usr/local/bin/entrypoint.sh
COPY vscode-setup.sh /usr/local/bin/vscode-setup.sh
COPY azure-auth.sh /usr/local/bin/azure-auth.sh

# Copy default VS Code settings and set permissions
COPY settings/ /tmp/vscode-settings/
RUN mkdir -p /home/vscode/.config/Code/User && \
    cp -r /tmp/vscode-settings/* /home/vscode/.config/Code/User/ 2>/dev/null || true && \
    chown -R vscode:vscode /home/vscode/.config

# Make scripts executable and fix line endings
RUN chmod +x /usr/local/bin/entrypoint.sh \
    && chmod +x /usr/local/bin/vscode-setup.sh \
    && chmod +x /usr/local/bin/azure-auth.sh \
    && chmod +x /app/start.sh \
    && sed -i 's/\r$//' /usr/local/bin/entrypoint.sh \
    && sed -i 's/\r$//' /usr/local/bin/vscode-setup.sh \
    && sed -i 's/\r$//' /usr/local/bin/azure-auth.sh \
    && sed -i 's/\r$//' /app/start.sh

# Expose ports: noVNC web interface and docker service
EXPOSE 6080 6081

# Everything runs with appropriate users
# System services (X11, supervisord) run as root
# VS Code runs as vscode user for security

WORKDIR /workspace

LABEL org.opencontainers.image.title="VscUse ATK VS Code Image"
LABEL org.opencontainers.image.description="VscUse ATK VS Code Docker container with noVNC, Chrome, Edge and recording capabilities"
LABEL org.opencontainers.image.licenses=MIT
LABEL org.opencontainers.image.base.name="vscuse-base:latest"

ENTRYPOINT ["/usr/local/bin/entrypoint.sh"]
CMD ["/app/start.sh"]
