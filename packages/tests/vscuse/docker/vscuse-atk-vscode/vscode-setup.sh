#!/bin/bash
# VS Code Setup Script - installs VSIX extensions and starts VS Code

set -e  # Exit immediately if a command exits with a non-zero status

# Function to log with timestamp
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1"
}

# Configuration
EXTENSIONS_DIR="${EXTENSIONS_DIR:-/tmp/extensions}"
DISPLAY=":99"
INSTALL_TIMEOUT=180

# Determine VS Code binary based on environment variable
if [ "${VSCODE_INSIDERS:-false}" = "true" ]; then
    VSCODE_BIN="code-insiders"
    log "Using VS Code Insider"
else
    VSCODE_BIN="code"
    log "Using VS Code stable"
fi

# Verify the binary exists and is executable
if ! command -v "$VSCODE_BIN" >/dev/null 2>&1; then
    log "ERROR: VS Code binary '$VSCODE_BIN' not found or not executable"
    log "Available VS Code binaries:"
    command -v code >/dev/null 2>&1 && log "  - code (stable): available"
    command -v code-insiders >/dev/null 2>&1 && log "  - code-insiders (insider): available"

    # Fallback logic
    if [ "${VSCODE_INSIDERS:-false}" = "true" ] && command -v "code" >/dev/null 2>&1; then
        VSCODE_BIN="code"
        log "Falling back to VS Code stable since Insider is not available"
    elif [ "${VSCODE_INSIDERS:-false}" = "false" ] && command -v "code-insiders" >/dev/null 2>&1; then
        # This case is less likely but included for completeness
        log "VS Code stable requested but only Insider available - this is unexpected"
        exit 1
    else
        log "ERROR: No VS Code binary found"
        exit 1
    fi
fi

log "Final VS Code binary: $VSCODE_BIN"

# Function to clear workspace directory
clear_workspace() {
    log "Clearing /workspace directory"

    if [ -d "/workspace" ]; then
        # Remove all contents of /workspace but keep the directory itself
        rm -rf /workspace/* /workspace/.[^.]* 2>/dev/null || true
        log "Workspace directory cleared"
    else
        # Create workspace directory if it doesn't exist
        mkdir -p /workspace
        log "Workspace directory created"
    fi
}

# Function to install a single VSIX extension
install_vsix() {
    local vsix_file="$1"
    local filename=$(basename "$vsix_file")
    local start_time=$(date +%s)
    local temp_user_data="$2"

    log "Installing: $filename"

    # Set environment variables for extension installation
    export HOME=/home/vscode
    export XDG_CONFIG_HOME=/home/vscode/.config

    # Run VS Code extension installation with temp user data directory
    if timeout $INSTALL_TIMEOUT env HOME=/home/vscode $VSCODE_BIN --install-extension "$vsix_file" --user-data-dir="$temp_user_data" --force >/dev/null 2>&1; then
        local end_time=$(date +%s)
        local duration=$((end_time - start_time))
        log "Installed: $filename (${duration}s)"
        return 0
    else
        local end_time=$(date +%s)
        local duration=$((end_time - start_time))
        log "Failed: $filename (${duration}s)"
        return 1
    fi
}

# Function to install all extensions
install_extensions() {
    local temp_user_data="$1"
    local phase_start_time=$(date +%s)
    log "Phase 2: Extension installation"

    # Quick check - if no extensions directory, skip immediately
    if [ ! -d "$EXTENSIONS_DIR" ]; then
        return 0
    fi

    # Fast check for VSIX files using shell glob (faster than find)
    local has_extensions=false
    for vsix_file in "$EXTENSIONS_DIR"/*.vsix; do
        if [ -f "$vsix_file" ]; then
            has_extensions=true
            break
        fi
    done

    if [ "$has_extensions" = false ]; then
        return 0
    fi

    # Count and install extensions
    local failed_count=0
    local installed_count=0
    local total_count=0

    for vsix_file in "$EXTENSIONS_DIR"/*.vsix; do
        [ -f "$vsix_file" ] || continue
        total_count=$((total_count + 1))

        if install_vsix "$vsix_file" "$temp_user_data"; then
            installed_count=$((installed_count + 1))
        else
            failed_count=$((failed_count + 1))
            log "ERROR: Extension installation failed"
            exit 1  # Fail fast as requested
        fi
    done

    # Summary logging
    if [ $failed_count -gt 0 ]; then
        log "ERROR: $failed_count extension(s) failed to install"
        exit 1
    elif [ $installed_count -gt 0 ]; then
        log "Extensions: $installed_count installed fresh"
    fi

    local phase_end_time=$(date +%s)
    local phase_duration=$((phase_end_time - phase_start_time))
    log "Phase 2 completed (${phase_duration}s)"
}

# Main function
main() {
    log "Phase 1: Starting VS Code with fresh user data"

    # Create completely fresh temporary user data directory each time
    TEMP_USER_DATA="/tmp/vscode-userdata-$(date +%s)-$$"
    mkdir -p "$TEMP_USER_DATA/User"
    chmod -R 755 "$TEMP_USER_DATA"

    # Copy our default settings to the fresh temp location
    if [ -f "/tmp/vscode-settings/settings.json" ]; then
        cp "/tmp/vscode-settings/settings.json" "$TEMP_USER_DATA/User/settings.json"
        log "Applied fresh settings to temporary user data directory"
    fi

    # Also copy from the vscode user's config if available
    if [ -f "/home/vscode/.config/Code/User/settings.json" ]; then
        cp "/home/vscode/.config/Code/User/settings.json" "$TEMP_USER_DATA/User/settings.json"
        log "Applied vscode user settings to temporary user data directory"
    fi

    # Install extensions to the temp user data directory
    install_extensions "$TEMP_USER_DATA"

    # Clear workspace directory before starting VS Code
    # clear_workspace

    # Debug: Check permissions before starting VS Code
    log "Debug: User data directory: $TEMP_USER_DATA"
    log "Debug: Directory permissions: $(ls -la /tmp | grep vscode-userdata || echo 'Not found')"
    log "Debug: Current user: $(whoami)"
    log "Debug: Build-time extensions: $(ls -la /home/vscode/.vscode/extensions 2>/dev/null | wc -l || echo '0') extensions found"
    log "Debug: Runtime extensions: $(ls -la $TEMP_USER_DATA/extensions 2>/dev/null | wc -l || echo '0') extensions found"

    # Start VS Code with completely fresh state - no persistence at all
    exec $VSCODE_BIN --no-sandbox --disable-dev-shm-usage --disable-gpu --password-store=basic --user-data-dir="$TEMP_USER_DATA" /workspace --wait
}

# Run main function
main "$@"
