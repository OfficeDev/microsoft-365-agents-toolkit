#!/bin/bash
# Build-time VS Code Extension Installation Script

set -e  # Exit immediately if a command exits with a non-zero status

# Configuration
REQUIRED_EXTENSIONS_DIR="/app/build-extensions"
OPTIONAL_EXTENSIONS_DIR="/app/build-extensions-optional"
STABLE_USER_DATA_DIR="/root/.vscode"
INSIDER_USER_DATA_DIR="/root/.vscode-insiders"
INSTALL_TIMEOUT=180

# Function to log with timestamp
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] [BUILD-TIME] $1"
}

# Function to install an extension by ID from marketplace
install_extension_id() {
    local ext_id="$1"
    local is_optional="$2"
    local start_time=$(date +%s)

    log "Installing from marketplace: $ext_id ${is_optional:+(optional)}"

    # Install for VS Code stable first
    local stable_success=false
    if timeout $INSTALL_TIMEOUT bash -c "
        export HOME=/root
        export XDG_CONFIG_HOME=/root/.config
        export DONT_PROMPT_WSL_INSTALL=1
        echo 'n' | code --install-extension '$ext_id' --user-data-dir='$STABLE_USER_DATA_DIR' --force --disable-extensions
    " 2>/tmp/vscode-install-stable.log; then
        log "✓ Installed for VS Code stable: $ext_id"
        stable_success=true
    else
        log "✗ Failed for VS Code stable: $ext_id"
        if [ -f /tmp/vscode-install-stable.log ]; then
            log "Stable error: $(tail -2 /tmp/vscode-install-stable.log | tr '\n' ' ')"
        fi
    fi

    # Install for VS Code Insider
    local insider_success=false
    if timeout $INSTALL_TIMEOUT bash -c "
        export HOME=/root
        export XDG_CONFIG_HOME=/root/.config
        export DONT_PROMPT_WSL_INSTALL=1
        echo 'n' | code-insiders --install-extension '$ext_id' --user-data-dir='$INSIDER_USER_DATA_DIR' --force --disable-extensions
    " 2>/tmp/vscode-install-insider.log; then
        log "✓ Installed for VS Code Insider: $ext_id"
        insider_success=true
    else
        log "✗ Failed for VS Code Insider: $ext_id"
        if [ -f /tmp/vscode-install-insider.log ]; then
            log "Insider error: $(tail -2 /tmp/vscode-install-insider.log | tr '\n' ' ')"
        fi
    fi

    local end_time=$(date +%s)
    local duration=$((end_time - start_time))

    # Consider success if at least one version succeeded
    if [ "$stable_success" = true ] || [ "$insider_success" = true ]; then
        local status=""
        [ "$stable_success" = true ] && [ "$insider_success" = true ] && status="both versions"
        [ "$stable_success" = true ] && [ "$insider_success" = false ] && status="stable only"
        [ "$stable_success" = false ] && [ "$insider_success" = true ] && status="insider only"

        log "✓ Completed: $ext_id ($status, ${duration}s)"
        return 0
    else
        log "✗ Failed: $ext_id (both versions failed, ${duration}s)"
        return 1
    fi
}

# Function to install a single VSIX extension
install_vsix() {
    local vsix_file="$1"
    local filename=$(basename "$vsix_file")
    local is_optional="$2"
    local start_time=$(date +%s)

    log "Installing: $filename ${is_optional:+(optional)}"

    # Install for VS Code stable first
    local stable_success=false
    if timeout $INSTALL_TIMEOUT bash -c "
        export HOME=/root
        export XDG_CONFIG_HOME=/root/.config
        export DONT_PROMPT_WSL_INSTALL=1
        echo 'n' | code --install-extension '$vsix_file' --user-data-dir='$STABLE_USER_DATA_DIR' --force --disable-extensions
    " 2>/tmp/vscode-install-stable.log; then
        log "✓ Installed for VS Code stable: $filename"
        stable_success=true
    else
        log "✗ Failed for VS Code stable: $filename"
        if [ -f /tmp/vscode-install-stable.log ]; then
            log "Stable error: $(tail -2 /tmp/vscode-install-stable.log | tr '\n' ' ')"
        fi
    fi

    # Install for VS Code Insider
    local insider_success=false
    if timeout $INSTALL_TIMEOUT bash -c "
        export HOME=/root
        export XDG_CONFIG_HOME=/root/.config
        export DONT_PROMPT_WSL_INSTALL=1
        echo 'n' | code-insiders --install-extension '$vsix_file' --user-data-dir='$INSIDER_USER_DATA_DIR' --force --disable-extensions
    " 2>/tmp/vscode-install-insider.log; then
        log "✓ Installed for VS Code Insider: $filename"
        insider_success=true
    else
        log "✗ Failed for VS Code Insider: $filename"
        if [ -f /tmp/vscode-install-insider.log ]; then
            log "Insider error: $(tail -2 /tmp/vscode-install-insider.log | tr '\n' ' ')"
        fi
    fi

    local end_time=$(date +%s)
    local duration=$((end_time - start_time))

    # Consider success if at least one version succeeded
    if [ "$stable_success" = true ] || [ "$insider_success" = true ]; then
        local status=""
        [ "$stable_success" = true ] && [ "$insider_success" = true ] && status="both versions"
        [ "$stable_success" = true ] && [ "$insider_success" = false ] && status="stable only"
        [ "$stable_success" = false ] && [ "$insider_success" = true ] && status="insider only"

        log "✓ Completed: $filename ($status, ${duration}s)"
        return 0
    else
        log "✗ Failed: $filename (both versions failed, ${duration}s)"
        return 1
    fi
}

# Function to install extensions from a directory
install_extensions_from_dir() {
    local ext_dir="$1"
    local is_optional="$2"
    local installed_count=0
    local failed_count=0
    local total_count=0

    if [ ! -d "$ext_dir" ]; then
        return 0
    fi

    log "Scanning directory: $ext_dir ${is_optional:+(optional)}"

    # Check for VSIX files (ignore README.md and other non-extension files)
    local has_extensions=false
    for vsix_file in "$ext_dir"/*.vsix; do
        if [ -f "$vsix_file" ]; then
            has_extensions=true
            break
        fi
    done

    if [ "$has_extensions" = false ]; then
        log "No .vsix files found in $ext_dir"
        return 0
    fi

    # Install extensions (only process .vsix files, skip README.md and other files)
    for vsix_file in "$ext_dir"/*.vsix; do
        [ -f "$vsix_file" ] || continue

        # Skip if this is just the glob pattern (no actual files)
        if [[ "$vsix_file" == "$ext_dir/*.vsix" ]]; then
            continue
        fi

        total_count=$((total_count + 1))

        if install_vsix "$vsix_file" "$is_optional"; then
            installed_count=$((installed_count + 1))
        else
            failed_count=$((failed_count + 1))

            # For required extensions, fail the build
            if [ -z "$is_optional" ]; then
                log "ERROR: Required extension installation failed"
                exit 1
            else
                log "WARNING: Optional extension installation failed"
            fi
        fi
    done

    log "Directory $ext_dir: $installed_count/$total_count extensions installed successfully"
}

# Function to install extensions from extensions.txt file
install_extensions_from_file() {
    local ext_file="$1"
    local is_optional="$2"
    local installed_count=0
    local failed_count=0
    local total_count=0

    if [ ! -f "$ext_file" ]; then
        return 0
    fi

    log "Processing extension IDs from: $ext_file ${is_optional:+(optional)}"

    # Read extensions.txt line by line
    while IFS= read -r line || [ -n "$line" ]; do
        # Remove carriage return and trim whitespace
        ext_id=$(echo "$line" | tr -d '\r' | xargs)

        # Skip empty lines and comments
        [[ -z "$ext_id" || "$ext_id" =~ ^# ]] && continue

        total_count=$((total_count + 1))

        if install_extension_id "$ext_id" "$is_optional"; then
            installed_count=$((installed_count + 1))
        else
            failed_count=$((failed_count + 1))

            # For required extensions, fail the build
            if [ -z "$is_optional" ]; then
                log "ERROR: Required extension installation failed: $ext_id"
                exit 1
            else
                log "WARNING: Optional extension installation failed: $ext_id"
            fi
        fi
    done < "$ext_file"

    if [ $total_count -gt 0 ]; then
        log "Extension IDs from $ext_file: $installed_count/$total_count installed successfully"
    else
        log "No extension IDs found in $ext_file"
    fi
}

# Main execution
main() {
    log "=== Build-time Extension Installation Started ==="
    local script_start_time=$(date +%s)

    # Ensure user data directories exist for both VS Code versions
    mkdir -p "$STABLE_USER_DATA_DIR"
    mkdir -p "$INSIDER_USER_DATA_DIR"

    # Install required extensions (fail build if any fail)
    log "Phase 1: Installing required extensions"
    install_extensions_from_file "$REQUIRED_EXTENSIONS_DIR/extensions.txt" ""
    install_extensions_from_dir "$REQUIRED_EXTENSIONS_DIR" ""

    # Install optional extensions (continue build even if they fail)
    log "Phase 2: Installing optional extensions"
    install_extensions_from_file "$OPTIONAL_EXTENSIONS_DIR/extensions.txt" "optional"
    install_extensions_from_dir "$OPTIONAL_EXTENSIONS_DIR" "optional"

    local script_end_time=$(date +%s)
    local script_duration=$((script_end_time - script_start_time))
    log "=== Build-time Extension Installation Completed in ${script_duration}s ==="
}

# Only run if being executed directly (not sourced)
if [ "${BASH_SOURCE[0]}" = "${0}" ]; then
    main "$@"
fi
