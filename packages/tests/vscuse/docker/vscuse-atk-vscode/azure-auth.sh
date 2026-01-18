#!/bin/bash
#
# Azure Authentication Script for VSC-Use Docker Container
#
# This script handles automatic Azure login using service principal credentials
# when the container starts up. It's called from the main startup scripts.
#

set -e

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Azure Auth: $1"
}

# Function to authenticate with Azure using service principal
authenticate_azure() {
    local client_id="$1"
    local client_secret="$2"
    local tenant_id="$3"

    if [[ -z "$client_id" || -z "$client_secret" || -z "$tenant_id" ]]; then
        log "Missing Azure credentials, skipping authentication"
        return 0
    fi

    log "Attempting Azure authentication with service principal..."
    log "Running as user: $(whoami), HOME: ${HOME:-/home/vscode}"

    # Ensure proper environment for vscode user
    export HOME="${HOME:-/home/vscode}"
    export USER="${USER:-vscode}"


    # Try to authenticate with ATK CLI
    if command -v atk >/dev/null 2>&1; then
        log "ATK CLI found, authenticating with service principal..."

        if atk auth login azure --interactive false --service-principal \
           -u "$client_id" \
           -p "$client_secret" \
           --tenant "$tenant_id" 2>&1; then
            log "Successfully authenticated with Azure via ATK CLI"

            # Verify authentication
            if atk auth list 2>&1; then
                log "Authentication verification successful"
            else
                log "Warning: Could not verify authentication status"
            fi
        else
            log "Failed to authenticate with Azure via ATK CLI"
            return 1
        fi
    else
        log "ATK CLI not found, skipping Azure authentication"
        return 1
    fi
}

# Main execution
main() {
    log "Starting Azure authentication process..."
    log "Current user: $(whoami)"
    log "Current HOME: ${HOME:-<not set>}"
    log "Environment variables:"
    log "  AZURE_AUTH_ENABLED: ${AZURE_AUTH_ENABLED:-<not set>}"
    [[ -n "$AZURE_CLIENT_ID" ]] && log "  AZURE_CLIENT_ID: <set>" || log "  AZURE_CLIENT_ID: <not set>"
    [[ -n "$AZURE_CLIENT_SECRET" ]] && log "  AZURE_CLIENT_SECRET: <set>" || log "  AZURE_CLIENT_SECRET: <not set>"
    [[ -n "$AZURE_TENANT_ID" ]] && log "  AZURE_TENANT_ID: <set>" || log "  AZURE_TENANT_ID: <not set>"

    # Get credentials from environment variables
    local client_id="${AZURE_CLIENT_ID:-}"
    local client_secret="${AZURE_CLIENT_SECRET:-}"
    local tenant_id="${AZURE_TENANT_ID:-}"
    local auth_enabled="${AZURE_AUTH_ENABLED:-false}"

    # Check if authentication is enabled
    if [[ "$auth_enabled" != "true" ]]; then
        log "Azure authentication disabled (AZURE_AUTH_ENABLED != true)"
        return 0
    fi

    # Perform authentication
    if authenticate_azure "$client_id" "$client_secret" "$tenant_id"; then
        log "Azure authentication completed successfully"
        return 0
    else
        log "Azure authentication failed"
        return 1
    fi
}

# Only run if called directly (not sourced)
if [[ "${BASH_SOURCE[0]}" == "${0}" ]]; then
    main "$@"
fi
