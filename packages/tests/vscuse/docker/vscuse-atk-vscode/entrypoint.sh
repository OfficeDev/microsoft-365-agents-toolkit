#!/usr/bin/env bash
set -e

echo "Starting VscUse VS Code Docker container..."

# Set umask for proper file permissions (permissive for Docker environment)
umask 000

# Clear workspace directory for fresh start
echo "Clearing workspace directory..."
if [ -d "/workspace" ]; then
    # Remove all contents of /workspace but keep the directory itself
    rm -rf /workspace/* /workspace/.[^.]* 2>/dev/null || true
    echo "Workspace directory cleared"
fi

# Ensure vscode user has access to workspace
chown -R vscode:vscode /workspace 2>/dev/null || true
echo "Workspace permissions set for vscode user"

echo "Container initialization complete. Executing command: $*"

# Execute the command passed to the container
exec "$@"