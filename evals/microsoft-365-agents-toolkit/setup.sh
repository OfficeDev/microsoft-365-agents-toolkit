# shellcheck shell=bash
if set -euo pipefail 2>/dev/null; then
  :
else
  set -euo
fi

export CI_ENABLED="${CI_ENABLED:-true}"

if command -v wiqd >/dev/null 2>&1 && wiqd --version >/dev/null 2>&1; then
	echo "WIQD is available for Declarative Agent execution tasks."
else
	cat >&2 <<'EOF'
WIQD is unavailable or its EULA has not been accepted. Declarative Agent
execution tasks require WIQD to be installed and initialized interactively;
they will not pass in this environment. Other evaluation tasks can continue.
EOF
fi

# Values are injected by CI (for GitHub Actions, use repository secrets).
required_vars=(
	M365_ACCOUNT_NAME
	M365_ACCOUNT_PASSWORD
	M365_TENANT_ID
	AZURE_ACCOUNT_NAME
	AZURE_ACCOUNT_OBJECT_ID
	AZURE_ACCOUNT_PASSWORD
	AZURE_TENANT_ID
	AZURE_SUBSCRIPTION_ID
)

for var_name in "${required_vars[@]}"; do
	if [ -z "${!var_name:-}" ]; then
		echo "Missing required environment variable: ${var_name}" >&2
		exit 1
	fi
done

npm install -g @microsoft/m365agentstoolkit-cli@beta
# az login -u "$AZURE_ACCOUNT_NAME" -p "$AZURE_ACCOUNT_PASSWORD" --tenant "$AZURE_TENANT_ID"