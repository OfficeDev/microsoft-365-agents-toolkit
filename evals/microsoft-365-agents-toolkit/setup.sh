# shellcheck shell=bash
if set -euo pipefail 2>/dev/null; then
  :
else
  set -euo
fi

export CI_ENABLED="${CI_ENABLED:-true}"

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
skill_dependencies_dir="$repo_root/node_modules/.waza-skills"
wiqd_live_skill_dir="$HOME/.wiqd/marketplace/wiqd/skills/wiqd"
wiqd_installed_skill_dir="$HOME/.copilot/installed-plugins/wiqd/wiqd/skills/wiqd"
azure_skills_dir="$HOME/.copilot/installed-plugins/azure-skills/azure/skills"

if [ -d "$wiqd_live_skill_dir" ]; then
	wiqd_skill_dir="$wiqd_live_skill_dir"
elif [ -d "$wiqd_installed_skill_dir" ]; then
	wiqd_skill_dir="$wiqd_installed_skill_dir"
else
	echo "Missing WIQD plugin skill directory." >&2
	exit 1
fi

if [ ! -d "$azure_skills_dir" ]; then
	echo "Missing Azure plugin skills directory: $azure_skills_dir" >&2
	exit 1
fi

rm -rf "$skill_dependencies_dir"
mkdir -p "$skill_dependencies_dir"
cp -R "$wiqd_skill_dir" "$skill_dependencies_dir/"

npm install -g @microsoft/m365agentstoolkit-cli@beta
