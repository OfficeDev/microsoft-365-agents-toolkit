#!/usr/bin/env bash
# Setup script for agent plugins & skills used by this repo.
#
# Run once on a fresh machine after cloning:
#   bash .github/scripts/setup-agent-skills.sh
#
# Requires: `copilot` CLI (GitHub Copilot CLI) and `npx` on PATH.

set -u

step() {
    local desc="$1"; shift
    echo ""
    echo "==> $desc"
    if "$@"; then
        echo "    OK"
    else
        echo "    FAILED (continuing)"
    fi
}

# --- Prerequisite checks ---------------------------------------------------
command -v copilot >/dev/null 2>&1 || \
    echo "WARN: 'copilot' CLI not found. Install from https://docs.github.com/copilot/github-copilot-in-the-cli"
command -v npx >/dev/null 2>&1 || \
    echo "WARN: 'npx' not found. Install Node.js 18+."

# --- 1. Copilot CLI marketplaces ------------------------------------------
marketplaces=(
    "obra/superpowers-marketplace"
    "anthropics/skills"
)
for mp in "${marketplaces[@]}"; do
    step "Register marketplace $mp" \
        copilot plugin marketplace add "$mp"
done

# --- 2. Copilot CLI plugins -----------------------------------------------
# Note: copilot plugins install user-globally to ~/.copilot/installed-plugins/,
# not into the repo. They're available to the `copilot` CLI from any directory.
copilot_plugins=(
    "superpowers@superpowers-marketplace"
    "document-skills@anthropic-agent-skills"
    "example-skills@anthropic-agent-skills"
    # Note: 'frontend-design' is NOT a plugin in anthropics/skills.
    # It exists only as a skill folder; installed via npx skills add below.
)
for plugin in "${copilot_plugins[@]}"; do
    step "copilot plugin install $plugin" \
        copilot plugin install "$plugin"
done

# --- 3. `npx skills add` entries ------------------------------------------
# These install into .agents/skills/ (the github-copilot agent dir).
# Each entry: "<source>|<skill1>,<skill2>" (use '*' for all skills)
npx_skills=(
    "https://github.com/microsoft/playwright-cli|*"
    "OthmanAdi/planning-with-files|planning-with-files"
    "anthropics/skills|frontend-design"
    "https://github.com/awesome-skills/code-review-skill|code-review-excellence"
    # TODO: fill in the actual repos below before running on a clean machine
    # "<code-review-skill-repo>|*"
    # "<commit-helper-repo>|*"
)
for entry in "${npx_skills[@]}"; do
    src="${entry%%|*}"
    skills_csv="${entry##*|}"
    skill_args=()
    IFS=',' read -r -a skill_list <<< "$skills_csv"
    for s in "${skill_list[@]}"; do
        skill_args+=(--skill "$s")
    done
    # Install only the specified skills into the github-copilot agent
    # directory (.github/). Avoids creating .claude/.cursor/.windsurf etc
    # folders, and avoids language variants we don't need.
    # See `npx skills add --help`.
    step "npx skills add $src (github-copilot, skills=$skills_csv)" \
        npx --yes skills add "$src" --yes --agent github-copilot "${skill_args[@]}"
done

echo ""
echo "Done."
echo "Next steps:"
echo "  - Fill in the remaining TODO entry (commit-helper) in this script."
echo "  - Run 'copilot' in the repo root to use the installed plugins."
