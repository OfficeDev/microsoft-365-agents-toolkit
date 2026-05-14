# Setup script for agent plugins & skills used by this repo.
#
# Run once on a fresh machine after cloning:
#   pwsh .github/scripts/setup-agent-skills.ps1
#
# Requires: `copilot` CLI (GitHub Copilot CLI) and `npx` on PATH.

$ErrorActionPreference = "Stop"

function Invoke-Step {
    param(
        [Parameter(Mandatory)] [string] $Description,
        [Parameter(Mandatory)] [scriptblock] $Action
    )
    Write-Host ""
    Write-Host "==> $Description" -ForegroundColor Cyan
    try {
        & $Action
        Write-Host "    OK" -ForegroundColor Green
    }
    catch {
        Write-Host "    FAILED: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "    (continuing)" -ForegroundColor Yellow
    }
}

# --- Prerequisite checks ---------------------------------------------------
if (-not (Get-Command copilot -ErrorAction SilentlyContinue)) {
    Write-Warning "GitHub Copilot CLI ('copilot') not found on PATH. Install from https://docs.github.com/copilot/github-copilot-in-the-cli and re-run."
}
if (-not (Get-Command npx -ErrorAction SilentlyContinue)) {
    Write-Warning "'npx' not found on PATH. Install Node.js 18+ and re-run."
}

# --- 1. Copilot CLI marketplaces ------------------------------------------
$marketplaces = @(
    "obra/superpowers-marketplace",
    "anthropics/skills"
)
foreach ($mp in $marketplaces) {
    Invoke-Step "Register marketplace $mp" {
        copilot plugin marketplace add $mp
    }
}

# --- 2. Copilot CLI plugins -----------------------------------------------
# Note: copilot plugins install user-globally to ~/.copilot/installed-plugins/,
# not into the repo. They're available to the `copilot` CLI from any directory.
$copilotPlugins = @(
    "superpowers@superpowers-marketplace",
    "document-skills@anthropic-agent-skills",
    "example-skills@anthropic-agent-skills"
    # Note: 'frontend-design' is NOT a plugin in anthropics/skills.
    # It exists only as a skill folder; installed via npx skills add below.
)
foreach ($plugin in $copilotPlugins) {
    Invoke-Step "copilot plugin install $plugin" {
        copilot plugin install $plugin
    }
}

# --- 3. `npx skills add` entries ------------------------------------------
# These install into .agents/skills/ (the github-copilot agent dir).
# Each entry: @{ source = '<repo>'; skills = @('name1','name2') } (skills='*' for all)
$npxSkills = @(
    @{ source = "https://github.com/microsoft/playwright-cli"; skills = @("*") },
    @{ source = "OthmanAdi/planning-with-files";              skills = @("planning-with-files") },
    @{ source = "anthropics/skills";                          skills = @("frontend-design") },
    @{ source = "https://github.com/awesome-skills/code-review-skill"; skills = @("code-review-excellence") }
    # TODO: fill in the actual repos below before running on a clean machine
    # @{ source = "<code-review-skill-repo>"; skills = @("*") },
    # @{ source = "<commit-helper-repo>";    skills = @("*") }
)
foreach ($entry in $npxSkills) {
    $src = $entry.source
    $skillArgs = @()
    foreach ($s in $entry.skills) { $skillArgs += @("--skill", $s) }
    Invoke-Step "npx skills add $src (github-copilot, skills=$($entry.skills -join ','))" {
        # Install only the specified skills into the github-copilot agent
        # directory (.github/). Avoids creating .claude/.cursor/.windsurf etc
        # folders, and avoids language variants we don't need.
        # See `npx skills add --help`.
        npx --yes skills add $src --yes --agent github-copilot @skillArgs
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Next steps:"
Write-Host "  - Fill in the remaining TODO entry (commit-helper) in this script."
Write-Host "  - Run 'copilot' in the repo root to use the installed plugins."
