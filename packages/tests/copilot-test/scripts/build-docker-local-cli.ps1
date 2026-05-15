<#
.SYNOPSIS
    Build ATK CLI from source and bake it into the Docker image.

.DESCRIPTION
    1. Runs `pnpm build` in the ATK CLI package
    2. Runs `npm pack` to produce a .tgz
    3. Copies the .tgz to packages\tests\copilot-test\docker\cli-local.tgz
    4. Builds the Docker image with ATK_CLI_SOURCE=local

.PARAMETER AtkSource
    Path to microsoft-365-agents-toolkit repo root.
    Default: C:\Users\quke\source\atk\microsoft-365-agents-toolkit

.PARAMETER ImageTag
    Docker image tag. Default: atk-copilot-test-local

.PARAMETER SkipBuild
    Skip pnpm build (use if you already built the CLI).

.PARAMETER PrepOnly
    Only copy the .tgz; skip docker build.

.EXAMPLE
    # Full flow: build CLI + build Docker image
    .\scripts\build-docker-local-cli.ps1

.EXAMPLE
    # Use pre-built CLI, skip pnpm build
    .\scripts\build-docker-local-cli.ps1 -SkipBuild

.EXAMPLE
    # Just prep .tgz, build Docker yourself
    .\scripts\build-docker-local-cli.ps1 -PrepOnly
    docker build --build-arg ATK_CLI_SOURCE=local -t my-tag -f packages\tests\copilot-test\docker\Dockerfile .
#>
param(
    [string]$AtkSource = "C:\Users\quke\source\atk\microsoft-365-agents-toolkit",
    [string]$ImageTag  = "atk-copilot-test-local",
    [switch]$SkipBuild,
    [switch]$PrepOnly
)

$ErrorActionPreference = "Stop"
$repoRoot  = Split-Path $PSScriptRoot
$dockerDir = Join-Path $repoRoot "packages\tests\copilot-test\docker"
$cliPkg    = Join-Path $AtkSource "packages\cli"

# 1. Validate
if (-not (Test-Path $cliPkg)) {
    Write-Error "ATK CLI package not found at: $cliPkg`nSet -AtkSource to your ATK repo root."
}
Write-Host "==> ATK source : $AtkSource"
Write-Host "==> CLI package: $cliPkg"

# 2. Build CLI
if (-not $SkipBuild) {
    Write-Host "`n==> Building ATK CLI (pnpm build)..."
    Push-Location $cliPkg
    try {
        & pnpm build
        if ($LASTEXITCODE -ne 0) { throw "pnpm build failed" }
    } finally { Pop-Location }
} else {
    Write-Host "`n==> Skipping pnpm build (-SkipBuild)"
}

# 3. npm pack
Write-Host "`n==> Packing CLI with npm pack..."
Push-Location $cliPkg
try {
    Get-ChildItem -Filter "*.tgz" | Remove-Item -Force
    & npm pack
    if ($LASTEXITCODE -ne 0) { throw "npm pack failed" }
    $tgz = Get-ChildItem -Filter "*.tgz" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $tgz) { throw "npm pack produced no .tgz file" }
    Write-Host "==> Packed: $($tgz.FullName)  ($([math]::Round($tgz.Length/1KB))KB)"
    $tgzPath = $tgz.FullName
} finally { Pop-Location }

# 4. Copy to docker context
$dest = Join-Path $dockerDir "cli-local.tgz"
Copy-Item $tgzPath $dest -Force
Write-Host "==> Copied to: $dest"

if ($PrepOnly) {
    Write-Host "`n==> PrepOnly — skipping docker build. Run manually:"
    Write-Host "      docker build --build-arg ATK_CLI_SOURCE=local -t $ImageTag -f packages\tests\copilot-test\docker\Dockerfile ."
    exit 0
}

# 5. Docker build
Write-Host "`n==> Building Docker image: $ImageTag ..."
Push-Location $repoRoot
try {
    & docker build --build-arg ATK_CLI_SOURCE=local -t $ImageTag -f "packages\tests\copilot-test\docker\Dockerfile" .
    if ($LASTEXITCODE -ne 0) { throw "docker build failed" }
} finally { Pop-Location }

Write-Host "`n✅ Done! Image: $ImageTag"
Write-Host "`nRun tests:"
Write-Host "  docker run --rm --shm-size=512m ``"
Write-Host "    -v `"C:\path\to\vscode-extension:/atk-ext:ro`" ``"
Write-Host "    -v `"`$(pwd)/test-output:/output`" ``"
Write-Host "    -e TEST_FILE=simple-bot-create ``"
Write-Host "    $ImageTag"
