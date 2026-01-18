# ============================================================================
# VscUse VS Code Docker Build Script
# ============================================================================
# This image depends on vscuse-base:latest
# The script will check for the base image and offer to build it if missing
# ============================================================================

# Add parameters for enhanced build options
param(
    [switch]$NoCache,
    [switch]$Clean,
    [switch]$Push,
    [switch]$MultiPlatform,
    [switch]$SkipBaseCheck,
    [string]$Tag = "latest",
    [string]$Repository = "microsoft",
    [string]$Platforms = "linux/amd64,linux/arm64"
)

# PowerShell version of build script
$IMAGE_NAME = "vscuse-atk-vscode"
$BASE_IMAGE = "vscuse-base:latest"

# Check if base image exists (unless skipped)
if (-not $SkipBaseCheck) {
    Write-Host "Checking for base image: $BASE_IMAGE" -ForegroundColor Cyan
    $baseImageExists = docker images -q $BASE_IMAGE 2>$null
    if (-not $baseImageExists) {
        Write-Host "Base image '$BASE_IMAGE' not found!" -ForegroundColor Yellow
        Write-Host "Building base image first..." -ForegroundColor Cyan
        
        $baseDir = Join-Path $PSScriptRoot "..\vscuse-base"
        if (Test-Path $baseDir) {
            Push-Location $baseDir
            try {
                & .\build.ps1
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "Failed to build base image!" -ForegroundColor Red
                    exit 1
                }
                Write-Host "Base image built successfully!" -ForegroundColor Green
            } finally {
                Pop-Location
            }
        } else {
            Write-Host "Base image directory not found: $baseDir" -ForegroundColor Red
            Write-Host "Please build vscuse-base:latest manually first:" -ForegroundColor Yellow
            Write-Host "  cd ../vscuse-base && docker build -t vscuse-base:latest ." -ForegroundColor Gray
            exit 1
        }
    } else {
        Write-Host "Base image found: $BASE_IMAGE" -ForegroundColor Green
    }
}
$IMAGE_VERSION = $Tag

# Always use GitHub Container Registry
$REGISTRY = "ghcr.io/${Repository}"
$FULL_IMAGE_NAME = "${REGISTRY}/${IMAGE_NAME}:${IMAGE_VERSION}"

Write-Host "=== VSCode Docker Build Script ===" -ForegroundColor Green
Write-Host "Registry: GitHub Container Registry (ghcr.io)" -ForegroundColor Cyan
Write-Host "Image: $FULL_IMAGE_NAME" -ForegroundColor Yellow
Write-Host "VS Code: Both stable and insider versions included" -ForegroundColor Cyan

if ($MultiPlatform) {
    Write-Host "Multi-platform build enabled for: $Platforms" -ForegroundColor Magenta
}

# Function to clean Docker system
function Clean-DockerSystem {
    Write-Host "Cleaning Docker system..." -ForegroundColor Yellow
    Write-Host "Current disk usage:" -ForegroundColor Cyan
    docker system df

    Write-Host "`nRemoving old image..." -ForegroundColor Yellow
    docker rmi $FULL_IMAGE_NAME -f 2>$null

    Write-Host "Pruning Docker system..." -ForegroundColor Yellow
    docker system prune -a --volumes -f

    Write-Host "Post-cleanup disk usage:" -ForegroundColor Cyan
    docker system df
}

# Function to setup Docker Buildx for multi-platform builds
function Setup-DockerBuildx {
    Write-Host "Setting up Docker Buildx for multi-platform builds..." -ForegroundColor Yellow

    # Check if buildx is available
    $buildxCheck = docker buildx version 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Docker Buildx is not available. Please update Docker Desktop." -ForegroundColor Red
        exit 1
    }

    # Create or use existing builder
    $builderName = "vscuse-multiarch-builder"

    # Check if builder exists
    $existingBuilder = docker buildx ls | Select-String $builderName
    if (-not $existingBuilder) {
        Write-Host "Creating new buildx builder: $builderName" -ForegroundColor Cyan
        docker buildx create --name $builderName --use --bootstrap
        if ($LASTEXITCODE -ne 0) {
            Write-Host "❌ Failed to create buildx builder" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Using existing buildx builder: $builderName" -ForegroundColor Cyan
        docker buildx use $builderName
    }

    # Inspect builder capabilities
    Write-Host "Builder capabilities:" -ForegroundColor Cyan
    docker buildx inspect --bootstrap
}

if ($Clean) {
    Clean-DockerSystem
}

# Setup buildx for multi-platform builds
if ($MultiPlatform) {
    Setup-DockerBuildx
}

# Check if build-extensions directory has any .vsix files
$buildExtensionsDir = Join-Path $PSScriptRoot "build-extensions"
$optionalExtensionsDir = Join-Path $PSScriptRoot "build-extensions-optional"

$requiredExtensions = @()
$optionalExtensions = @()

if (Test-Path $buildExtensionsDir) {
    $requiredExtensions = Get-ChildItem -Path $buildExtensionsDir -Filter "*.vsix" -ErrorAction SilentlyContinue
}

if (Test-Path $optionalExtensionsDir) {
    $optionalExtensions = Get-ChildItem -Path $optionalExtensionsDir -Filter "*.vsix" -ErrorAction SilentlyContinue
}

Write-Host "Build-time Extensions Summary:" -ForegroundColor Cyan
Write-Host "  Required extensions: $($requiredExtensions.Count)" -ForegroundColor Yellow
Write-Host "  Optional extensions: $($optionalExtensions.Count)" -ForegroundColor Yellow

if ($requiredExtensions.Count -gt 0) {
    Write-Host "  Required extensions:" -ForegroundColor White
    $requiredExtensions | ForEach-Object { Write-Host "    - $($_.Name)" -ForegroundColor Gray }
}

if ($optionalExtensions.Count -gt 0) {
    Write-Host "  Optional extensions:" -ForegroundColor White
    $optionalExtensions | ForEach-Object { Write-Host "    - $($_.Name)" -ForegroundColor Gray }
}

Write-Host ""

# Check if --push flag is provided or PUSH environment variable is set
$SHOULD_PUSH = ($Push -or ($env:PUSH -eq "true"))

if ($SHOULD_PUSH) {
    Write-Host "Building and pushing images..." -ForegroundColor Yellow
} else {
    Write-Host "Building images locally..." -ForegroundColor Yellow
}

# Build command with optional flags
if ($MultiPlatform) {
    # Use buildx for multi-platform builds
    $buildArgs = @("buildx", "build")

    # Add platform specification
    $buildArgs += @("--platform", $Platforms)

    # For multi-platform builds, we need to push or use --load (but --load doesn't work with multiple platforms)
    if ($SHOULD_PUSH) {
        $buildArgs += "--push"
    } else {
        # For local multi-platform builds, we'll build without loading to local Docker
        Write-Host "Note: Multi-platform builds cannot be loaded to local Docker." -ForegroundColor Yellow
        Write-Host "The image will be built but not available locally unless pushed to registry." -ForegroundColor Yellow
    }

    $buildArgs += @("-t", $FULL_IMAGE_NAME)
} else {
    # Use regular build for single platform
    $buildArgs = @("build", "-t", $FULL_IMAGE_NAME)
}

# Add labels similar to GitHub workflow
$buildArgs += @(
    "--label", "org.opencontainers.image.source=https://github.com/${Repository}",
    "--label", "org.opencontainers.image.description=VscUse VS Code Docker container with noVNC and recording capabilities",
    "--label", "org.opencontainers.image.licenses=MIT",
    "--label", "org.opencontainers.image.version=${IMAGE_VERSION}"
)

if ($NoCache) {
    $buildArgs += "--no-cache"
    Write-Host "Building with --no-cache flag..." -ForegroundColor Yellow
} else {
    Write-Host "Building with cache..." -ForegroundColor Yellow
}

$buildArgs += "."

# Build the image
Write-Host "Build command: docker $($buildArgs -join ' ')" -ForegroundColor Gray
& docker $buildArgs

if ($LASTEXITCODE -eq 0) {
    if ($SHOULD_PUSH) {
        if (-not $MultiPlatform) {
            Write-Host "Pushing images to registry..." -ForegroundColor Yellow
            Write-Host "Note: Make sure you're logged into GitHub Container Registry:" -ForegroundColor Yellow
            Write-Host "  docker login ghcr.io -u YOUR_USERNAME" -ForegroundColor Gray

            docker push $FULL_IMAGE_NAME
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Image built and pushed successfully!" -ForegroundColor Green
                Write-Host "Image available at: $FULL_IMAGE_NAME" -ForegroundColor Cyan
            } else {
                Write-Host "Push failed!" -ForegroundColor Red
                exit 1
            }
        } else {
            Write-Host "Multi-platform image built and pushed successfully!" -ForegroundColor Green
            Write-Host "Image available at: $FULL_IMAGE_NAME" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Image built successfully!" -ForegroundColor Green
        Write-Host "Image: $FULL_IMAGE_NAME" -ForegroundColor Cyan
        Write-Host "VS Code: Both stable and insider versions included" -ForegroundColor Yellow
        Write-Host "To push to GitHub Container Registry:" -ForegroundColor Yellow
        if ($MultiPlatform) {
            Write-Host "  .\build.ps1 -MultiPlatform -Push -Tag latest" -ForegroundColor Gray
        } else {
            Write-Host "  .\build.ps1 -Push -Tag latest" -ForegroundColor Gray
        }
    }
} else {
    Write-Host "Build failed!" -ForegroundColor Red
    exit 1
}
