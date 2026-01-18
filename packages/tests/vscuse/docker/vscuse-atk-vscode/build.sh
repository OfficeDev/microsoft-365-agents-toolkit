#!/bin/bash

# Add parameters for enhanced build options
NO_CACHE=false
CLEAN=false
PUSH=false
MULTI_PLATFORM=false
TAG="latest"
REPOSITORY="1openwindow"
PLATFORMS="linux/arm64,linux/amd64"

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        --no-cache)
            NO_CACHE=true
            shift
            ;;
        --clean)
            CLEAN=true
            shift
            ;;
        --push)
            PUSH=true
            shift
            ;;
        --multi-platform)
            MULTI_PLATFORM=true
            shift
            ;;
        --tag)
            TAG="$2"
            shift 2
            ;;
        --repository)
            REPOSITORY="$2"
            shift 2
            ;;
        --platforms)
            PLATFORMS="$2"
            shift 2
            ;;
        *)
            echo "Unknown option: $1"
            echo "Usage: $0 [--no-cache] [--clean] [--push] [--multi-platform] [--tag TAG] [--repository REPO] [--platforms PLATFORMS]"
            exit 1
            ;;
    esac
done

# Bash version of build script
IMAGE_NAME="vscuse-atk-vscode"
IMAGE_VERSION="$TAG"

# Always use GitHub Container Registry
REGISTRY="ghcr.io/${REPOSITORY}"
FULL_IMAGE_NAME="${REGISTRY}/${IMAGE_NAME}:${IMAGE_VERSION}"

echo "=== VscUse ATK VSCode Docker Build Script ==="
echo "Registry: GitHub Container Registry (ghcr.io)"
echo "Image: $FULL_IMAGE_NAME"
echo "VS Code: Both stable and insider versions included"

if [ "$MULTI_PLATFORM" = true ]; then
    echo "Multi-platform build enabled for: $PLATFORMS"
fi

# Function to clean Docker system
clean_docker_system() {
    echo "Cleaning Docker system..."
    echo "Current disk usage:"
    docker system df

    echo
    echo "Removing old image..."
    docker rmi "$FULL_IMAGE_NAME" -f 2>/dev/null || true

    echo "Pruning Docker system..."
    docker system prune -a --volumes -f

    echo "Post-cleanup disk usage:"
    docker system df
}

# Function to setup Docker Buildx for multi-platform builds
setup_docker_buildx() {
    echo "Setting up Docker Buildx for multi-platform builds..."

    # Check if buildx is available
    if ! docker buildx version >/dev/null 2>&1; then
        echo "❌ Docker Buildx is not available. Please update Docker Desktop."
        exit 1
    fi

    # Create or use existing builder
    local builder_name="vscuse-multiarch-builder"

    # Check if builder exists
    if ! docker buildx ls | grep -q "$builder_name"; then
        echo "Creating new buildx builder: $builder_name"
        docker buildx create --name "$builder_name" --use --bootstrap
        if [ $? -ne 0 ]; then
            echo "❌ Failed to create buildx builder"
            exit 1
        fi
    else
        echo "Using existing buildx builder: $builder_name"
        docker buildx use "$builder_name"
    fi

    # Inspect builder capabilities
    echo "Builder capabilities:"
    docker buildx inspect --bootstrap
}

if [ "$CLEAN" = true ]; then
    clean_docker_system
fi

# Setup buildx for multi-platform builds
if [ "$MULTI_PLATFORM" = true ]; then
    setup_docker_buildx
fi

# Check for build-time extensions
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REQUIRED_EXTENSIONS_DIR="$SCRIPT_DIR/build-extensions"
OPTIONAL_EXTENSIONS_DIR="$SCRIPT_DIR/build-extensions-optional"

REQUIRED_COUNT=0
OPTIONAL_COUNT=0

if [ -d "$REQUIRED_EXTENSIONS_DIR" ]; then
    REQUIRED_COUNT=$(find "$REQUIRED_EXTENSIONS_DIR" -name "*.vsix" 2>/dev/null | wc -l)
fi

if [ -d "$OPTIONAL_EXTENSIONS_DIR" ]; then
    OPTIONAL_COUNT=$(find "$OPTIONAL_EXTENSIONS_DIR" -name "*.vsix" 2>/dev/null | wc -l)
fi

echo "Build-time Extensions Summary:"
echo "  Required extensions: $REQUIRED_COUNT"
echo "  Optional extensions: $OPTIONAL_COUNT"

if [ $REQUIRED_COUNT -gt 0 ]; then
    echo "  Required extensions:"
    find "$REQUIRED_EXTENSIONS_DIR" -name "*.vsix" 2>/dev/null -exec basename {} \; | sed 's/^/    - /'
fi

if [ $OPTIONAL_COUNT -gt 0 ]; then
    echo "  Optional extensions:"
    find "$OPTIONAL_EXTENSIONS_DIR" -name "*.vsix" 2>/dev/null -exec basename {} \; | sed 's/^/    - /'
fi

echo ""

# Check if --push flag is provided or PUSH environment variable is set
if [ "$PUSH" = true ] || [ "$PUSH_ENV" = "true" ]; then
    echo "Building and pushing images..."
else
    echo "Building images locally..."
fi

# Build command with optional flags
if [ "$MULTI_PLATFORM" = true ]; then
    # Use buildx for multi-platform builds
    BUILD_ARGS=("buildx" "build")

    # Add platform specification
    BUILD_ARGS+=("--platform" "$PLATFORMS")

    # For multi-platform builds, we need to push or use --load (but --load doesn't work with multiple platforms)
    if [ "$PUSH" = true ] || [ "$PUSH_ENV" = "true" ]; then
        BUILD_ARGS+=("--push")
    else
        # For local multi-platform builds, we'll build without loading to local Docker
        echo "Note: Multi-platform builds cannot be loaded to local Docker."
        echo "The image will be built but not available locally unless pushed to registry."
    fi

    BUILD_ARGS+=("-t" "$FULL_IMAGE_NAME")
else
    # Use regular build for single platform
    BUILD_ARGS=("build" "--platform" "linux/arm64" "-t" "$FULL_IMAGE_NAME")
fi

# Add labels similar to GitHub workflow
BUILD_ARGS+=(
    "--label" "org.opencontainers.image.source=https://github.com/${REPOSITORY}"
    "--label" "org.opencontainers.image.description=VscUse VS Code Docker container with noVNC and recording capabilities"
    "--label" "org.opencontainers.image.licenses=MIT"
    "--label" "org.opencontainers.image.version=${IMAGE_VERSION}"
)

if [ "$NO_CACHE" = true ]; then
    BUILD_ARGS+=("--no-cache")
    echo "Building with --no-cache flag..."
else
    echo "Building with cache..."
fi

BUILD_ARGS+=(".")

# Build the image
echo "Build command: docker ${BUILD_ARGS[*]}"
docker "${BUILD_ARGS[@]}"

if [ $? -eq 0 ]; then
    if [ "$PUSH" = true ] || [ "$PUSH_ENV" = "true" ]; then
        if [ "$MULTI_PLATFORM" = false ]; then
            echo "Pushing images to registry..."
            echo "Note: Make sure you're logged into GitHub Container Registry:"
            echo "  docker login ghcr.io -u YOUR_USERNAME"

            docker push "$FULL_IMAGE_NAME"
            if [ $? -eq 0 ]; then
                echo "Image built and pushed successfully!"
                echo "Image available at: $FULL_IMAGE_NAME"
            else
                echo "Push failed!"
                exit 1
            fi
        else
            echo "Multi-platform image built and pushed successfully!"
            echo "Image available at: $FULL_IMAGE_NAME"
        fi
    else
        echo "Image built successfully!"
        echo "Image: $FULL_IMAGE_NAME"
        echo "VS Code versions: Both stable and insider included"
        echo "To push to GitHub Container Registry:"
        if [ "$MULTI_PLATFORM" = true ]; then
            echo "  ./build.sh --multi-platform --push --tag latest"
        else
            echo "  ./build.sh --push --tag latest"
        fi
    fi
else
    echo "Build failed!"
    exit 1
fi
