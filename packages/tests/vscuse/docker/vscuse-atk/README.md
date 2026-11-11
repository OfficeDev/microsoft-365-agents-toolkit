# VscUse ATK Profile

This directory contains the **ATK (Microsoft 365 Agent Toolkit)** profile for VscUse.

## What's Included

This profile extends the base image with ATK-specific tools:

- **.NET SDK 8.0** - Required for .NET-based agents
- **Azure Functions Core Tools v4.4.0** - Local functions development
- **Microsoft 365 Agents Toolkit CLI** - M365 agent scaffolding
- **Dev Tunnel CLI** - Secure tunneling for local development
- **Microsoft Edge** - Browser for M365 testing
- **M365 Template Directory** - Pre-configured devTools structure
- **ATK Extension** - Microsoft 365 Agent Toolkit VS Code extension

## Building the Image

### Using docker build directly

```bash
docker build -t ghcr.io/<your-github-username>/vscuse-atk:latest .
```

### Overriding the Node.js version

The image uses the Node.js provided by the base image by default. To bake a different
Node.js version, pass the `NODE_VERSION` build arg. It accepts either a major line
(e.g. `22`, resolved to the latest matching release) or a full version (e.g. `22.11.0`):

```bash
docker build --build-arg NODE_VERSION=22 -t ghcr.io/<your-github-username>/vscuse-atk:latest .
```

The same override is exposed as the `node_version` input on the
`Build VscUse ATK Docker Image` workflow.

### Selecting an npm registry for local builds

`NPM_CONFIG_REGISTRY` is an optional build argument for npm installations during
image creation. Its default is `https://registry.npmjs.org/`. For environments
that require the Microsoft package proxy, pass it explicitly without disabling
TLS verification:

```bash
docker build --build-arg NODE_VERSION=22 --build-arg NPM_CONFIG_REGISTRY=https://packagefeedproxy.microsoft.io/npm/ -t vscuse-atk-dev:local .
```

This build argument does not configure npm inside a later test container; use
the run config's `docker.environment` for that.

## Extending from Base Image

The Dockerfile and CI workflow default to the same immutable Linux AMD64 base:

```dockerfile
ARG BASE_IMAGE=ghcr.io/microsoft/vscuse-base@sha256:2db458897652e039d87f3bd28928c74003be4224d3ee49484b245f63e48f7977
FROM ${BASE_IMAGE}
```

This is the Chrome 155.0.8059.39 baseline used by the `20261013` ATK image. Dev CD
and PR smoke builds use this same base, while installing their own VSIX artifacts.
The workflow supplies the pinned build argument even when checking out a PR's
older Dockerfile. Local builds use the Dockerfile's matching default.

Keep this digest fixed until vulnerabilities in the pinned image require an
update. Do not refresh it automatically from `latest` or on a periodic schedule.
To remediate vulnerabilities, validate a replacement image's Chrome startup and
login behavior, then update both the Dockerfile's `BASE_IMAGE` default and the
workflow's `PINNED_BASE_IMAGE`, run the image regression tests, and rebuild the dev
and PR images. Existing published tags are not changed by updating the pin.

The optional workflow `base_image_tag` input is an explicit override for candidate
validation; normal dev and PR builds leave it empty. Local candidate builds can
override `BASE_IMAGE` with `--build-arg`. Pinning the base does not pin the
ATK-specific packages installed in later layers or add vulnerability scanning.

## Configuration

### Using with vscuse CLI

Update `config.yaml`:

```yaml
docker:
  image_name: "ghcr.io/<your-github-username>/vscuse-atk:latest"
```

The image entrypoint mirrors a declared `TEAMSFX_CEA_ENABLED=true|false` test precondition into
`M365AgentsToolkit.enableLaunchAgentForTeamsInCopilot` before VS Code starts. Generate a
plan-scoped run config with `vscode-test-cases/engine/prepare-vscuse-config.cjs`; do not add the
setting to the shared image settings.

On AMD64, Chrome desktop actions and the `vscode` user's HTTP/HTTPS associations
use the base image's `chrome-wrapper`, including its `--no-first-run` flag.
Setting `BROWSER` alone is insufficient because `xdg-open` tries desktop URL
associations first; those associations must not launch the unwrapped Chrome binary.

The test image also adds `--test-type` to the Chrome wrapper to suppress startup
warning bars that shift the page during UI tests. This keeps the existing container
flags unchanged; it does not enable sandboxing or make those flags safe for regular
browsing. Rebuild the image to consume these changes; updating a test plan does not
modify an existing image's browser startup configuration.

## Differences from Base Image

| Feature         | Base Image | ATK Profile      |
| --------------- | ---------- | ---------------- |
| Ubuntu 22.04    | ✅         | ✅               |
| Python 3.12     | ✅         | ✅               |
| Node.js 20 LTS  | ✅         | ✅               |
| VS Code         | ✅         | ✅               |
| Chrome/Chromium | ✅         | ✅               |
| Microsoft Edge  | ❌         | ✅               |
| .NET SDK 8.0    | ❌         | ✅               |
| Azure Functions | ❌         | ✅ v4.4.0        |
| M365 CLI        | ❌         | ✅ Alpha         |
| Dev Tunnel      | ❌         | ✅ Latest        |
| ATK Extension   | ❌         | ✅ Pre-installed |

## Directory Structure

```
vscuse-atk/
├── Dockerfile                    # Profile build definition
├── build-extensions/             # ATK-specific extensions
├── sync-vscode-feature-flags.cjs # Test-only setting synchronization
└── vscuse-atk-entrypoint.sh      # Pre-VS Code startup hook
```
