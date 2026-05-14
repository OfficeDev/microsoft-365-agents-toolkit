# Publishing

How packages and the VS Code extension reach end users.

## npm packages

Lerna-orchestrated publish via `npm publish --workspace ...`.

| Package | npm name | Publish channel |
|---------|----------|-----------------|
| `api` | `@microsoft/teamsfx-api` | `latest` |
| `manifest` | `@microsoft/app-manifest` | `latest` |
| `fx-core` | `@microsoft/teamsfx-core` | `latest` (v3 line) |
| `cli` | `@microsoft/m365agentstoolkit-cli` | `latest` |
| `core-next` | `@microsoft/teamsfx-core-next` | preview (v4.0.0-x) |
| `cli-next` | `@microsoft/m365agentstoolkit-cli-next` | preview |
| `sdk`, `sdk-react`, `adaptivecards-tools-sdk`, `mcp-server`, `spec-parser` | various | `latest` |

## VS Code extension

`vsce package` + `vsce publish` to the VS Code Marketplace and Open VSX. Two channels:

- **Stable** — monthly cadence.
- **Insiders / Prerelease** — weekly cadence.

Daily builds also published to a GitHub release for opt-in install.

## Visual Studio extension

Published from a separate repository to the Visual Studio Marketplace. Co-versioned with the .NET SDK.

## NuGet

`packages/dotnet-sdk` publishes `Microsoft.TeamsFx` and friends to NuGet.org.

## Provenance

npm publishes use **provenance attestations** for supply-chain integrity. `--provenance` flag in the publish workflow.

## CI/CD for publishing

Publish workflows run in a GitHub Environment with **required reviewers**. Secrets (`NPM_TOKEN`, marketplace PATs) are pulled from that environment, never from the build job.

## Pre-publish checks

| Check | Tool |
|-------|------|
| Tests pass | CI gate |
| Lint clean | CI gate |
| `npm pack` dry-run | manual |
| Package size sanity | manual |
| Release notes drafted | `CHANGELOG.md` per package |

## Yanking / unpublishing

Avoid `npm unpublish` (breaks downstream installs). Use `npm deprecate` to mark a bad version with a guidance message, and publish a fixed patch.
