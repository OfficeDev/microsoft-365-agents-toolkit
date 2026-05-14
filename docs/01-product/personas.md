# Personas

Personas are the **anchors** for every UX decision and PRD scenario. They are intentionally small in number; sub-segments belong in scenarios.

## P1 — TypeScript / JavaScript developer

| | |
|-|-|
| Primary surface | VS Code extension; `atk` CLI |
| Stack | Node.js, npm, Microsoft Agents SDK / Teams AI Library |
| Top jobs | Scaffold a Copilot agent or Teams bot · F5 to running · provision to Azure · publish to tenant |
| Pains | Manifest schema confusion · token-cache friction · provisioning errors with no clear next step |
| Wins from us | Opinionated templates · in-editor errors with help links · single-command provision |

## P2 — .NET developer

| | |
|-|-|
| Primary surface | Visual Studio extension; `dotnet` CLI |
| Stack | .NET 8+, ASP.NET, Microsoft Agents SDK |
| Top jobs | Scaffold C# agent · debug with VS · publish via VS |
| Pains | Project layout differs from Web SDK templates · separate auth flows |
| Wins from us | Project templates that fit Visual Studio patterns · shared engine behaviour with VS Code |

## P3 — Python developer

| | |
|-|-|
| Primary surface | VS Code extension; `atk` CLI |
| Stack | Python 3.10+, Microsoft Agents SDK for Python |
| Top jobs | Scaffold bot or custom-engine agent in Python · run locally · deploy to Azure Functions / App Service |
| Pains | Smaller template surface than TS · Python tooling vs Node tooling friction in CI |
| Wins from us | First-class Python templates; see [capabilities-matrix.md](capabilities-matrix.md) for current list and coverage status. Same lifecycle commands as TS. |

## P4 — CI / DevOps engineer

| | |
|-|-|
| Primary surface | `atk` CLI in headless mode (`CI_ENABLED=true`) |
| Stack | GitHub Actions, Azure DevOps |
| Top jobs | Reproducible provision/deploy/publish in pipelines · service-principal auth · per-environment secrets |
| Pains | Interactive prompts breaking automation · unclear which env vars must be set |
| Wins from us | Pure non-interactive mode · service-principal flow (`azureLoginCI`) · shipped pipeline templates |

## P5 — IT admin

| | |
|-|-|
| Primary surface | Output artefacts (manifest zip, validation reports) |
| Top jobs | Validate apps before tenant rollout · publish to org catalog · monitor adoption |
| Pains | Late-discovery of manifest issues · unclear what permissions an app requires |
| Wins from us | Pre-publish validation drivers (`teamsApp/validateManifest`, `teamsApp/validateAppPackage`) · clear permission summaries |

## P6 — Open-source contributor

| | |
|-|-|
| Primary surface | This monorepo |
| Top jobs | Add a template / driver / CLI command · fix a bug · run lint/test locally |
| Pains | Two engines (v3, v4) · large generated code surface · long build times |
| Wins from us | [Codebase instructions](../../.github/instructions/codebase.instructions.md) · [skills](../../.github/skills/) · [contributor playbooks](../07-contributing/README.md) |
