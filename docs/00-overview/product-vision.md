# Product vision

## What it is

**Microsoft 365 Agents Toolkit** is a developer toolset for **Visual Studio**, **Visual Studio Code**, and the **command line** that makes building agents and apps for Microsoft 365 Copilot, Microsoft Teams, and the broader Microsoft 365 ecosystem fast, simple, and delightful.

It is the successor to *Teams Toolkit* (the npm package family is still `@microsoft/teamsfx-*` and the binary is `atk`).

## Who it serves

| Persona | Primary surface | Primary outcome |
|---------|-----------------|-----------------|
| TypeScript / JavaScript developer | VS Code extension, `atk` CLI | Scaffold, debug, ship Copilot agents and Teams apps |
| .NET developer | Visual Studio extension | Same outcome on the .NET stack |
| Python developer | VS Code extension, `atk` CLI | Bot / custom-engine agent scaffolds in Python |
| CI / DevOps engineer | `atk` CLI (non-interactive) | Reproducible provision/deploy/publish in pipelines |
| IT admin | Output artifacts (manifest, app package) | Validate and distribute apps via tenant catalogs |

See [01-product/personas.md](../01-product/personas.md) for the full persona model.

## What it does end-to-end

```
   create ──▶ run/debug ──▶ provision ──▶ deploy ──▶ publish
   (scaffold)   (F5/Playground)   (Azure infra)   (code/zip)   (M365)
```

Every step is exposed across all three surfaces with a single shared engine underneath.

## Surfaces (at a glance)

| Surface | Package | UI primitive |
|---------|---------|-------------|
| VS Code | [`vscode-extension`](../05-engineering/package-reference/vscode-extension.md) | Tree views, webviews, Copilot Chat participant |
| Visual Studio | (separate VS extension repo, consumes `dotnet-sdk` and `function-extension`) | Project templates and tool windows |
| CLI v3 | [`cli`](../05-engineering/package-reference/cli.md) | `atk` binary (npm: `@microsoft/m365agentstoolkit-cli`) |
| CLI v4 (preview) | [`cli-next`](../05-engineering/package-reference/cli-next.md) | Same `atk` binary, Commander.js, registry-driven |

## Engine: two generations, one product

| | v3 (current) | v4 (next) |
|-|--------------|-----------|
| Core engine | [`packages/fx-core`](../05-engineering/package-reference/fx-core.md) | [`packages/core-next`](../05-engineering/package-reference/core-next.md) |
| API contracts | [`packages/api`](../05-engineering/package-reference/api.md) (v3) | `core-next/src/api/` (v4, inlined) |
| Context | `TOOLS` global singleton | injected `AtkContext` |
| Composition | `FxCore` class methods | `Operation` pipeline (`runOperation`, `defineOperation`) |
| Drivers | implicit modules | explicit `DriverRegistry` + `createDriver()` factory |
| Feature flag | — | `TEAMSFX_V4_CORE` |

The v3 engine ships today; v4 is feature-flagged preview. The product behaviour the user sees is the same; the internals differ. Every page in this site that touches the engine documents both lenses.

## Non-goals

- This site is **not** the public end-user docs. Those live at <https://aka.ms/teamsfx-docs>.
- It does **not** replace the README in each package. Per-package READMEs remain the front door for npm consumers.
- It does **not** mirror `.github/instructions/*`. Those files are the contract for Copilot; this site explains them in prose.

## North-star principles

1. **Day-1 platform support** — every new Microsoft 365 capability is reachable from day one.
2. **Familiar tools** — TypeScript, JavaScript, Python, C# with the editors and frameworks developers already use.
3. **Get started fast** — opinionated scaffolds with clear next steps.
4. **Development velocity** — F5 to a running agent; hot reload; secure tunnels; M365 Agents Playground.
5. **Ship with confidence** — declarative lifecycle (`m365agents.yml`), CI/CD templates, env-scoped secrets.
