# 5 — Building blocks

C4 levels 1 and 2 of the toolkit.

## Level 1 — System

The toolkit is a **single system** decomposed into surface adapters and an engine.

```mermaid
flowchart TB
    subgraph Surfaces
      VSC[vscode-extension]
      VS[Visual Studio extension]
      CLI3[cli v3]
      CLI4[cli-next v4]
      SRV[server JSON-RPC]
    end

    subgraph "Engine v3"
      API[api]
      MAN[manifest]
      FXC[fx-core]
    end

    subgraph "Engine v4"
      CN[core-next]
    end

    subgraph Templates
      TPL[templates/]
    end

    VSC --> FXC
    VS --> FXC
    CLI3 --> FXC
    SRV --> FXC
    CLI4 --> CN

    FXC --> API
    FXC --> MAN
    CN -.merges.-> API
    CN -.merges.-> MAN

    FXC --> TPL
    CN --> TPL
```

## Level 2 — v3 engine

```mermaid
flowchart TB
    subgraph fx-core
      Coord[Coordinator]
      Gen[Generators]
      Drv[Drivers]
      QM[Question Model]
      Env[EnvironmentManager]
      FF[FeatureFlagManager]
      Sec[secretmasker SVM+BloomFilter]
    end
    Coord --> Gen
    Coord --> Drv
    Coord --> QM
    Coord --> Env
    Coord --> FF
    Drv --> Sec
```

## Level 2 — v4 engine (`core-next`)

```mermaid
flowchart TB
    subgraph core-next
      OP[Operation pipeline]
      LC[Lifecycle engine]
      DR[DriverRegistry]
      TR[TemplateRegistry]
      QM[Question Model]
      DA[Declarative Agent module]
      Env[Environment manager]
      Tel[Telemetry helpers]
      Sec[secretMasker keywords]
      FF[FeatureFlagRegistry]
      L10n[Localizer]
      Http[HTTP client]
      Cli[Service clients]
    end

    OP --> LC
    LC --> DR
    OP --> TR
    OP --> QM
    OP --> DA
    OP --> Env
    OP --> Tel
    OP --> Cli
    Cli --> Http
    Cli --> Sec
    DA --> DR
```

## Per-package roles

See [05-engineering/package-reference/README.md](../05-engineering/package-reference/README.md) for per-package detail.

| Package | Role | Bundler |
|---------|------|---------|
| [`api`](../05-engineering/package-reference/api.md) | v3 public contracts | tsc |
| [`manifest`](../05-engineering/package-reference/manifest.md) | Typed manifest wrapper | tsc |
| [`fx-core`](../05-engineering/package-reference/fx-core.md) | v3 engine | webpack |
| [`vscode-extension`](../05-engineering/package-reference/vscode-extension.md) | VS Code adapter + UI | esbuild |
| [`cli`](../05-engineering/package-reference/cli.md) | v3 CLI | webpack |
| [`core-next`](../05-engineering/package-reference/core-next.md) | v4 engine | tsc |
| [`cli-next`](../05-engineering/package-reference/cli-next.md) | v4 CLI | esbuild |
| [`server`](../05-engineering/package-reference/server.md) | JSON-RPC bridge to fx-core | webpack |
| [`sdk`](../05-engineering/package-reference/sdk.md) | App-side SDK (auth helpers) | rollup |
| [`sdk-react`](../05-engineering/package-reference/sdk-react.md) | React hooks for SDK | rollup |
| [`mcp-server`](../05-engineering/package-reference/mcp-server.md) | Model Context Protocol server | webpack |
| [`spec-parser`](../05-engineering/package-reference/spec-parser.md) | OpenAPI parser/validator/filter | rollup |
| [`templates`](../05-engineering/package-reference/templates.md) | Project scaffolds | npm scripts |
