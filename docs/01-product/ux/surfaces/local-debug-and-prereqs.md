# Local debug and prerequisites (v3)

> **v4 design status.** This page is an allowed input to v4 design. It captures the v3 user-facing surfaces around local debugging — what the user sees, what gets validated, what gets started — without committing to v3's internal subsystem boundaries.

When a user presses F5 in VS Code or runs `atk preview` / `atk doctor`, the toolkit takes them through several *prerequisite* + *runtime* steps before the app is ready to interact with. This page enumerates each of those surfaces.

## Dependency checker

Run as part of `atk teamsapp doctor` and during local debug startup. Source: [`packages/fx-core/src/component/deps-checker/`](../../../packages/fx-core/src/component/deps-checker/).

| `DepsType` | What it checks | Probe | Help link surfaced on failure |
|-----------|----------------|-------|-------------------------------|
| `LtsNode` | Required LTS Node.js version | `node --version` | "Install supported Node.js" |
| `ProjectNode` | Per-project Node.js version (when stricter than LTS) | `node --version` against project requirement | Same |
| `Dotnet` | .NET SDK (for C# projects) | `dotnet --version` | "Install .NET SDK" |
| `FuncCoreTools` | Azure Functions Core Tools (`func`) | `func --version` | "Install Azure Functions Core Tools" |
| `TestTool` | M365 Agents Playground npm package | `npx m365agentsplayground --version` (or installed binary lookup) | "Install M365 Agents Playground" |
| `VxTestApp` | (Internal test prerequisite) | — | — |

Each check returns `DependencyStatus { isInstalled, installVersion, supportedVersions, helpLink }`. The CLI renders results as a table; VS Code renders them in a tree view with per-row install buttons.

User-facing failure modes:

- `NodejsNotLtsError` — Node version below the required LTS line.
- `NodejsNotFoundError` — `node` not on PATH.
- `NodejsNotRecommendedError` — Node above LTS (warning, not blocking).
- `InstallNodeJSError` — auto-install failed.
- `DepsCheckerError` — generic check failure (typically with the underlying error chained).

## Port and certificate checks

Run during local debug startup. Source: [`packages/fx-core/src/component/local/`](../../../packages/fx-core/src/component/local/).

| Check | What it does | Failure name |
|-------|-------------|--------------|
| Port availability | Probes default debug ports (e.g. 3007 for tab, 3978 for bot) | `PortsConflictError` (lists conflicting ports + processes) |
| SSL certificate | Verifies the local dev cert exists and isn't expired; auto-generates if missing; trusts in OS keychain | Certificate-trust dialog if first-run, or warning notification if expired |
| `npmLogHelper` | Tails npm install/build log for "ready" signal | Surfaces npm errors verbatim if the dev server fails to start |

The user sees progress in the VS Code "Microsoft 365 Agents Toolkit" output channel, plus per-step status in the debug task output.

## Dev tunnels

Used to expose a local bot or backend at a public HTTPS URL so Microsoft 365 / Teams can reach it. Source: VS Code task type `dev-tunnel` defined in `packages/vscode-extension/package.json` `contributes.taskDefinitions`; runtime in `packages/vscode-extension/src/debug/`.

### Task type

```jsonc
{
  "type": "dev-tunnel",
  "ports": [
    {
      "portNumber": 3978,
      "protocol": "http" | "https",
      "access": "public" | "private",
      "writeToEnvironmentFile": {
        "endpoint": "BOT_ENDPOINT",
        "domain": "BOT_DOMAIN"
      }
    }
  ]
}
```

### User flow

1. F5 → the `dev-tunnel` task starts.
2. Toolkit calls `@microsoft/dev-tunnels-management` to allocate a tunnel.
3. Tunnel URL (`https://{domain}.devtunnels.ms`) is written to env file via `writeToEnvironmentFile`.
4. Subsequent local-debug actions consume `BOT_ENDPOINT` to register the bot channel and update manifest `validDomains`.
5. On debug stop → tunnel is reclaimed.

### Failure modes

- Tunnel quota exceeded (max concurrent tunnels per account) → error with reclaim guidance.
- Authentication failed → re-auth via `atk auth login m365`.
- Network blocked (corporate proxy) → fallback to `ngrok` (separate task type) if installed.

### Feature flag

`TEAMSFX_DEV_TUNNEL_TEST` — gates experimental tunnel behaviours (testing only).

## M365 Agents Playground (Test Tool)

A local emulator for bots and message extensions. The user can debug without sideloading into Teams.

Auto-detected when the project's environment is named `testtool`. Source: VS Code command `fx-extension.debugInTestToolWithIcon`, helper `isTestToolEnabledProject()`.

### User flow

1. Project scaffolded with `m365agents.local.testtool.yml` + an env named `testtool` → toolkit detects support.
2. Tree view shows "Debug in Microsoft 365 Agents Playground" button on the `testtool` env item.
3. Click → toolkit launches `m365agentsplayground` (CLI binary) with the project's local endpoint pointed at it.
4. User chats with the bot inside the Playground UI; Adaptive Cards render natively.
5. Telemetry: `MessageDebugInTestTool` event when launched.

### Failure modes

- Test Tool not installed → DepsChecker prompts install via npm.
- Local bot endpoint not reachable → "Failed to connect" with link to local-debug troubleshooting.

## Local lifecycle (m365agents.local.yml)

The local lifecycle is a separate YAML file from `m365agents.yml`. It typically runs:

- Bot AAD app creation (`botAadApp/create`) in the local env (single tenant).
- (Optional) ARM deploy for any local Azure infra (rare; usually Test Tool replaces this).
- `script` actions to start dev tunnel + start the bot process.
- `teamsApp/zipAppPackage` + `teamsApp/extendToM365` to sideload the app for local testing.

The user does not invoke these explicitly — they run automatically as part of F5 / `atk preview --env local`.

## `atk preview` runtime flow

CLI command for previewing a local app in a chosen M365 host (Teams web, Teams desktop, Outlook, M365 Copilot).

### Full options (15+)

| Option | Purpose |
|--------|---------|
| `--env <name>` | Which environment to preview (typically `local`) |
| `--manifest-file <path>` | Override manifest path |
| `-c, --run-command <cmd>` | Custom command to start the local server (default: auto-detect `npm run dev` / `func start` / `python app.py`) |
| `-p, --running-pattern <regex>` | Regex matched against the `--run-command`'s stdout to detect "ready" |
| `-o, --open-only` | Skip starting the local service; just open the M365 client |
| `-b, --browser <chrome\|edge\|default>` | Which browser to launch |
| `-ba, --browser-arg <arg>` | Additional arg to pass to the browser (e.g. `--guest`) |
| `-ep, --exec-path <path>` | Path to `func` / other tooling executable |
| `--m365-host <teams\|outlook\|m365>` | Which M365 surface to open the app in |
| `--desktop` | Launch Teams desktop instead of Teams web |
| `--ignore-env-file` | Skip loading `.env.*.user` |
| (others) | Less commonly used; see `atk preview --help` |

### Telemetry properties

- `PreviewType` — `local` or `remote`
- `PreviewHub` — `teams`, `outlook`, `m365`
- `PreviewBrowser` — chosen browser
- `PreviewAppId` — Teams App ID

## Problem matchers (for VS Code task output)

Contributed via `contributes.problemMatchers` in `package.json`. Each matches output patterns from a specific subsystem so VS Code can render errors inline:

| Name | Subsystem | Matches |
|------|-----------|---------|
| `teamsfx-frontend-watch` | Frontend (React tab) | webpack-style errors |
| `teamsfx-backend-watch` | Backend (Functions / App Service) | TypeScript / func tooling errors |
| `teamsfx-bot-watch` | Bot dev server | Microsoft Agents SDK errors |
| `teamsfx-auth-watch` | SimpleAuth (legacy SSO proxy) | Auth proxy errors |
| `teamsfx-tunnel` | Dev tunnel | Tunnel allocation / disconnect errors |
| `teamsfx-ngrok` | ngrok fallback | ngrok errors |
| `teamsfx-local-tunnel` | Local tunnel server | Tunnel-server errors |

## VS Code "Microsoft 365 Agents Toolkit" output channel

All extension log output flows through one named channel:

- Channel name: `Microsoft 365 Agents Toolkit`.
- Log levels controlled by setting `M365AgentsToolkit.logLevel` (`Info` / `Verbose` / `Debug`).
- Format: `[TIMESTAMP] [LEVEL] [COMPONENT] message`.
- Auto-shown on first error (configurable).
- Custom language `teamsfx-toolkit-output` + grammar contributed to provide syntax highlighting in the channel.

There is **no separate telemetry-trace channel** — telemetry events are not echoed to the user; debug them via `TEAMSFX_TELEMETRY_TRACE=true` env var which routes events to the same output channel.

## See also

- [vscode-extension-commands.md](vscode-extension-commands.md) — the commands that trigger these flows
- [cli-v3-command-reference.md](cli-v3-command-reference.md) — `atk doctor`, `atk preview`
- [tenant-and-sovereign-cloud.md](../flows/tenant-and-sovereign-cloud.md) — auth side of local debug
- [../../01-product/v3-feature-inventory.md](../../01-product/v3-feature-inventory.md) — Test Tool / dev tunnels in the integration surface table
