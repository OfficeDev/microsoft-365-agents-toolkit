# Local development environment

What the toolkit needs and provides on the developer's local machine.

## Prerequisites

| Tool | Minimum | Notes |
|------|---------|-------|
| Node.js | 18.x or 20.x | esbuild target; MSAL native plugins |
| Python | 3.10+ | For Python templates only |
| .NET | 8.0+ | For C# templates only |
| Visual Studio Code | Stable or Insiders | For the VS Code extension |
| Visual Studio | 2022 17.10+ | For the VS extension |

## Tunnels

Local debugging of bots and message extensions requires a public HTTPS endpoint reachable by the M365 service. The toolkit launches **dev tunnels** (the Microsoft tunneling service used by VS) automatically as part of the `m365agents.local.yml` lifecycle.

Behind the scenes:

- VS Code extension calls into the dev-tunnel CLI / API.
- Tunnel URL becomes `BOT_ENDPOINT` and is written into `env/.env.local`.
- The Bot Channel registration (or its local equivalent) is updated to point at the tunnel.

## M365 Agents Playground

`@microsoft/m365agentsplayground` (formerly Teams App Test Tool) emulates Teams locally for bots and message extensions. F5 from VS Code launches the agent against the Playground UI rather than real Teams. Faster inner loop, no sideload required.

## Token cache

`~/.fx/account/` stores MSAL token caches:

| File | Contents |
|------|----------|
| `m365TokenCache.bin` | M365 access + refresh tokens (encrypted) |
| `azureTokenCache.bin` | Azure tokens (encrypted) |

Both are AES-256-GCM-encrypted. If `keytar` is installed and the OS keychain is accessible, the encryption key is stored there; otherwise it falls back to a machine-derived key.

This cache is **shared between v3 and v4** — `cli-next` reads / writes the same files as `cli`.

## Project layout after scaffold

```
my-agent/
├── appPackage/                 # manifest + icons
├── env/
│   ├── .env.dev
│   └── .env.local
├── infra/
│   ├── azure.bicep
│   └── azure.parameters.json
├── m365agents.yml
├── m365agents.local.yml
├── src/                        # app code
└── .vscode/launch.json         # F5 wired to local lifecycle
```

`.fx/` (legacy v3 state) is created on first command if needed; v4 does not require it.
