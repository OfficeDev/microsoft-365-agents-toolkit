# Storage locations

Where every persistent thing lives.

## Per user (machine-global)

| Path | Contents | Lifetime |
|------|----------|----------|
| `~/.fx/account/m365TokenCache.bin` | M365 MSAL token cache (AES-256-GCM) | Until logout |
| `~/.fx/account/azureTokenCache.bin` | Azure MSAL token cache (AES-256-GCM) | Until logout |
| `~/.fx/account/keytarKey` (or OS keychain) | Symmetric encryption key | Until logout / reinstall |

## Per project (in repo)

| Path | Contents | Should commit? |
|------|----------|----------------|
| `appPackage/manifest.json` | Teams manifest | yes |
| `appPackage/declarativeAgent.json` | DA definition (DA templates) | yes |
| `appPackage/aiPlugin.json` | API plugin manifest (plugin templates) | yes |
| `appPackage/icons/*.png` | App icons | yes |
| `infra/azure.bicep` + parameters | Azure IaC | yes |
| `m365agents.yml` | Lifecycle actions | yes |
| `m365agents.local.yml` | Local debug lifecycle | yes |
| `env/.env.{envName}` | Non-secret env vars per environment | yes |
| `env/.env.{envName}.user` | Secret env vars (passwords, client secrets) | **no** (gitignored by default) |
| `src/` | App source | yes |
| `node_modules/`, `bin/`, `obj/` | Build artefacts | no (gitignored) |
| `.fx/` (v3 only) | Local engine state | no (gitignored) |

## In-process

| Structure | Lifetime |
|-----------|----------|
| `AtkContext` | Per operation invocation |
| `envMap` (in lifecycle executor) | Per `executeLifecycle` call |
| `correlationId` | Per user action (via `AsyncLocalStorage`) |
| `TemplateRegistry` / `DriverRegistry` | Process-global singletons |

## Telemetry sinks

| Sink | Where |
|------|-------|
| Application Insights | Microsoft-internal AI instance keys per package |
| VS Code "Microsoft 365 Agents" output channel | Per-session log |
| File logs | None by default; user can pipe `atk --debug` to a file |
