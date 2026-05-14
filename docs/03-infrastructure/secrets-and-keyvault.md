# Secrets and Key Vault

## How secrets flow today

| Origin | Lifetime | Storage |
|--------|----------|---------|
| MSAL access tokens | Session + silent refresh | `~/.fx/account/` (AES-256-GCM, optional keytar) |
| Bot password (`botAadApp/create` output) | Long-lived | Project `env/.env.{envName}.user` (gitignored) |
| AAD app client secret (`aadApp/create` output) | Long-lived | Project `env/.env.{envName}.user` |
| API key (`apiKey/register`) | Long-lived | TDP store (server-side); local copy in `.env.{envName}.user` |
| OAuth client secret (`oauth/register`) | Long-lived | TDP store; local copy in `.env.{envName}.user` |

`env/.env.{envName}.user` is the **shipped pattern** for secret storage — it is gitignored by default in every scaffolded project's `.gitignore`.

## Promoting to Key Vault

For production, users typically:

1. Add a Key Vault to their Bicep (`Microsoft.KeyVault/vaults`).
2. Switch `appSettings` from `${{BOT_PASSWORD}}` to `@Microsoft.KeyVault(SecretUri=...)`.
3. Grant the App Service / Function App's managed identity `get` permission on the vault.

The toolkit does not automate this — see the **azure-keyvault-expiration-audit** skill for ongoing maintenance recommendations.

## Secret masking in logs and telemetry

| Layer | Masker |
|-------|--------|
| v3 | `common/secretmasker/` — keyword + SVM + BloomFilter |
| v4 | `core-next/src/secretMasker/` — keyword regex (100+ suffixes) |

Both apply to log lines, error messages, and telemetry properties before they leave the process. URLs in HTTP client logs are also masked (query-string secrets stripped).

See [02-architecture/08-crosscutting-concepts.md](../02-architecture/08-crosscutting-concepts.md) §"Security".
