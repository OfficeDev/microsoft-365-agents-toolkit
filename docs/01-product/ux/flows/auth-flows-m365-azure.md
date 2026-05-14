# Auth flows — M365 and Azure

Two independent identities, two independent caches.

## M365 identity

Used for: Teams Developer Portal · Microsoft Graph · M365 PackageService (sideloading) · publishing to org catalog.

```
ensureM365Auth(ctx)
  ↓
M365TokenProvider.getAccessToken(scopes)
  ↓
silent attempt (read encrypted cache at ~/.fx/account/m365TokenCache.bin)
  ↓ (cache miss / expired)
interactive attempt (browser opens to login.microsoftonline.com)
  ↓
write tokens back to cache
  ↓
extract tenantId from JWT claims
```

Default scopes: M365 (`https://graph.microsoft.com/.default`), App Studio (TDP).

## Azure identity

Used for: ARM/Bicep deploy · Kudu zip deploy · subscription / resource group operations.

```
ensureAzureAuth(ctx)
  ↓
AzureAccountProvider.getIdentityCredentialAsync()
  ↓
DefaultAzureCredential chain (CI: env vars → service principal; local: MSAL interactive)
  ↓
write tokens to ~/.fx/account/azureTokenCache.bin
  ↓
ensureSubscription(ctx, envMap)
  - 0 subs → fail with help link
  - 1 sub → auto-select
  - >1 subs → prompt
  ↓
ensureResourceGroup(ctx, envMap, ...)
  - prompt with default rg-{safeName}{suffix}-{env}
  - existing RG: confirm reuse
  - new RG: confirm location
```

## CI auth

Detected via env vars: `CI`, `TF_BUILD`, `GITHUB_ACTIONS`, or `CI_ENABLED=true`.

| CI scenario | Auth |
|-------------|------|
| GitHub Actions with workload identity | OIDC token exchange — `azure/login` action upstream |
| ADO with workload identity | Same |
| Service principal (legacy) | `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` env vars |
| M365 in CI | Service principal with admin consent — typically use a dedicated automation account |

CI flows go through `azureLoginCI.ts` (v4) instead of the interactive `azureLogin.ts`.

## Cache compatibility

`~/.fx/account/` is **shared between v3 and v4**. Logging in via `cli` makes `cli-next` see the same tokens, and vice versa. Encryption: AES-256-GCM. If `keytar` is installed and an OS keychain is reachable, the symmetric key is stored there; otherwise it derives from a machine-stable secret.

## Sign-out

`atk account logout m365` and `atk account logout azure` clear the respective cache files. The VS Code extension exposes equivalent commands in the Account tree view.
