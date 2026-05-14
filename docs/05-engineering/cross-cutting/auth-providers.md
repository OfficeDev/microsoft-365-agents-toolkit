# Auth providers

Source: [`packages/cli-next/src/auth/`](../../../packages/cli-next/src/auth/) (v4) and [`packages/cli/src/commonlib/`](../../../packages/cli/src/commonlib/) (v3).

## Two identities

| Identity | Used for |
|----------|----------|
| **M365 (Microsoft 365)** | Teams Developer Portal · Microsoft Graph · M365 PackageService · org catalog publish |
| **Azure** | Subscription / RG operations · ARM/Bicep deploy · Kudu zip deploy |

Each has its own MSAL flow, token cache, and provider class. They are independent — a user can be signed into one but not the other.

## Provider interfaces

Defined in `core-next/src/api/auth/`:

```typescript
interface M365TokenProvider {
  getAccessToken(scopes: string[] | string): Promise<Result<string, AtkError>>;
  getJsonObject(scopes?: string[]): Promise<Result<Record<string, unknown>, AtkError>>;
  signout(): Promise<Result<void, AtkError>>;
  setStatusChangeMap(name, statuses, callback): Promise<Result<boolean, AtkError>>;
  removeStatusChangeMap(name): Promise<Result<boolean, AtkError>>;
}

interface AzureAccountProvider {
  getIdentityCredentialAsync(): Promise<TokenCredential | undefined>;
  signout(): Promise<boolean>;
  listSubscriptions(): Promise<SubscriptionInfo[]>;
  setSubscription(subscriptionId: string): Promise<void>;
}
```

## `createTokenProvider()` factory (v4)

`packages/cli-next/src/auth/index.ts` exports a factory that:

1. Detects CI mode (`CI`, `TF_BUILD`, `GITHUB_ACTIONS`, `CI_ENABLED=true`).
2. In CI: returns service-principal-backed providers (`azureLoginCI.ts`).
3. Otherwise: returns interactive MSAL providers (`m365Login.ts`, `azureLogin.ts`).

Both paths use the same encrypted cache at `~/.fx/account/`.

## Cache (`cacheAccess.ts`)

- AES-256-GCM at rest.
- Symmetric key:
  - If `keytar` installed and an OS keychain reachable → key in keychain.
  - Otherwise → derived from a machine-stable secret.
- Files: `~/.fx/account/m365TokenCache.bin`, `~/.fx/account/azureTokenCache.bin`.
- **Shared between v3 and v4** — sign in once via either CLI; both see the tokens.

## MSAL configuration

| Constant | Value |
|----------|-------|
| Client ID (M365) | Standard atk client ID — see `constants.ts` |
| Authority | `https://login.microsoftonline.com/common` (interactive) or tenant-specific (CI) |
| M365 default scopes | M365 (`https://graph.microsoft.com/.default`), App Studio (TDP) |
| Azure default scopes | `https://management.core.windows.net/.default` |

## Token refresh

- Silent refresh on every `getAccessToken` call — uses MSAL's `acquireTokenSilent`.
- Failure falls back to interactive flow (browser opens to `login.microsoftonline.com`).
- In CI, silent failure is fatal — interactive flow disabled.

## Windows broker (WAM)

When the native MSAL broker plugin is present and the OS is Windows, login goes through WAM (Web Account Manager) for SSO with the user's signed-in Windows account.

## CI: service principal

`azureLoginCI.ts` reads:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET` (or workload-identity OIDC token)

…and returns a provider backed by `ClientSecretCredential` (or `ClientAssertionCredential` for OIDC).

## Sign out

```bash
atk account logout m365
atk account logout azure
```

Each clears the corresponding cache file. VS Code extension exposes equivalent commands.

## Status callbacks

Providers expose `setStatusChangeMap()` so consumers can react to login / logout events (e.g. tree view refresh in VS Code).
