# Service clients

Source: [`packages/core-next/src/clients/`](../../../packages/core-next/src/clients/).

## Layout

```
clients/
├── teamsDevPortal/  — Teams Developer Portal API (app CRUD, validation, publish, OAuth, API keys)
├── graphApi/        — Microsoft Graph (Entra ID app reg, app catalog publish)
├── azure/           — Azure ARM (deployments, Kudu zip deploy)
├── m365/            — M365 PackageService (V1 classic + V2 DA sideloading)
└── index.ts         — barrel
```

## Common patterns

| Pattern | Why |
|---------|-----|
| Constructor takes `(ctx: AtkContext)` | Uses `ctx.auth.m365TokenProvider` or `ctx.auth.azureAccountProvider` for tokens; `ctx.logger` for logging; `ctx.telemetry` for metrics |
| HTTP via `sendWithRetry()` from `src/http/` | Automatic retry on 5xx + telemetry interceptors |
| 404 → `undefined` (not error) | Idempotent check-before-create patterns |
| Types co-located in client dir (`types.ts`) | Each client encapsulates its own surface |
| Returns `Result<T, AtkError>` | Universal error pattern |

## TeamsDevPortalClient

Methods cover the full TDP REST API surface:

- App CRUD (create, get, update, delete, list).
- Manifest validation (`/devportal/validate`).
- Package validation (`/devportal/validateAppPackage`).
- OAuth registration (Custom + MicrosoftEntra providers).
- API key registration.
- Publishing (legacy path; new publishing goes through `GraphApiClient`).

ZIP uploads validate **magic bytes** (`PK\x03\x04`) before sending.

## GraphApiClient

Microsoft Graph endpoints:

- Entra ID app registration: `POST /applications`, `PATCH /applications/{id}`.
- Password credentials: `POST /applications/{id}/addPassword`.
- App catalog publish: `POST/PUT /beta/appCatalogs/teamsApps` — replaces legacy TDP publish path.

Token scopes resolved per-call via `m365TokenProvider`.

## AzureArmClient

Azure ARM operations:

- ARM/Bicep deployment (`PUT /deployments/{name}`).
- Deployment status polling.
- Kudu zip deploy (`POST /api/zipdeploy`) — uses Azure credential to acquire SCM token, validates ZIP magic bytes.

## M365PackageService

Microsoft 365 sideloading via the MOS PackageService API:

- V1 (classic Teams apps).
- V2 (declarative agents).

Constructor takes `(ctx, token)` — the token is acquired separately by the caller (typically via `m365TokenProvider`). Returns `Result` for all operations.

## HTTP layer

`src/http/`:

- `createHttpClient(ctx)` — Axios instance with telemetry interceptors (request / response logging, secret-masked URLs).
- `sendWithRetry(client, config, options)` — retries on 5xx with exponential backoff.
- `sendWithTimeout(client, config, ms)` — `AbortController`-based timeout (replaces deprecated CancelToken).
- Pure functions, no `TOOLS` dependency.

## Tests

`tests/unit/clients/` — `teamsDevPortor` (sic, file name typo preserved), `graphApi`, `azure`. Each client mocks `axios` directly to keep tests deterministic.
