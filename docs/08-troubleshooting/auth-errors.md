# Auth errors

## `M365LoginFailed` / `AzureLoginFailed`

**Trigger.** MSAL flow could not complete.

**Mitigation.**

- Interactive: ensure a browser opens and you complete the login. Behind a corporate proxy? Set `HTTPS_PROXY` env var.
- CI: verify service-principal env vars (`AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`) or workload-identity OIDC config.

## `TokenCacheCorrupted`

**Trigger.** `~/.fx/account/*.bin` cannot be decrypted (key changed, or file truncated).

**Mitigation.** `atk account logout m365` and `atk account logout azure`, then sign in again. The cache will be regenerated.

## `TokenExpired`

**Trigger.** Cached refresh token is past its lifetime.

**Mitigation.** Sign out and back in. In CI, this typically means the service-principal secret expired — rotate it.

## `Forbidden` on M365

**Trigger.** Signed-in user lacks the M365 permission required (e.g. publish to org catalog).

**Mitigation.** The error's `helpLink` points at the relevant admin permission. Either obtain the role yourself or have an admin grant it.

## `Unauthorized` (401) intermittent

**Trigger.** Token race or stale token used after rotation.

**Mitigation.** Re-run the command. If persistent, sign out and back in.

## Keytar / OS keychain failure

**Trigger.** `keytar` cannot reach the OS keychain (Linux without libsecret, sandbox restrictions).

**Mitigation.** This is not fatal — the toolkit falls back to a machine-derived key. Cache encryption is weaker but still active. Install `libsecret-1-dev` (Linux) to enable keytar fully.

## CI: workload-identity setup

If using workload identity federation:

- The federated credential subject must match the workflow's claims.
- Pre-flight: `az login --service-principal --federated-token "$ID_TOKEN"` should succeed locally before CI is set up.
- The OIDC `audience` in `azure/login` action must be `api://AzureADTokenExchange`.

See [03-infrastructure/cicd-pipelines.md](../03-infrastructure/cicd-pipelines.md).
