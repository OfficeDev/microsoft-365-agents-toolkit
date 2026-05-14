# Tenant and sovereign cloud flows (v3)

> **v4 design status.** This page is an allowed input to v4 design. It captures user-facing flows around multi-tenant accounts and sovereign cloud routing — what the user does and sees — without committing to v3's internal MSAL plumbing.

## Tenant switching

Users with multiple M365 or Azure tenants can switch the active tenant without signing out. Two parallel flows: M365 and Azure.

### M365 tenant switch

| Trigger | Flow |
|---------|------|
| Command palette: `Switch between your available tenants (M365)` | Lists tenants the signed-in account belongs to → user picks one → re-auth as that tenant |
| Tree view: hover signed-in M365 account → switch icon | Same flow |
| CLI: `atk auth login m365 --tenant <id>` | Direct switch by tenant ID, no list |

Source: [`packages/vscode-extension/src/handlers/accounts/switchTenantHandler.ts`](../../../packages/vscode-extension/src/handlers/accounts/switchTenantHandler.ts), `packages/vscode-extension/src/commonlib/m365Login.ts` (`switchTenant`).

What the user sees:

1. Status notification: "Switching to tenant {id}…".
2. Brief auth pop-up if the new tenant requires re-consent.
3. Tree views (Environment, Accounts) refresh to reflect new tenant's apps and resources.
4. In-flight provision/deploy operations are **not cancelled** — they continue against the original tenant. (This is a known sharp edge.)

### Azure tenant switch

| Trigger | Flow |
|---------|------|
| Command: `Switch between your available tenants (Azure)` | Same shape as M365 |
| CLI: `atk auth login azure --tenant <id>` | Direct switch |

The Azure switch additionally re-fetches the subscription list (subscriptions are scoped to the active tenant).

### Telemetry

Property: `TelemetryProperty.SovereignCloudType` (records which cloud — for cross-cloud usage analytics).

## Sovereign cloud routing

Microsoft's sovereign clouds — **GCC M** (commercial-equivalent for US government), **GCC H** (high-security US gov), **DoD** (Department of Defense) — use different login endpoints, different Microsoft Graph hostnames, and different Azure ARM endpoints from commercial cloud.

### Configuration

VS Code setting:

```jsonc
"M365AgentsToolkit.sovereignCloudEnvironment": "" | "GCC M" | "GCC H" | "DoD"
```

Source: `packages/vscode-extension/package.json` `contributes.configuration`. Setting key constant: `FeatureFlags.SovereignCloudEnvironment`.

For GCC H or DoD, the user must **also** set the VS Code built-in setting:

```jsonc
"microsoft-sovereign-cloud.environment": "USGovernment"
```

Then **reload the window** for the change to take effect.

### What changes when set

Internally the auth subsystem (`packages/vscode-extension/src/commonlib/codeFlowLogin.ts`, `packages/vscode-extension/src/commonlib/azureLogin.ts`) routes to the appropriate endpoints:

| Cloud | M365 login authority | Graph host | ARM host |
|-------|---------------------|------------|----------|
| Commercial (default) | `login.microsoftonline.com` | `graph.microsoft.com` | `management.azure.com` |
| GCC M | `login.microsoftonline.com` | `graph.microsoft.com` (commercial-equivalent) | `management.azure.com` |
| GCC H | `login.microsoftonline.us` | `graph.microsoft.us` | `management.usgovcloudapi.net` |
| DoD | `login.microsoftonline.us` | `dod-graph.microsoft.us` | `management.usgovcloudapi.net` |

### What the user sees

No separate "GCC setup" wizard. The setting toggle is the entire user-facing surface. After reload:

- Login pop-ups go to sovereign Microsoft sign-in pages.
- Tokens returned are sovereign-issued.
- App registrations land in the sovereign tenant.
- Deployed Azure resources land in sovereign-cloud subscriptions.

If the user accidentally signs into a commercial account while configured for sovereign (or vice versa), the toolkit shows an authentication-failure error pointing at the mismatch.

### CLI

`atk` does not currently expose a sovereign-cloud setting CLI flag — sovereign routing is via env var `TEAMSFX_SOVEREIGN_CLOUD` (when present) or via the VS Code setting (when launched from a VS Code-managed terminal that inherits the workspace setting).

### Migration path

There is **no automated migration** between commercial and sovereign clouds. Moving an existing project from commercial to GCC requires:

1. Provisioning a new app registration in the sovereign tenant.
2. Re-running `atk provision --env <new-env>` against a sovereign-cloud subscription.
3. Updating manifest IDs in source control.

## Auth-related env vars (for CI / scripted flows)

Set these before invoking `atk` in non-interactive mode:

| Env var | Purpose |
|---------|---------|
| `M365_ACCOUNT_NAME` / `M365_ACCOUNT_PASSWORD` | M365 user-flow auth (legacy; not recommended for production CI) |
| `M365_TENANT_ID` | M365 tenant target |
| `TEAMS_APP_TENANT_ID` | Override Teams app tenant (used for cross-tenant publishing) |
| `AZURE_ACCOUNT_NAME` / `AZURE_ACCOUNT_PASSWORD` | Azure user-flow auth (legacy) |
| `AZURE_SERVICE_PRINCIPAL_ID` / `AZURE_SERVICE_PRINCIPAL_SECRET` | Azure SP auth (recommended for CI) |
| `AZURE_TENANT_ID` | Azure tenant |
| `AZURE_SUBSCRIPTION_ID` | Default subscription |
| `AZURE_ACCOUNT_OBJECT_ID` | Account object ID (for advanced auth scenarios) |

CI mode is triggered by `CI_ENABLED=true` (forces non-interactive). See [cli-v3-command-reference.md §"Interactive vs non-interactive"](../surfaces/cli-v3-command-reference.md#interactive-vs-non-interactive).

## See also

- [auth-flows-m365-azure.md](auth-flows-m365-azure.md) — base M365 + Azure sign-in flow
- [../surfaces/cli-v3-command-reference.md](../surfaces/cli-v3-command-reference.md) — `atk auth` commands
- [../surfaces/vscode-extension-commands.md](../surfaces/vscode-extension-commands.md) — account-management commands
- [../../01-product/v3-feature-inventory.md](../../01-product/v3-feature-inventory.md) — sovereign clouds in the integration surface
