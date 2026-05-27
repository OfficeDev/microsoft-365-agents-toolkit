# Identity & Login

External-dependency fact page. Captures the **non-negotiable** identity / login
substrate the Microsoft 365 Agents Toolkit binds to today. Every fact below is
anchored in current source — change the code, and you must update this page.

This page does **not** decide engine-internal shape; that belongs in an ADR
under [`../adr/`](../adr/README.md). It only
records what the world outside the engine forces us to honor.

## 1. Facts the toolkit is bound to

### 1.1 First-party AAD application

| Field | Value | Notes |
|---|---|---|
| Client ID | `7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0` | Used by **both** VS Code and CLI surfaces for M365 user login. Public client (no secret). |
| Authority (public) | `https://login.microsoftonline.com/common` | |
| Authority (GCC H / DoD) | `https://login.microsoftonline.us/common` | Applies in sovereign-high environments (see §1.5). |
| Reply URLs registered on the app | `ms-appx-web://Microsoft.AAD.BrokerPlugin/7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0`, `https://vscode.dev/redirect`, `http://localhost`, `http://localhost:8400` | Closed set on the AAD side. The toolkit code today uses the first three (see §1.4); `http://localhost:8400` is registered for out-of-band use and is not referenced by this codebase. |

### 1.2 Auth libraries

| Surface | Library | Notes |
|---|---|---|
| VS Code (M365) | `@azure/msal-node` + `@azure/msal-node-extensions` | MSAL `PublicClientApplication`; token cache persisted via `msal-node-extensions`. |
| CLI (M365) | `@azure/msal-node` + `@azure/msal-node-extensions` | Same client ID as VS Code (§1.1). |
| VS Code (Azure) | **VS Code built-in `vscode.authentication.getSession("microsoft", …)`** | Does **not** use MSAL. The session is wrapped into a `TokenCredential`. |
| CLI (Azure) | `@azure/msal-node` + `@azure/msal-node-extensions` | MSAL-based, independent of the M365 CLI path. |

### 1.3 Native broker (WAM)

| Aspect | Behavior |
|---|---|
| Library | `@azure/msal-node-extensions` `NativeBrokerPlugin` |
| Platforms | Windows only. |
| Loading | The native dependency must not load on non-Windows platforms (avoids keytar / libsecret pull-in). |
| VS Code gating | Opt-in at runtime; off by default. |
| CLI gating | On whenever the plugin loads; silent fallback otherwise. |
| Quirk | Broker does **not** support `forceRefresh`; surfaces pass `claims` to force a refresh. |
| Sign-out scope | When broker is available, sign-out clears **only the cached account**, not all accounts. |

### 1.4 Redirect / reply URIs per flow

| Flow | Redirect URI used at runtime |
|---|---|
| Native broker (WAM, Windows) | `ms-appx-web://Microsoft.AAD.BrokerPlugin/7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0` (constructed by MSAL's `NativeBrokerPlugin`) |
| VS Code web (vscode.dev / Codespaces) | `https://vscode.dev/redirect` |
| Local browser code flow (CLI + VS Code desktop) | `http://localhost:<port>` where `<port>` is chosen by the OS (the local Express server is started with `SERVER_PORT = 0`); matched on the AAD side by the loopback-wildcard registration `http://localhost` |

### 1.5 Sovereign-cloud endpoint binding

Resource endpoints, scopes, and the Entra authority are **all sovereign-aware**.

Four environments are supported: `Public`, `GCC M`, `GCC H`, `DoD`. Selection
is runtime-configurable.

Per-environment endpoint matrix:

| Service | Public | GCC M | GCC H | DoD |
|---|---|---|---|---|
| AuthSvc | `teams.microsoft.com/api/authsvc` | same | `authsvc.gov.teams.microsoft.us` | `authsvc.dod.teams.microsoft.us` |
| AuthSvc audience | `api.spaces.skype.com` | same | `authsvc.teams.microsoft.com` | same |
| TDP (Teams Dev Portal) | `dev.teams.microsoft.com` | same | `gov.dev.teams.microsoft.us` | `dod.dev.teams.microsoft.us` |
| MOS3 | `titles.prod.mos.microsoft.com` | `titles.gccm.mos.microsoft.com` | `titles.gcch.mos.svc.usgovcloud.microsoft` | `titles.dod.mos.svc.usgovcloud.microsoft` |
| Graph | `graph.microsoft.com` | same | `graph.microsoft.us` | `dod-graph.microsoft.us` |
| Azure ARM | `management.azure.com` | same | `management.usgovcloudapi.net` | same as GCC H |
| TeamsGraph | `teams.microsoft.com/api/platform` | `teams.microsoft.com/gcc/api/platform` | `gov.teams.microsoft.us/api/platform` | `dod.teams.microsoft.us/api/platform` |
| TeamsGraph audience | `teamsgraph.teams.microsoft.com` | same | `teamsgraph.gov.teams.microsoft.us` | `teamsgraph.dod.teams.microsoft.us` |

An **internal dogfood TDP** exists at `dev-int.teams.microsoft.com`. When the
toolkit is pointed at the dogfood mode, TDP requests target this host
instead of the sovereign-default; the selector mechanism is toolkit-defined
(see code map).

### 1.6 Scope contracts

What each external service requires the toolkit to ask for. "Resource" in the
"Scope string" column shifts per sovereign environment (see §1.5); `<tenant>`
is the user's tenant URL. **Microsoft Graph permission strings are
sovereign-independent literals**; all other resource-prefixed scopes shift
per §1.5.

| Purpose | Scope string the toolkit must request |
|---|---|
| Teams Developer Portal (manifest CRUD, bot registration) | `<TDP-resource>/AppDefinitions.ReadWrite` |
| AuthSvc region discovery | `<AuthSvc-audience>/Region.ReadWrite` |
| Azure ARM | `<ARM-resource>/.default` |
| M365 launch / titles service (MOS3) | `<MOS3-resource>/.default` |
| Teams Graph | `<TeamsGraph-audience>/.default` |
| Microsoft Graph — baseline app management | `Application.ReadWrite.All`, `TeamsAppInstallation.ReadForUser` |
| Microsoft Graph — group / member search | `GroupMember.Read.All` |
| Microsoft Graph — external connection (Graph Connectors) | `ExternalConnection.Read.All` |
| Microsoft Graph — basic user lookup | `User.ReadBasic.All` |
| Microsoft Graph — Teams app settings (tenant admin read) | `TeamworkAppSettings.Read.All` |
| Microsoft Graph — Teams team creation | `Team.Create` |
| Microsoft Graph — Teams channel creation | `Channel.Create` |
| Microsoft Graph — Teams team listing | `Team.ReadBasic.All` |
| Microsoft Graph — Teams channel listing | `Channel.ReadBasic.All` |
| Microsoft Graph — install Teams app to a team | `TeamsAppInstallation.ReadWriteAndConsentForTeam` |
| Microsoft Graph — tenant app catalog (CRUD) | `AppCatalog.ReadWrite.All` |
| Microsoft Graph — sensitivity label list | `InformationProtectionPolicy.Read` |
| SPFx deploy | `<tenant>/Sites.FullControl.All` |

One conditional behavior: on the **launch and provision** paths under
sovereign-high (GCC H / DoD), the toolkit substitutes the Microsoft Graph
baseline scopes for the TDP-resource scope because **AppDefinitions.ReadWrite
is not granted on sovereign-high TDP**.

### 1.7 Region endpoint discovery (TDP)

TDP is **region-sharded**. Before serving a user's first TDP request, the
toolkit must perform a handshake against AuthSvc with an AuthSvc-scoped token;
AuthSvc returns the per-region TDP base URL that user is assigned to.
Subsequent TDP requests for that user are expected to hit the returned regional
URL; when no per-region URL is known, the sovereign-default TDP endpoint from
§1.5 is used.

Region discovery does **not** apply to the internal-dogfood TDP (§1.5),
which is a single-region service.

> **Where these facts live in code today:** see
> [`identity-and-login.code-map.md`](identity-and-login.code-map.md). The code
> map is a navigation aid and is expected to churn with refactors; it is not
> part of the contract.

## 2. Constraints derived from these facts

Rules a refactor must honor; each is a direct consequence of §1.

1. **No new AAD client ID without coordination.** The `7ea7c24c-…` first-party
   app is the only client ID approved for M365 user login from this codebase.
2. **No new reply URI without coordination.** The closed set is what is
   registered on the AAD app (§1.1). Any new flow must reuse one of those
   registered URIs.
3. **Sovereign-aware URL resolution is mandatory.** Any code that talks to
   AuthSvc, TDP, MOS3, Graph, Azure ARM, or TeamsGraph must resolve the host
   per the current sovereign environment (§1.5). Inline hosts are forbidden.
4. **Sovereign-aware Entra authority resolution is mandatory.** No literal
   `login.microsoftonline.*` host anywhere in this codebase; the authority
   must follow the sovereign environment selection in §1.1.
5. **Broker is optional, Windows-only, dynamically loaded.** No top-level
   `import` of `@azure/msal-node-extensions` from cross-platform code paths.
6. **Only the login-owning module may import MSAL.** No other code in this
   codebase may take a direct dependency on `@azure/msal-node` or
   `@azure/msal-node-extensions`; token acquisition is concentrated in one
   place per surface.
7. **Azure credentials on VS Code do not have MSAL semantics.** The Azure
   credential there is backed by `vscode.authentication.getSession`; cache
   invalidation and silent refresh behave differently from the MSAL path on
   the CLI.
8. **TDP usage is unsafe until region discovery has resolved.** Per §1.7, no
   TDP request may be issued for a user before the AuthSvc handshake has
   completed (or the sovereign default has been adopted).
9. **Sovereign-high (GCC H / DoD) has no AppStudio scope.** Code that picks
   scopes for the launch / provision path must request the Microsoft Graph
   baseline (§1.6) instead of the TDP-resource scope on these environments.

## 3. Open questions (candidates for ADRs)

These are *not* facts — they are gaps the refactor needs to close. Each is
tracked as a `Proposed` ADR stub under
[`../adr/`](../adr/README.md). The
removed engine ↔ surface token-boundary discussion is tracked there as
[ADR-0001](../adr/ADR-0001-engine-surface-token-boundary.md).

- **CLI ↔ VS Code login duplication.** The M365 MSAL plumbing is duplicated
  between the CLI and VS Code surfaces (same client ID, same broker handling
  on Windows). Should this live in a shared package, or is the duplication
  deliberate (e.g. to avoid leaking VS Code APIs into CLI)? →
  [ADR-0002](../adr/ADR-0002-m365-login-plumbing-sharing.md)
- **Broker gating asymmetry.** VS Code defaults broker off (opt-in); CLI
  defaults broker on. Intentional, or should both surfaces share one policy?
  →
  [ADR-0003](../adr/ADR-0003-broker-gating.md)
- **Region-discovery side effect.** Should TDP region state live on the
  TDP-client module as singleton state, or be carried per-request (relevant
  for any future multi-tenant / multi-region work)? →
  [ADR-0004](../adr/ADR-0004-tdp-region-state.md)
- **Internal-dogfood TDP selector.** Selection is currently a one-off env
  read; should it go through the same configuration model as other
  env-driven behavior? →
  [ADR-0005](../adr/ADR-0005-env-override-config-model.md)

---

> **How to update this page:** changes to the facts in §1 require updating
> [`identity-and-login.code-map.md`](identity-and-login.code-map.md) in the
> same PR. Adding a new constraint to §2 requires an accepted ADR under
> [`../adr/`](../adr/README.md); link it from
> the constraint line.
