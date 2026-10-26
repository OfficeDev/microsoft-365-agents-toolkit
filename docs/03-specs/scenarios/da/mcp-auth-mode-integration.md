# Remote MCP Authentication Mode Integration

- **Status:** Requester-defined contract, 2026-10-21. Registration-boundary tests and live UI integration are reported separately.
- **Product scenario:** [Create DA with MCP server](../../../01-product/scenarios/da/create-da-with-mcp-server.md), `SCN-DA-CREATE-WITH-MCP-SERVER`.
- **Modes:** Static OAuth, DCR, and Microsoft Entra SSO, each with an independent project, environment mapping, registration state, and browser/token state.
- **Execution target:** Microsoft 365 Copilot with the licensed atktesting08 account.
- **Fixture:** The requester-owned remote MCP facade.
- **Credential source:** An explicitly selected SSO-specific environment file. Do not fall back to the generic account profile when this file is unavailable.
- **Current resource identity:** atktesting08 tenant with client ID
  `7c5898c1-99fb-4d83-8474-02bf7985096c`, read from the latest consumer export.
  Reject a stale client identity rather than loading the previous application's credentials.

## Required Separation

| Mode                  | Registration behavior                                                                                                   | Authorization/token behavior                                                                                                                 |
| --------------------- | ----------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
| Static OAuth          | Execute its own `oauth/register` as Custom; never seed its configuration ID from the SSO ID                             | Use the fixture's Entra authorization/token endpoints and securely injected static AAD credentials.                                          |
| DCR                   | Execute its own `dcr/register`; never reuse the SSO configuration ID                                                    | Use MCP-facade discovery, registration, authorization and dynamically managed credentials; do not insert static Entra credentials/endpoints. |
| Entra SSO, ID present | Read and validate the pre-registered record through `oauth/register`; skip creation and all creation-time configuration | Execute real SSO/token/MCP calls using the existing registration. Preserve mini-owned AAD URI, callbacks, and preauthorization.              |
| Entra SSO, ID absent  | Retain the normal complete registration/configuration workflow                                                          | Execute real SSO/token/MCP calls after creation.                                                                                             |

Only the SSO test preparation branch may read `ENTRA_SSO_REGISTRATION_ID`. The
product driver continues to receive the ordinary mapped configuration-ID
environment variable; it does not introduce a global read of the SSO test setting.

A supplied but inaccessible/missing record must fail, not warn-and-succeed or
create a replacement. A record of another authentication kind, or one whose
client ID conflicts with the supplied SSO client, must also fail.

## Composed Operations

- [Register OAuth](../../operations/auth/register-oauth.md).
- [Register DCR](../../operations/auth/register-dcr.md).
- [Create MCP server](./create-mcp-server.md).

## Acceptance Criteria

| ID          | Runtime | Purpose               | Gate                    | Harness                                                              | Given / When                                                                                              | Then                                                                                                                                                                                               |
| ----------- | ------- | --------------------- | ----------------------- | -------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| MCP-AUTH-01 | L1      | operation-integration | required                | Real OAuth driver with fake TGS HTTP                                 | SSO test ID is present but the static-OAuth case runs with its own empty output mapping                   | POST a separate Custom registration using Entra endpoints/static credentials; do not GET or reuse the SSO ID.                                                                                      |
| MCP-AUTH-02 | L1      | operation-integration | required                | Real DCR driver with fake TGS HTTP                                   | SSO test ID is present but the DCR case runs with its own empty output mapping                            | POST a separate Dynamic configuration using the MCP facade; do not GET/reuse the SSO ID or send static credentials/Entra endpoints.                                                                |
| MCP-AUTH-03 | L1      | operation-integration | required                | Real OAuth driver with fake TGS HTTP                                 | SSO supplies a valid matching MicrosoftEntra ID                                                           | GET the existing record, skip POST and creation-time Graph/configuration work, and leave existing resource URI/reference state unchanged.                                                          |
| MCP-AUTH-04 | L1      | operation-integration | required                | Real OAuth driver with fake TGS HTTP                                 | SSO omits its optional registration ID                                                                    | Execute normal MicrosoftEntra registration with SSO enabled and AnyApp/HomeTenant, preserving the creation result mappings.                                                                        |
| MCP-AUTH-05 | L1      | operation-integration | required                | Real OAuth driver with fake TGS HTTP                                 | Supplied SSO ID returns an error, another provider, or a conflicting client ID                            | Return an explicit error; no successful skip, no replacement POST and no configuration mutation.                                                                                                   |
| MCP-AUTH-06 | L1      | compatibility         | required                | Independent mode fixtures and fake TGS HTTP                          | Each case creates/loads its own mode-specific output variables                                            | SSO reuse never populates static-OAuth/DCR outputs; fresh static/DCR registrations return distinct IDs; preserve legacy Custom reuse-error behavior outside SSO.                                   |
| MCP-AUTH-07 | L1      | operation-integration | required                | Real drivers with failing token provider                             | Authentication fails in any mode                                                                          | Return the authentication error before HTTP registration or reuse; do not fall back to another mode.                                                                                               |
| MCP-AUTH-08 | L3      | surface               | required-for-live-claim | Three independent Copilot UI executions plus real flow evidence      | Run all three modes against the dedicated fixture/account                                                 | Each returns the required tenant, canary and arithmetic result; actual registration/token evidence matches the selected mode.                                                                      |
| MCP-AUTH-09 | L3      | surface               | required-for-live-claim | Copilot UI with isolated SSO present/absent/invalid fixtures         | Exercise all SSO ID variants                                                                              | Present ID skips creation only, absent ID retains full setup, invalid ID visibly fails; valid variants still obtain tokens and invoke MCP.                                                         |
| MCP-AUTH-10 | L1      | compatibility         | required                | Mode profiles and prepared project fixtures                          | Load the latest consumer export for each mode                                                             | Validate the requested tenant/current client, resolve the endpoint at runtime, inject matching static credentials only for static OAuth, and give each run a distinct ownership ID.                |
| MCP-AUTH-11 | L1      | operation-integration | required                | Real teardown coordinator with temporary projects and deletion ports | Teardown receives shared identity IDs, a reused/invalid SSO ID, foreign-run resources, or owned resources | Skip mini-managed AAD, SP, monitoring apps and supplied SSO records even with a direct-by-ID deletion port; delete only supported resources whose creation baseline belongs to this run.           |
| MCP-AUTH-12 | L1/L3   | surface               | required                | Compiled terminal steps and real vscuse desktop                      | Prepare a new project or reject a broken fixture before provision                                         | Execute the exact supplied script, validate files and mode wiring, require a current-run success marker, and stop before cloud operations on failure; code-agent completion is not script success. |

## Live Result Oracle

All three live cases must assert actual tool results:

- `whoami.tenant_id == "16801f2e-7163-4bd5-8542-f083fa59993c"`.
- `auth_flow_probe.canary == "PURPLE-NARWHAL-7788"`.
- `add(20, 7) == 27`.

These values alone do not prove authentication mode. Correlate the UI/tool run
with its actual registration and token process, including static Entra
authorization versus facade DCR versus Entra SSO. Use fresh case state and
redacted evidence; never archive tokens, credentials, or authorization codes.

The authenticated `whoami` result must contain a non-empty `client_id`; the UI
test does not prescribe a Microsoft broker or downstream JWT caller GUID.
`MCP_AUTH_EXPECTED_JWT_CLIENT_ID` is not required or generated. The caller can
differ from the resource application and from the facade's dynamically
registered OAuth client. Static OAuth and DCR can share a downstream caller and
probe label; their actual registration and authorization/token paths must still
distinguish them. The exact tenant, `entra_user`, canary and arithmetic checks
remain unchanged.

`MCP_AUTH_RUN_ID` is generated with a new UUID for every profile/run. It binds
fixture preparation and ownership-safe teardown, not the remote authentication
protocol. It is not a pre-registered ID or a manually configured GitHub variable.

Health checks, probe classification, fake tokens, and mocked responses cannot
qualify a case as a live end-to-end pass. The L1 tests cover only registration
selection/isolation and failure semantics; they do not simulate the three tool
results or claim Copilot coverage.

## Flow

```mermaid
flowchart TD
  Env[Explicit dedicated environment file] --> Isolate[Independent project and auth state]
  Isolate --> Mode{Selected mode}
  Mode -- Static OAuth --> Static[Create Custom registration with Entra endpoints]
  Mode -- DCR --> Dynamic[Create configuration using MCP facade discovery]
  Mode -- Entra SSO --> HasId{SSO registration ID supplied?}
  HasId -- yes --> Read[Validate matching existing MicrosoftEntra record]
  Read -- invalid --> Fail[Fail without mutation or replacement]
  Read -- valid --> Preserve[Reuse without mini AAD changes]
  HasId -- no --> Create[Normal SSO creation and configuration]
  Static --> Tokens[Real mode-specific authorization and token acquisition]
  Dynamic --> Tokens
  Preserve --> Tokens
  Create --> Tokens
  Tokens --> Copilot[Invoke and inspect all three tools in Copilot UI]
  Copilot --> Evidence[Correlate tool results with registration and token evidence]
```

## Boundary and Live Fixture

- Do not overwrite mini-owned AAD settings when reusing a registration.
- Do not reuse a browser, registration, environment or token artifact from another mode.
- Every generated auth fixture records its initial resource state and run ID before
  provisioning. MCP teardown must use this ownership record; it must not call
  generic AAD/SP cleanup or inherit a resource-group ID from the ambient account.
  Mini-managed identities remain shared even in the absent-SSO-ID case.
  A supplied SSO record, including an invalid one, is never test-owned.
  For the forced `absent` variant, the file's pre-registered ID is retained only
  as protected teardown metadata, never as the active registration input.
  Missing/invalid ownership must not fall back to destructive generic cleanup.
- Cleanup APIs not supported by the existing test harness are reported as retained,
  not silently deleted or treated as successfully cleaned. In particular, the
  generic cleaner does not invent an OAuth/DCR registration DELETE contract.
- The run-scoped guard also applies to direct-by-ID AAD/SP cleanup helpers.
  Outside an MCP run, their existing behavior is preserved. Missing ownership
  for an authenticated MCP project fails closed rather than invoking generic
  identity or ambient-resource-group deletion.
- Do not guess the dedicated environment filename or silently substitute the generic profile.
- Use the mini project's secret-free `.generated/da-sso.atktesting08.env` as the
  explicit consumer profile. Static OAuth additionally requires its matching
  private Entra credential file, loaded only for that mode. Do not import local
  HMAC/static bearer credentials or substitute the generic account profile's AAD
  fixture.
- The consumer profile supplies the tenant/MCP/client identity. The `atk08`
  account profile supplies the licensed M365 account and runner infrastructure
  credentials, not the resource application's static credentials.
- The launcher `run-mcp-auth-integration.ps1` accepts the consumer path explicitly
  and applies mode-scoped environment references. Its SSO `absent` variant retains
  registration and adds guarded AAD configuration from the mini-owned manifest,
  preserving the existing URI/callback/preauthorization data. The `invalid`
  variant must fail registration validation rather than proceed to Copilot.

## Running the Independent UI Cases

These cases use the [product feature defaults](../../../../packages/fx-core/src/common/featureFlags.ts)
instead of pinning `TEAMSFX_MCP_FOR_DA_DT`, `TEAMSFX_MCP_FOR_DA_DCR`, or
`TEAMSFX_V4_ENABLED` in case metadata or temporary runner configuration.
They do not force the v4 preview route. MCP endpoint, account, credential, and
assertion-value references still come from the selected runtime fixture.

The semantic cases are in
[`da-mcp-auth-modes.yml`](../../../../packages/tests/vscuse/vscode-test-cases/cases/da-mcp-auth-modes.yml).
Each generated plan is flat and independently creates its own project. The
[launcher](../../../../packages/tests/vscuse/vscode-test-cases/run-mcp-auth-integration.ps1)
refuses an existing vscuse container rather than inheriting another case's browser
or token state. Use a current branch-specific local VSIX/image.

Set the two local path variables below outside source control. The consumer
export is secret-free; the static credential file contains the matching Entra
client secret and is used only by static OAuth:

```powershell
.\packages\tests\vscuse\vscode-test-cases\run-mcp-auth-integration.ps1 `
  -Mode static-oauth -SsoEnvironmentFile $env:MCP_AUTH_CONSUMER_ENV_FILE `
  -StaticCredentialFile $env:MCP_AUTH_STATIC_CREDENTIAL_FILE

.\packages\tests\vscuse\vscode-test-cases\run-mcp-auth-integration.ps1 `
  -Mode dcr -SsoEnvironmentFile $env:MCP_AUTH_CONSUMER_ENV_FILE

.\packages\tests\vscuse\vscode-test-cases\run-mcp-auth-integration.ps1 `
  -Mode entra-sso -SsoEnvironmentFile $env:MCP_AUTH_CONSUMER_ENV_FILE
```

For the SSO group, `-SsoVariant absent` deliberately exercises creation/configuration
even when the consumer export contains a shared ID. `-SsoVariant invalid` supplies
an invalid ID and must produce a failing registration, not a positive Copilot
pass. `auto` follows the file: reuse a supplied ID or retain full setup when absent.
`-Author` opens the real authoring UI with an ignored copy of just that flat plan.
The absent-ID run writes a literal empty optional ID into its temporary Docker
configuration; it does not depend on empty environment-variable interpolation.

The OAuth cases first use the recorded Copilot `Sign in to <app>` connection
fragment, not `Allow` as an equivalent entry state. DCR then requires the facade's
`Authorize an MCP client` page and `Allow this client` consent before Entra
authorization and selection of the expected signed-in M365 account. After
returning to Copilot, the OAuth cases reuse the ordinary `Allow` capability-consent
fragment. Static OAuth must not visit the facade consent page, and SSO retains
ordinary capability consent without invoking the OAuth connection adapter.
This first-use client authorization is
unrelated to the removed provisioning-time DCR-update confirmation. Completing
these interactions alone never satisfies the real token/tool-result checks.

The identity and probe prompts request only their asserted fields from the
actual tool results, including a new `probe_id` and fresh `server_time` for the
probe. This keeps unrelated identity details from crowding the
required values out of Copilot's JSON rendering. Expected tenant, client, canary,
and arithmetic answers are not embedded in those user messages.

The UI invocation assertion accepts returned tool-specific data in the completed
response, including ordinary inline JSON. It does not require an execution
marker, citation or debugger panel. As confirmed by the requester, missing
debugger information after `-developer on` is a known upstream issue, so these
cases neither enable developer mode nor depend on its output. Refusals, pending
calls, descriptions and promises without returned data still fail. Registration
and token-flow evidence remain separate from these UI assertions.

Runner logs and reports are stored under a unique ignored
`.local/test-reports/` directory. Known secrets, JWTs and callback authorization
codes are redacted from text artifacts. A setup failure (including invalid Azure
OpenAI/vision runner credentials) is not a business-flow pass or a verified
authentication-mode failure.

## Validation Status

The three generated plans remain on the manual gate. Local registration-boundary
tests and fixture/compiler tests pass, including mode isolation, SSO reuse and
creation, invalid IDs, and authentication failures. Tool-result checks require the
named `tenant_id`, `canary`, and `sum` fields, not just a matching UUID, canary, or
number in the response.

A real noVNC DCR authoring run reached the generated project. It selected
`OAuth (with dynamic registration)` and verified `dcr/register`, the
`RemoteMCPServer` runtime, empty functions, wildcard dispatch, and the absence of
the legacy static tool file. This is create-flow evidence only.

The initial runner `401 PermissionDenied` was traced to stale OpenAI and Vision
keys. Authorized retrieval of the current keys restored runner access without
persisting key values. A clean DCR CLI rerun completed 109 steps, including real
M365 sign-in, successful provisioning, opening the owned agent in Copilot, and
requesting `whoami`. Its next screen required `Sign in to <app>`, whereas the
existing action-consent component expected `Allow`. This is an uncovered OAuth
connection step, not an authentication or tool-call pass.

That run also exposed a false-success preparation path: the runner rewrote a
CodeAgent example, omitted `PROJECT_DIR`, and accepted completion despite the
script's nonzero exit. Preparation now uses exact native-terminal input and a
mandatory current-run ready marker, emitted only after project validation,
mode-specific configuration, and ownership persistence succeed. A real
missing-project UI run failed at this gate before provisioning; a separate real
DCR project passed it and produced matching run/project ownership with initially
absent registration, Teams app, and M365 title state.
Preparation also clears the scaffold notification before opening the terminal:
an actual rerun showed that the toast could cover the UUID in successful output.
The current-run assertion remains strict; a bare or obscured marker is not success.

The recorded Sign in action subsequently reached the real facade authorization
endpoint and returned `invalid_client`. The fixture revision and consumer export
were refreshed during that authoring window. This is a failed authentication
observation, not proof of token acquisition, and must be revalidated with fresh
latest-profile registration state.

A subsequent latest-profile DCR registration reached the real facade client
consent page and its Entra authorization redirect; the former `invalid_client`
was not reproduced for that fresh registration. The two public interactions were
recorded and extracted into reusable components with captured visual guards.
Compiler regressions verify their ordering and authentication-mode isolation.
This remains authorization-entry evidence, not a completed token/MCP run.

Further authoring recorded the Entra account-picker selection and Copilot's
post-authentication `Allow` card. Correlated facade logs confirmed a real Entra
callback and token exchange, and Copilot returned the requested tenant and
`entra_user` identity. The complete `client_id` was visible in a subsequent
compact response, but the full raw-JSON response did not visibly expose its
complete value and the compact turn did not satisfy the explicit invocation
assertion. Those observations are not an all-tool or clean-run pass.

A separate preview reached the owned agent's Store details with `Open` instead
of chat. That transition was recorded and used during authoring, but it has not
been wired as a guessed optional fallback around the existing chat-ready guard.
That preview-entry variation remains a separate follow-up, not a reason to
bypass the chat-ready guard.

The archived clean three-mode batch on 2026-10-23 used the short result projections.
All three reached Copilot and passed every `whoami` identity assertion. DCR and
static OAuth used the current resource-client identity; SSO used the expected
native token-store caller. Their mode-specific preparation/connection paths
remained separate.

| Mode         | Successful / executed steps | Final result                                                                                 |
| ------------ | --------------------------- | -------------------------------------------------------------------------------------------- |
| DCR          | 139 / 140                   | HOLD at the explicit completed-invocation assertion for `auth_flow_probe`.                   |
| Static OAuth | 144 / 145                   | HOLD at the same probe-invocation assertion.                                                 |
| Entra SSO    | 134 / 135                   | HOLD at the same probe-invocation assertion after reuse and the native identity checks pass. |

These archived failures are test-plan drift: the old assertion demanded a
separate completed-call marker that the Copilot UI does not guarantee. They
are not evidence of an authentication failure. Those runs stopped before `add`
and are not retroactively counted as passing.

The corrected contract accepts the returned tool result and retains every
identity/canary/arithmetic assertion. The probe additionally requests and checks
server-generated `probe_id` and `server_time` rather than relying only on constants
from the tool description. The experimental runtime observer, collector and
witness-tab approach has been removed; no MCP service deployment or shared
identity changes are required for this repair.

The corrected cases were executed from separate clean containers with atk08 on
2026-10-26, without developer mode or feature-flag overrides:

| Mode         | Successful / executed steps | Result |
| ------------ | --------------------------- | ------ |
| Entra SSO    | 146 / 146                   | PASS   |
| DCR          | 151 / 151                   | PASS   |
| Static OAuth | 156 / 156                   | PASS   |

Each completed its mode-specific registration or reuse, authentication and
Copilot connection, the required `whoami` identity checks, the probe's canary,
mode and server-generated fields, and the real `add(20,7)` request with sum 27.
No separate debugger or execution-marker UI was required. Reports and redacted
logs are archived under the unique run directories
`20261026-100636-da-mcp-entra-sso-copilot-auto`,
`20261026-102927-da-mcp-dcr-copilot-auto`, and
`20261026-105305-da-mcp-static-entra-copilot-auto` in `.local/test-reports/`.

The SSO live pass exercises the consumer export's supplied-ID reuse path.
Absent-ID and invalid-ID behavior remains covered by registration-boundary and
fixture regressions; these runs do not claim live coverage for those variants.
The invalid-ID variant is an expected failure of the positive SSO plan, not a
separately passing negative UI plan.

The later non-empty `client_id` assertion removes the fixed JWT-caller
expectation without changing registration or authentication. Compiler/profile
regressions cover this follow-up; the archived UI runs above precede it and are
not presented as reruns of the updated assertion.
