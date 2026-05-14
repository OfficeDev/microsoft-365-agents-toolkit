# PRD overview

A condensed product requirements document. The strategic narrative lives in [`ROADMAP.md`](../../ROADMAP.md); this page is the durable, structured summary engineers can build from.

## Problem

Building agents and apps for Microsoft 365 (Copilot, Teams, Outlook, the Microsoft 365 app) requires composing many independent systems: app manifests, Entra ID identities, Azure infrastructure, Bot Framework registration, Teams Developer Portal, Microsoft Graph publishing, and the M365 sideloading service. A new developer has to understand all of them before producing anything that runs.

## Solution shape

A single toolset (VS Code, Visual Studio, CLI) that:

1. **Scaffolds** opinionated, runnable starting points for every supported capability.
2. **Runs and debugs** via two modes: M365 Agents Playground (local simulator, no login or tunnel required) or sideloading to Teams / Microsoft 365 (requires dev tunnel and M365 tenant login).
3. **Provisions** the required Azure / Entra / Teams resources via declarative IaC.
4. **Deploys** code (zip, container, npm, dotnet) to the provisioned resources.
5. **Publishes** the resulting app to a tenant catalog or the M365 store.

A shared engine implements steps 1–5 once; each surface is a thin adapter.

## In scope (must)

- Day-1 support for new Microsoft 365 capabilities.
- TypeScript, JavaScript, Python, C# scaffolds. **Note:** VS Code language scaffolds live under `templates/vsc/{ts,js,python}/`; `templates/vsc/common/` holds language-agnostic templates shared across them. C# templates live separately under `templates/vs/csharp/` for the Visual Studio extension.
- All five lifecycle stages reachable from VS Code, Visual Studio, and CLI. v3 ships **4 driver-run stages** in `m365agents.yml` (`provision`, `deploy`, `publish`, `share`); the fifth ("run/debug") is local-debug tooling driven by `m365agents.local.yml` plus dev tunnels and the M365 Agents Playground.
- Reproducible non-interactive runs in CI (`CI_ENABLED=true`).
- Localised UX in 13 languages (see [`Localize/loc/`](../../Localize/loc/)).
- Telemetry for product improvement, with strict secret masking.

## In scope (should)

- AI-assisted authoring inside the editor (Copilot Chat). The user-visible `@m365agents` participant is **registered by a separate companion extension** (install via `aka.ms/install-m365agents`); this repo's VS Code extension contributes only the `office` participant plus link buttons + pre-filled queries. See [01-product/ux/surfaces/copilot-chat-participant.md](../01-product/ux/surfaces/copilot-chat-participant.md).
- E2E lifecycle tests covering shipped templates. Current coverage tracked in [capabilities-matrix.md](capabilities-matrix.md).

## Aspirational / not yet shipped

- **Containerised deploy targets** — listed in [`ROADMAP.md`](../../ROADMAP.md) but **no container or Docker driver exists** in either v3 (`packages/fx-core/src/component/driver/`) or v4 (`packages/core-next/src/drivers/builtin/`) at extraction time. Treat as roadmap, not shipped capability.

## Out of scope

- Hosting the user's deployed app (we provision Azure; we do not run it).
- End-user product documentation (lives at <https://aka.ms/teamsfx-docs>).
- Substituting for `az` / `azd` for non-M365 Azure workflows.

## Success criteria

See [success-metrics.md](success-metrics.md). At a glance:

- *Time to first run* — minutes from "install toolkit" to "agent responds in Playground".
- *Lifecycle reliability* — provision/deploy/publish success rate per template.
- *Coverage* — number of supported capabilities in the [capability matrix](capabilities-matrix.md).
- *Adoption* — telemetry events per active developer.

## Constraints

- **Backward compatibility.** Existing v3 projects must keep working unchanged. The v4 engine is opt-in via `TEAMSFX_V4_CORE`.
- **Microsoft platform contracts.** Manifest schemas, TDP, Graph, Bot Framework, MSAL — we are downstream of all of them.
- **Security baselines.** No PII in telemetry, EAFP filesystem access, Zip-Slip protection on extraction, magic-byte validation on uploads.
- **Open source.** MIT licensed; contributors must sign the CLA.

## Risks

| Risk | Mitigation |
|------|------------|
| Two engines (v3 and v4) drift in behaviour | The v4 engine is gated by the `TEAMSFX_V4_CORE` feature flag (defined in [`packages/core-next/src/featureFlags/flags.ts`](../../packages/core-next/src/featureFlags/flags.ts) — v4-side only; v3 is the default until the flag flips). Cross-engine integration tests cover both. |
| Manifest schema churn (1.x → 2.4 → next) | `@microsoft/app-manifest` is the single typed wrapper layer |
| Token cache compatibility across v3 / v4 | Both write to `~/.fx/account/`. v3 cli writes via [`packages/cli/src/commonlib/cacheAccess.ts`](../../packages/cli/src/commonlib/cacheAccess.ts) (uses `os.homedir() + '/.${ConfigFolderName}/account'`); v4 cli-next writes via `packages/cli-next/src/auth/cacheAccess.ts` to the same path. |
| Long-tail of one-off generators in fx-core | v4 replaces with explicit `TemplateRegistry` / `DriverRegistry` |
