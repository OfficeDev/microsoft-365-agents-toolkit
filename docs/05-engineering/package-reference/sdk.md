# `packages/sdk` — `@microsoft/teamsfx`

App-side SDK consumed by the **user's deployed app** (not the toolkit itself).

## Role

Provides authentication helpers and Microsoft Graph integration for Teams apps:

- SSO token acquisition for Teams tabs.
- On-behalf-of (OBO) flow for backend-to-Graph.
- `MicrosoftGraphProvider` and `OnBehalfOfUserCredential`.
- App credential helpers for bots and Functions.

## Targets

- Browser (tabs)
- Node.js (bots, Azure Functions, app services)

Built with **rollup** to produce both CJS and ESM bundles plus `.d.ts` declarations.

## Why it lives in this monorepo

- Templates depend on stable SDK API contracts; co-locating allows atomic changes.
- Telemetry / logging conventions stay consistent with the toolkit's own.

## Versioning

Independent semver. SDK breaking changes are coordinated with template updates so a fresh scaffold always compiles with the SDK version it pins.

## Documentation

End-user SDK docs live at <https://aka.ms/teamsfx-docs>. This page is a pointer for contributors making SDK changes inside the monorepo.
