# ADR-0004 — TDP region state: singleton vs request-scoped

- **Status:** Proposed
- **Date:** 2026-05-27
- **Source:** [`../external-dependencies/identity-and-login.md` §3](../external-dependencies/identity-and-login.md#3-open-questions-candidates-for-adrs)

## Context

TDP is region-sharded (see fact-page §1.7): the toolkit must do an AuthSvc
handshake before its first TDP call for a user and use the returned regional
URL thereafter. Today the regional URL is stored as singleton state on the
TDP client module. That works for a single-user CLI / VS Code session; it
does not generalize to scenarios where the engine processes more than one
identity in the same process (multi-tenant orchestration, server / RPC host,
test runners), and it entangles the TDP client lifetime with global state.

## Options considered

- **A —** Status quo: module-scoped TDP client, region stashed as a singleton.
- **B —** TDP client per operation; region resolved (and cached) per user identity.
- **C —** TDP client per user / tenant; region carried on the client instance.

## Decision

(Pending.)

## Consequences

(Pending.)
