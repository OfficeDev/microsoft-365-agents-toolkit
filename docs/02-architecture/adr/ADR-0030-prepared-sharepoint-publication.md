# ADR-0030 - Prepared SharePoint publication through YAML lifecycle actions

- **Status:** Proposed
- **Date:** 2026-11-07
- **Source:** SharePoint Copilot component publication requirements.

## Context

The in-process publish client executes the project's configured YAML
`publish:` lifecycle. A SharePoint-backed app publication needs to derive two
immutable artifacts from a completed SPFx package and submit them in a strict
order: deploy and verify the SharePoint package, then submit the application
package. The existing `spfx/deploy` and `teamsApp/publishAppPackage` actions
must remain compatible for ordinary projects.

The flow crosses local package parsing, SharePoint catalog writes and Graph
submission. Environment outputs alone are not durable enough to reconcile a
process interruption between a remote write and the next lifecycle action.

## Decision

1. Add one opt-in `spfx/preparePublication` YAML action. It reads a completed
   package and the project's application package, validates both, extracts
   only the supported API plugin, produces separate staged artifacts, and
   records their SHA-256 lineage in a token-free, atomically updated journal.
2. Keep local SharePoint package inspection available through a lean public
   entry point that does not initialize resources, authentication, or network
   clients. The YAML preparation action may perform read-only tenant discovery
   required to resolve tenant-bound plugin values.
3. Bind tenant ID, verified catalog origin, app/product identity, source and
   prepared hashes, and a versioned transformation identity in one token-free
   journal. Identity-wide coordination spans preparation, deployment, and
   submission; changing source bytes or version does not bypass unresolved
   outcomes. Active legacy journals that lack tenant identity fail closed and
   remain available for explicit reconciliation.
4. Extend the existing SharePoint deployment and application submission
   actions with optional prepared-publication inputs. In that mode they verify
   the journal and exact bytes, persist stage transitions, require genuine
   approval before writes, and reject incompatible identity/version states.
   The ordinary action path is unchanged when prepared-publication inputs are
   absent.
5. Record an in-flight state before each non-idempotent remote write. A
   definitive rejection is distinct from an unknown outcome. Unknown outcomes
   and cancellation after a possible write remain durable and block automatic
   replay; the engine observes possible writes through a bounded settlement
   window and does not infer rollback or success from cancellation.
6. Use invocation-bound SharePoint catalog context for prepared requests,
   rather than mutable module-global base URL state. Verify the uploaded
   SharePoint package bytes and deployment state before
   allowing application submission. Submit the exact frozen application ZIP
   once; the prepared path does not use retry wrappers or create-to-update
   fallback.
7. Support either a declarative-agent action reference or the typed
   `AgentConnector.toolSource.plugin` representation. Preserve other
   application content and validate the final staged manifests against the
   supported schema machinery before freezing deterministic artifacts. The
   public lean package entry exposes an in-memory tenant-neutral assembly API
   backed by this same transformer. It returns exact source and output hashes,
   preserves tenant placeholders, performs local structural validation only,
   and leaves persistence to the caller. Tenant-bound preparation recognizes
   and replaces that deterministic managed contribution in place before final
   schema validation; conflicting or ambiguous ownership fails closed. No
   `TENANT_ORIGIN` CSP path is qualified yet; do not perform broad string
   replacement.
8. Do not change `FxCoreClient.publish`, route through the legacy Teams-only
   API, add a CLI action, or extend the direct-action allowlist.

## Consequences

- YAML ordering remains explicit and auditable; the application package cannot
  be submitted before SharePoint verification.
- The journal, rather than a second publisher or client callback, owns
  cross-action recovery state. Persisted lifecycle outputs expose the staged
  paths and hashes to later YAML actions but are not the recovery source of
  truth.
- The implementation supports only the qualified package shape and catalog
  operations it can verify. Unsupported clouds, malformed packages, ambiguous
  package content, collisions, and unresolved state fail closed.
- The public inspector is a local file operation. It never prompts, reads
  credentials, or makes service requests.
- The public in-memory assembler accepts application/SPPKG buffers through the
  lean entry and returns a neutral application ZIP plus hash/identity
  provenance. It does not initialize a lifecycle client, access the
  filesystem, authenticate, or make network requests. Its structural checks
  do not replace tenant-bound schema, identity, audience, catalog, approval,
  or deployment verification.
- The draft PR must retain the open package-format, service-verification,
  approval, and recovery questions for maintainer review before the decision
  is accepted.
