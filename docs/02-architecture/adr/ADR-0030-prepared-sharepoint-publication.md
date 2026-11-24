# ADR-0030 - Prepared SharePoint publication through YAML lifecycle actions

- **Status:** Proposed
- **Date:** 2026-11-07
- **Source:** SharePoint Copilot component publication requirements.

## Context

The in-process publish client executes the project's configured YAML
`publish:` lifecycle. A SharePoint-backed app publication needs to derive two
immutable artifacts from a completed SPFx package and submit them in a strict
order: deploy and verify the SharePoint package, then submit the application
package. The existing `spfx/deploy` and `copilotAgent/publish` actions must
remain compatible for ordinary projects.

The flow crosses local package parsing, SharePoint catalog writes and a
Microsoft 365 Titles service (MOS3 Builder API) submission. Agent and connector
applications are published to MOS3, not to the Microsoft Graph Teams catalog,
so submission uses `copilotAgent/publish` and leaves `teamsApp/publishAppPackage`
untouched. Environment outputs alone are not durable enough to reconcile a
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
   journal. Identify the source application package by a canonical digest of its
   entry names and bytes rather than by its ZIP bytes, because the YAML
   lifecycle rebuilds that ZIP (with a new timestamp) on every run; the
   `.sppkg` source and both staged artifacts stay bound by the SHA-256 of their
   exact bytes. Only a lineage staged under the current transformation identity
   resumes: an unsubmitted lineage staged under an earlier one blocks new
   preparation until it is reconciled, because its bytes may differ from what
   the current transformation would stage, while a submitted lineage still
   replays. Identity-wide coordination spans preparation, deployment, and
   submission; changing source content or version does not bypass unresolved
   outcomes. Active legacy journals that lack tenant identity fail closed and
   remain available for explicit reconciliation.
4. Extend the existing SharePoint deployment and `copilotAgent/publish` actions
   with optional prepared-publication inputs. In that mode they verify
   the journal and exact bytes, persist stage transitions, require explicit
   approval before writes, and reject incompatible identity/version states.
   Approval is the positive answer to a modal prompt on the host's
   `UserInteraction`; a host that provides none fails with `ApprovalUnavailable`.
   Each action re-verifies the signed-in token tenant (and, for Graph and
   SharePoint, audience) and the discovered catalog against the journal before
   any prompt, read, or write. Submission needs no Graph token: it binds the
   SharePoint and MOS3 tokens to the journal tenant. The ordinary action path is
   unchanged when prepared-publication inputs are absent, and
   `teamsApp/publishAppPackage` is not extended.
5. Record an in-flight state before each non-idempotent remote write. A
   definitive rejection is distinct from an unknown outcome. Unknown outcomes
   and cancellation after a possible write remain durable and block automatic
   replay; the engine observes possible writes through a bounded settlement
   window and does not infer rollback or success from cancellation.
6. Use invocation-bound SharePoint catalog context for prepared requests,
   rather than mutable module-global base URL state. Verify the uploaded
   SharePoint package bytes and deployment state before
   allowing application submission. Submit the exact frozen application ZIP
   once to MOS3 and poll its status within a bounded window; the prepared path
   does not use retry wrappers, create-to-update fallback, or incremental skill
   rewriting, so the hash-bound bytes are what the service receives. A
   confirmed upload is persisted before any read-only share-link lookup.
7. Support either a declarative-agent action reference or the typed
   `AgentConnector.toolSource.plugin` representation. The connector form is
   emitted only when the application's manifest schema defines
   `ToolSource.plugin` (currently `devPreview`); GA schemas through v1.30 are
   closed to `remoteMcpServer`, so connector-only applications on those
   versions are refused with a typed error rather than rewritten, converted to
   `remoteMcpServer`, or given a synthetic agent. Preserve other
   application content and validate the final staged manifests against the
   supported schema machinery before freezing deterministic artifacts. The
   public lean package entry exposes an in-memory tenant-neutral assembly API
   backed by this same transformer. It returns the output hash and the source
   identities that preparation records, preserves tenant placeholders, performs
   local structural validation only, and leaves persistence to the caller.
   Tenant-bound preparation recognizes
   and replaces that deterministic managed contribution in place before final
   schema validation; conflicting or ambiguous ownership fails closed. Resolve
   only the managed endpoint and the `{{TENANT_ORIGIN}}` entries of the
   `frameDomains` and `resourceDomains` lists in a tool's `_meta["ui/csp"]`,
   both from the verified catalog origin. Refuse, instead of publishing
   literally, any other string that is entirely a placeholder, including the
   managed authentication reference that this operation does not provision, and
   any plugin nested too deeply to serialize safely. A placeholder inside longer
   text stays as it is, and no broad string replacement is performed.
8. Do not change `FxCoreClient.publish`, route through the legacy Teams-only
   API, change `teamsApp/publishAppPackage` or the Graph catalog client, add a
   CLI action, or extend the direct-action allowlist.

## Consequences

- YAML ordering remains explicit and auditable; the application package cannot
  be submitted before SharePoint verification.
- The journal, rather than a second publisher or client callback, owns
  cross-action recovery state. Persisted lifecycle outputs expose the staged
  paths and hashes to later YAML actions but are not the recovery source of
  truth. The MOS3 title and app IDs and the scope are recorded in the journal
  when the submission is confirmed.
- The post-publish completion link of the in-process coordinator follows the
  action: a lifecycle containing `copilotAgent/publish` opens the Microsoft 365
  admin center.
- Approval is only as strong as the host's `UserInteraction`: the toolkit cannot
  verify that a person answered, and the ATK CLI confirms without asking when
  it runs non-interactively. `atk publish` runs `teamsapp publish` rather than
  the YAML `publish:` lifecycle, so the CLI is unchanged and in-process hosts
  that call `FxCoreClient.publish` own the approval surface.
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
