# Incremental skill upload

## Metadata

- Status: approved
- Created: 2026-10-10T06:29:18Z
- Last updated: 2026-10-15T06:06:39Z
- PM owner: Yan Jin
- Engineer owner: HuihuiWu-Microsoft
- Owner source: PM owner confirmed in chat on 2026-10-12; engineer owner from ../owner.md.
- Related request: Initial plan approved on 2026-10-10; MOS recovery guidance on 2026-10-14; ATK-only skip-copy scope selected on 2026-10-15.

## Problem

Bulk agent-package ingestion can throttle when MOS uploads many skill files to
SharePoint Embedded (SPE). Developers need a bounded, repeatable publish path
without changing their portable local package.

## Goals

Avoid bulk MOS skill ingestion and handle throttling without submitting an
incomplete agent package. Avoid repeated writes where a retained draft already
contains matching files. Implement and test the client logic before first-party
Graph/SPE onboarding completes; do not represent offline tests as live interoperability.

## Non-goals

Raising platform limits, changing store submission, adding new upload commands, granting
permissions, remote-content downloads, disk-backed intermediate artifacts, or a
persistent local synchronization cache.
Durable copy-monitor storage, concurrent-client draft merging, adoption of drafts
created with published-content copying enabled, and destructive draft recovery
are excluded from this milestone.

## Users and scenario map

Developers sideloading or publishing packages with standalone/root-manifest skills,
declarative-agent skills, or both through the existing CLI
and lifecycle actions. See [publish with incremental skills](../scenarios/da/publish-incremental-skills.md).
Supported skill drafts must have been created exclusively by the new ATK
skip-copy flow. Using only ATK today is insufficient if an older client previously
created the draft with copying enabled.

## Requirements

| ID | Requirement |
|---|---|
| REQ-01 | Default-off opt-in; existing uploads remain unchanged when disabled or no local skills exist. |
| REQ-02 | Within the ATK-only skip-copy scope, upload all local files into new empty drafts. For retained drafts, compare remote metadata hashes, upload new/changed files and remove stale files only inside that draft. Without a usable hash, upload conservatively; never download remote content. |
| REQ-03 | Keep derived inventories, manifests, metadata and transport ZIP in memory. Never rewrite the original package or create intermediate files. |
| REQ-04 | Use sequential Graph operations with at most five retries, honoring Retry-After. File synchronization failures stop publication and retain the draft. Never fall back to uploading skill payloads in the MOS package or automatically replay ambiguous directory creation or package publication. |
| REQ-05 | Apply consistently to CLI install, teamsApp/extendToM365 and copilotAgent/publish; preserve scope, sharing, RAI validation and existing certificate checks. |
| REQ-06 | Request no published-content copy on every skill-directory creation. Support only single-writer agents whose skill drafts are created exclusively through this ATK skip-copy flow. The complete local skill is authoritative. Drafts created by other clients or older copy-enabled ATK flows are outside the supported scope. |
| REQ-07 | First uploads and subsequent uploads within scope need no draft-reset confirmation and support noninteractive execution. A later user-initiated run may synchronize a retained copy-free draft after interrupted preparation or file upload. Do not delete/recreate draft roots in this milestone; unexpected reported copying stops before file writes or package submission with an actionable scope error. |
| REQ-08 | Empty monitor URLs are not general proof of readiness. Proceeding relies on the documented exclusive skip-copy provenance and deployed MOS support, neither of which ATK can establish from an empty response. Do not persist or poll monitor URLs, forward credentials to them, reset published versions, or replay an ambiguously completed package publication. |
| REQ-09 | The same default-off incremental-upload flag covers root-manifest `agentSkills` and declarative-agent `agent_skills` / supported preview `x-agent_skills`. Root-only packages need no declarative agent. Mixed packages synchronize each distinct local skill folder once and update all its references only in the in-memory transport package. Preserve declaration-specific settings. With the flag disabled, submit the original package through the existing path without incremental validation or Graph requests. |

## Success metrics

Hash-verified no-op synchronization against retained copy-free drafts performs
zero file-content writes. First publication and updates that create fresh drafts
upload all local files; zero-write updates across publications are not promised.
Synchronization has
one active Graph operation, no remote-content downloads and no local artifact writes.
Failed synchronization makes zero final package submissions.
All uploads perform zero MOS draft deletions and require no recovery prompt.

## Constraints and risks

Graph/SPE onboarding and deployed MOS support for copy skipping are external
prerequisites. An older service may ignore the flag; empty monitor URLs do not
acknowledge support. Finalization can still throttle server-side. Retained partial
drafts are not a transaction. Missing hashes may cause unchanged files to be
uploaded again. Fresh drafts require full local uploads after each publication.
Users must not edit these agents through other clients or run concurrent uploads.
The API cannot distinguish a safe ATK-created draft from an old copying draft with
missing monitors; the opt-in scope is an operational prerequisite, not an enforced
client-identity or provenance guarantee.

## Open questions

The user confirmed root-manifest skill support on 2026-10-15. This corrects the
earlier accidental DA-only scope; it does not widen the ATK-only draft-provenance
prerequisite or add a separate feature flag.

The user selected the ATK-only skip-copy scope on 2026-10-15. This revision
supersedes the unimplemented confirmation/reset proposal from 2026-10-14;
cross-client adoption and MOS's delete/recreate recovery path are deferred.
MOS guidance supplied in chat provides the copy-skipping contract; the inspected
checkout predates that flag. Deployment support must be verified before live
testing. Persisted copy progress is tracked by MOS design work item 12372190.
Live validation and PM release approval remain outstanding.

The narrowed flow and failure behavior were approved in chat on 2026-10-15.
The [operation spec](../../03-specs/operations/m365/incremental-skill-upload.md)
records the approved engineering criteria and implementation status. Approval
does not assert live interoperability.
