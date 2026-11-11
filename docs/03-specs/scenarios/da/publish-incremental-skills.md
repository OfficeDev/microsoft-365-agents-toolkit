# Publish with incremental skills

- Status: ATK-only skip-copy with root/DA skill support implemented and L1 tests passing; live validation pending
- Domain: M365 package upload
- Scenario ID: SCN-DA-PUBLISH-INCREMENTAL-SKILLS
- Product scenario: [publish with incremental skills](../../../01-product/scenarios/da/publish-incremental-skills.md)

## Acceptance Criteria

| ID | Runtime | Purpose | Gate | Harness | Expected outcome |
|---|---|---|---|---|---|
| SCN-ISU-01 | L1 | scenario | required | real ZIP and service boundary fakes | First publish stages every local file and submits a reference-only in-memory package, preserving the original. |
| SCN-ISU-02 | L1 | scenario | required | retained-draft sync | Retained copy-free drafts skip hash-verified files; edited/new files upload and stale draft files are removed. Missing hashes cause reupload without downloads. |
| SCN-ISU-03 | L1 | scenario | required | failed publish | Partial synchronization cannot submit the package; no automatic reset, directory/package replay or disk artifacts after file edits. |
| SCN-ISU-04 | L2 | surface | deferred | authorized CLI tenant | CLI install and lifecycle publish produce usable remote skills with existing outputs. |
| SCN-ISU-05 | L3 | surface | deferred | VS Code lifecycle | Opt-in publish reports actionable errors or existing successful publication outputs. |
| SCN-ISU-06 | L1 | scenario | required | real engine + HTTP boundaries | Fresh drafts after publication upload all local files; interrupted file uploads resume on a later user-initiated run using the retained copy-free draft without reset. |
| SCN-ISU-07 | L1 | surface | required | actual lifecycle drivers | First uploads work without UI in noninteractive mode; both lifecycle drivers and CLI install retain the shared engine with no recovery prompt. |
| SCN-ISU-08 | L1 | scenario | required | preparation boundaries | Unexpected copy monitors, invalid local data or malformed directory responses stop before file writes or package submission. No polling or draft deletion occurs. |
| SCN-ISU-09 | L1 | scenario | required | root-only/mixed ZIP and real service boundaries | With the shared flag enabled, root and DA references are synchronized and rewritten in memory, shared folders upload once, and original package bytes and declaration settings remain unchanged. Disabled uploads retain the original path. |

## Composed operations

The required L1 workflows run in
[incrementalSkillUpload.test.ts](../../../../packages/fx-core/tests/component/m365/skills/incrementalSkillUpload.test.ts).
These tests compose the real engine with HTTP/token/filesystem boundaries faked;
they are not live-service E2E evidence. The L2/L3 rows remain deferred until the
test account has Graph/SPE authorization.

[Incremental skill upload](../../operations/m365/incremental-skill-upload.md) owns
metadata comparison, retry, identity, draft and transport mechanics. Existing MOS
package submission owns RAI validation and share-link retrieval.

## Flow

See the [product flow](../../../01-product/scenarios/da/publish-incremental-skills.md#flow).
The package is submitted only after every referenced skill synchronizes successfully.

## Boundary

No new upload commands/store submission, runtime guarantees without live evidence,
permission grants, durable monitor storage, copy polling, draft recreation,
older copy-enabled ATK drafts, other-client drafts, concurrent writers, or cross-process
transaction/rollback guarantees. An ambiguous final publication is not recoverable
by resetting drafts and republishing.
