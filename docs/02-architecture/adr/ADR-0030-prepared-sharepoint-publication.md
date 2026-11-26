# ADR-0030 - Prepare SharePoint-backed application packages through YAML actions

- **Status:** Proposed
- **Date:** 2026-11-07
- **Source:** SharePoint Copilot component publication requirements.

## Context

The in-process publish client executes the project's configured YAML
`publish:` lifecycle. A SharePoint-backed application needs a separate
SharePoint package without the embedded standalone-agent ZIP and an
application package that uses the project's own identity. The existing
`spfx/deploy` and `copilotAgent/publish` actions already provide the required
remote operations and ordering.

## Decision

1. Add an opt-in `spfx/preparePublication` action. It reads a completed SPPKG
   and the project's application ZIP, extracts the supported API plugin,
   resolves qualified tenant values, creates a sanitized SharePoint deployment
   copy and an assembled application ZIP, and returns their project-relative
   paths.
2. Keep the original SPPKG and authored application package unchanged. The
   preparation action writes only separate staged artifacts; it does not
   persist a publication journal or claim durable recovery semantics.
3. Add an optional `packagePath` input to `spfx/deploy`. When present, it
   overrides the archive selected by `package-solution.json`; the solution
   identity still comes from `package-solution.json`. Existing catalog
   discovery/creation, overwrite, upload, and deployment behavior is reused.
4. Pass the staged application ZIP path to the ordinary
   `copilotAgent/publish` action. It keeps its existing MOS3 submission,
   retry, polling, and output behavior. No prepared-only mode, approval
   contract, journal, or replay guard is added.
5. Express SharePoint-first ordering in project YAML. The lifecycle does not
   execute the later application publish action when SharePoint deployment
   fails.
6. Keep local package inspection and in-memory tenant-neutral assembly
   available through the lean public package entry. Neither operation
   initializes authentication, network clients, or lifecycle services.
7. Do not change `FxCoreClient.publish`, the CLI, `teamsApp/publishAppPackage`,
   or the Graph Teams catalog client.

## Consequences

- Reusing existing actions substantially reduces duplicate SharePoint and
  MOS3 transport code and preserves established behavior for ordinary
  projects.
- SharePoint upload and MOS3 submission do not gain hash-bound requests,
  approval prompts, durable journals, bounded settlement, or special
  uncertain-outcome handling. Existing action retry behavior remains in
  effect. A subsequent lifecycle run prepares new artifacts and can repeat a
  remote operation.
- The `packagePath` override is optional, so existing deploy YAML remains
  compatible. Its use changes only which archive is uploaded.
- YAML order remains explicit. SharePoint action success is the prerequisite
  for the MOS3 action to run, but no additional read-back verification is
  promised.
- The public in-memory assembler remains tenant-neutral and side-effect-free.
  Its output is a supported input to tenant-bound preparation.
- The draft PR retains package-producer, catalog/cloud, and submission
  retry/recovery questions for maintainer review before the decision is
  accepted.
