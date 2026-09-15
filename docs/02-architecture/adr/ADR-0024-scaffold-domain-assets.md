# ADR-0024 - Scaffold domain assets and expression ownership

- **Status:** Accepted
- **Date:** 2026-09-09
- **Approval:** Chat authorization to implement the three remaining high-priority refactors.
- **Extends:** [ADR-0016](ADR-0016-declarative-template-format.md) and [ADR-0017](ADR-0017-named-pipeline-step-whitelist.md).

## Context

The expression whitelist mixes registration with MCP and Office implementations.
Dynamic OpenAPI and MetaOS generators embed static code and manifest content in
TypeScript. OpenAPI authentication insertion depends on an English YAML comment.
These are ownership and maintenance problems, not new product requirements.
No PRD or scenario flow change is required.

## Decision

1. Keep the closed expression whitelist as explicit registration and compatibility
   exports. Generic, MCP, and Office functions have separate implementation modules.
   Existing names, signatures, defaults, and byte output remain compatible.
2. Store engine-owned static generation fragments in dedicated JSON assets beside
   their domain implementation. Code fragments use arrays of lines; structured
   manifests stay JSON objects. Static imports let TypeScript and webpack package
   the assets without runtime filesystem discovery or a copy-build step.
   Ordinary new files still belong in template `content/`; these assets support
   operation-dependent generation and mutations of existing files only.
3. Domain functions bind dynamic values and perform mutations. Assets do not
   execute code or introduce an author-facing template dialect. Existing code
   interpolation semantics and public step parameters remain unchanged.
4. OpenAPI authentication actions are inserted into the parsed YAML `provision`
   sequence before the first `teamsApp/zipAppPackage` action, or at its end if no
   such action exists. Other top-level sections and comments are preserved.
   Invalid YAML or a missing/non-sequence `provision` is an explicit step error;
   no malformed fallback append is allowed. No registrations is a byte-preserving
   no-op. Missing optional YAML files retain their existing no-op behavior.

## Alternatives

- Move literals into another TypeScript file: easy, but content still requires
  editing executable source and does not establish a data ownership boundary.
- Add a template-channel fragment API: independently versionable, but requires
  archive, schema, parameter, and compatibility changes outside this refactor.
- Parse and serialize the whole YAML object: simpler, but discards comments.
  The existing `yaml` document API preserves document structure and comments.

## Verification

- [RCTX-14](../../03-specs/operations/scaffolding/build-render-context.md):
  module ownership, whitelist identity, existing expression/Office scenarios.
- [AC-30 and AC-31](../../03-specs/operations/scaffolding/run-scaffold-pipeline.md):
  asset-backed output compatibility and comment-independent YAML insertion.
- Focused tests after each edit, source and scenario-tooling builds, affected lint
  and formatting, full fx-core coverage gate, and independent code review.

## Boundary

No new functions, steps, question semantics, template version floor, registry
unification, Office input alias changes, telemetry changes, or removal of legacy
OpenAPI compatibility. YAML whitespace may normalize; generated source and
manifest output must not change. Protocol constraints remain in domain code.
