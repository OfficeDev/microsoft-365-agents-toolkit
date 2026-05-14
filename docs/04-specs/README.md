# Specs — Behavioral Contracts

This directory is the **authoritative source of truth** for system behavior in v4.
Tests are derived from specs. Code is constrained by tests. Specs cannot drift.

> **Docs explain the system. Specs constrain it.**
> When docs and specs conflict, specs win.

---

## Why Specs Exist

AI coding has three failure modes that specs directly address:

| Failure | Root cause | Spec fixes it via |
|---|---|---|
| AI changes the wrong things | Doesn't know code boundaries | `## Boundary` section |
| AI breaks hidden constraints | Doesn't know why something was designed a certain way | `## Invariants` section |
| AI can't self-validate | No executable criteria | `## Acceptance Criteria` → tests |

---

## Layer Hierarchy

```
Requirements (inform specs, can drift)
  docs/01-product/     — WHY: business goals, user needs (PM-owned)
  docs/01-product/ux/          — HOW users interact: flows, question model (UX-owned)
  GitHub Issue / ADO   — WHO/WHAT: atomic user story (PM + Dev)

Constraints (specs must respect these, cannot be overridden)
  docs/02-architecture/ — system structure, ADRs
  docs/03-infrastructure/ — platform boundaries, Azure limits

Contracts (source of truth, enforced by tests)
  docs/04-specs/domains/     — domain boundaries + inter-domain interfaces
  docs/04-specs/operations/  — WHAT EXACTLY each operation does
        ↓ derives
  Tests (packages/*/tests/) — EXECUTABLE SOURCE OF TRUTH
        ↓ constrains
  Code (packages/*/src/)    — IMPLEMENTATION
```

**Dependency rules:**
- Specs are constrained by architecture + infrastructure — specs cannot require behavior that contradicts an ADR or platform limit
- Specs are informed by PRD + UX — but these are inputs, not enforcement; they can drift
- If a spec conflicts with architecture → update the spec or file a new ADR
- If a spec cannot be written (ambiguity) → surface to PM, update PRD first
- One User Story → one Operation Spec
- One Operation Spec → one or more test files (never zero)

---

## Directory Structure

```
docs/04-specs/
  README.md                     ← this file
  domains/                      ← 7 domain specs (cross-package, cross-cutting)
    01-scaffolding.md
    02-lifecycle.md
    03-identity.md
    04-manifest.md
    05-extensibility.md
    06-environment.md
    07-collaboration.md
  operations/                   ← operation specs (one per atomic operation)
    scaffolding/
      new-project.md
      new-sample.md
    lifecycle/
      provision.md
      deploy.md
      publish.md
    manifest/
      validate.md
      package.md
    extensibility/
      add-action.md
      add-knowledge.md
    environment/
      add-environment.md
    identity/
      login.md
    collaboration/
      grant-permission.md
```

Operation specs are organized by domain, not by package.
A spec that touches multiple packages still lives in one file.

---

## Domain Spec Format

```markdown
# Domain: <Name>

## Purpose
One sentence: what this domain is responsible for.

## Boundary
**In scope:** what this domain owns
**Out of scope:** what other domains own (with references)

## Interfaces with other domains
| Consumer domain | What it consumes | Contract |
|---|---|---|
| Lifecycle | getToken(scope) → Token | Identity must provide a valid token or return AtkError |

## Invariants
Constraints that must never be violated by any operation in this domain:
- [ ] Invariant 1
- [ ] Invariant 2

## Operations
List of operation specs in this domain:
- [operation-name](../operations/<domain>/<operation>.md)
```

---

## Operation Spec Format

Spec review = test case review. The AC table is the only artifact humans need to review.
Each AC row maps 1:1 to a test case. Tests reference AC IDs in their names.

```markdown
# Operation: <name>

**Domain:** <domain>
**Issue:** #<number> — <one-line user story>
**Status:** draft | review | approved | implemented

## Purpose
One sentence: what this operation does.

## Inputs
| Parameter | Type | Required | Validation |
|---|---|---|---|
| name | string | yes | 1–120 chars |
| generateClientSecret | boolean | yes | — |

## Outputs
| Key | Type | Condition |
|---|---|---|
| clientId | UUID string | always |
| clientSecret | string | only when generateClientSecret=true |

## Acceptance Criteria

| ID | Scenario | Expected Result |
|---|---|---|
| AC-01 | Valid inputs, Entra API reachable | Returns clientId (UUID) and objectId (UUID) |
| AC-02 | generateClientSecret=false | Output does not include clientSecret |
| AC-03 | name exceeds 120 characters | UserError(AadAppNameTooLong) |
| AC-04 | Entra API unavailable | SystemError(AadAppCreateFailed) |
| AC-05 | Same existingClientId called twice | Returns same clientId (idempotent) |

## Boundary
This operation does NOT:
- Register OAuth connections (Identity domain)
- Modify manifest.json (Manifest domain)
- Create project files (Scaffolding domain)

## Invariants
- clientSecret must never appear in logs or telemetry
- Partial failure must not leave resources created (atomic)

## Error Cases
| Condition | Error type | Error name |
|---|---|---|
| name > 120 chars | UserError | AadAppNameTooLong |
| Entra API unavailable | SystemError | AadAppCreateFailed |

## Related
- Domain spec: [docs/04-specs/domains/03-identity.md](../domains/03-identity.md)
- ADR: [ADR-0004](../../02-architecture/09-architecture-decisions/0004-zod-driver-validation.md)
- Implementation: `packages/core-next/src/drivers/aadApp/create.ts`
- Tests: `packages/core-next/tests/unit/drivers/aadApp/create.test.ts`
```

Tests reference AC IDs directly:

```typescript
// AC-01: valid inputs
it("returns clientId and objectId on success", ...)

// AC-03: name > 120 chars
it("returns UserError(AadAppNameTooLong) when name exceeds limit", ...)

// AC-05: idempotency
it("returns same clientId when called with same existingClientId", ...)
```

---

## Complete Workflow

Three entry points, one flow:

```
Entry A: GitHub Issue
Entry B: ADO Work Item
Entry C: Chat (informal description)
         ↓
    [Chat only] AI classifies feature/bug/chore,
    identifies domain, creates Issue/WI
         ↓
AI reads Issue + Domain Spec + PRD/UX docs
         ↓
AI attempts to draft Operation Spec
    ├── Completes → proceed to review
    └── Blocked → AI surfaces specific questions:
                  "Cannot define AC-03 because:
                   should unreachable MCP server fail
                   immediately or retry? How many times?"
         ↓
── HUMAN GATE 1: Answer AI's questions ──
   (may trigger PRD update if scope is unclear)
         ↓
Spec completed (draft → review status)
         ↓
── HUMAN GATE 2: Review AC table ──
   One question per row: scenario clear? result correct?
   Approve → status: approved
         ↓
AI generates tests from AC table (all failing)
test names include AC IDs: "AC-01: ..."
         ↓
AI implements code until all tests pass
         ↓
AI updates affected docs/ files
         ↓
PR: spec + tests + code + docs in one PR
    references Issue/WI number
         ↓
Reviewer confirms AC ↔ implementation match
Merge → close Issue/WI
```

**Human intervention is only at the two gates.**
Everything else is AI-executed.

---

## Maintenance

| Who | When |
|---|---|
| Dev (before code) | Write or update spec when behavior changes |
| Dev | When PRD clarification resolves an ambiguity |
| AI (after code) | Update affected docs/, minor spec corrections |
| Never | Update spec to match wrong code — fix the code instead |

**Blocked spec = upstream ambiguity.** Do not guess.
Surface to PM → update PRD → complete spec.

---

## How AI Uses Specs

When given a task:

1. Find or create Operation Spec in `docs/04-specs/operations/<domain>/`
2. Read Domain Spec for boundary and invariants
3. Try to complete all sections — list specific questions for anything missing
4. On approval: generate tests using AC IDs as test name prefixes
5. Implement code until tests pass
6. Update spec status to `implemented`
7. Update affected `docs/` files (architecture docs, contributing guides)
