---
name: docs-guard
description: "Continuous documentation guard — checks if instructions, features, or plans need updating after each implementation turn. Use when: implementing features, fixing bugs, refactoring code, adding templates, changing conventions. Adds update tasks to todo list automatically. DO NOT USE FOR: pure research, reading files, answering questions, session review (use session-review skill instead)."
argument-hint: "Describe what was just changed (e.g. 'fixed scaffold pipeline' or 'added new template')"
---

# Documentation Guard

## When to Use

Run this check after every implementation turn that modifies code. Skip for read-only
turns (research, file reads, answering questions).

## Trigger Conditions

After any turn that **creates, edits, or deletes** files in the workspace, evaluate
whether the change impacts documented conventions, feature metadata, or tracked plans.

## Procedure

### Step 1 — Classify the Change

Determine which documentation categories the change could affect:

| Change Type | Check Instructions? | Check Features? | Check Plans? |
|-------------|-------------------|-----------------|--------------|
| New file/module added | Yes — package layout trees | Maybe — if it's a template | Yes — task completion |
| Function signature changed | Yes — if documented in conventions | No | Maybe |
| Template descriptor added/modified | Yes — descriptor counts, scaffold docs | Yes — feature table | Yes |
| Bug fix | Yes — if it reveals a wrong convention | No | Yes — task completion |
| New dependency added | Yes — if affects build/setup docs | No | Maybe |
| Error handling pattern changed | Yes — error conventions section | No | Maybe |
| CLI command added/changed | Yes — CLI integration section | Yes — if new capability | Yes |
| Test added/modified | Yes — test count if documented | No | Yes |
| Build/config changed | Yes — setup/build instructions | No | Maybe |

### Step 2 — Quick Staleness Scan

For each applicable category, do a **fast heuristic check** (no need to read full files):

**Instructions** (`.github/instructions/*.instructions.md`):
- Does the change contradict any **named constant, count, or path** in instructions?
- Does the change establish a **new pattern** that should be documented?
- Were **test counts** affected?
- Were **package layout trees** affected?

**Features** (`.github/instructions/features.instructions.md`):
- Was a **template descriptor** added, removed, or renamed?
- Did **language support** change for any template?
- Was a template **E2E-verified** for the first time?
- Did a **template name constant** change (e.g., folder name mapping)?

**Plans** (`.dev/plans/*.md`, session memory):
- Was a **tracked task** completed or unblocked?
- Did a **new blocker** emerge that should be noted?
- Did scope **expand or shrink**?

### Step 3 — Add to Todo List

If any staleness is detected, append documentation update tasks to the **existing todo list**
(do not replace existing items). Use these standard task titles:

| Detected Staleness | Todo Title |
|-------------------|------------|
| Instruction file has wrong count/path/name | `Update {filename} instructions` |
| Features table needs new row or name fix | `Update features.instructions.md` |
| Plan task completed | `Mark {task} complete in plan` |
| New convention established | `Document {pattern} in {filename}` |
| Test counts changed | `Update test counts in codebase.instructions.md` |

**Priority rules:**
- Documentation todos go **after** all implementation todos
- Group related doc updates into a single todo when possible (e.g., "Update instructions after scaffold fix" covers multiple files)
- Mark as `not-started` — never auto-complete doc todos

### Step 4 — Skip Silently When Clean

If no staleness is detected, **do nothing**. Do not mention the check to the user.
The guard should be invisible when documentation is up to date.

## Decision Shortcuts

**Skip entirely when:**
- The turn only read files or searched
- The turn only ran tests without changing code
- The change was to test files only (unless test counts are documented)
- The change was to a non-documented package (e.g., `simpleauth`, `dotnet-sdk`)

**Always check when:**
- Any file in `packages/core-next/src/templates/` changed
- Any file in `packages/cli-next/src/` changed
- Any `.instructions.md` file was edited (meta-check: did the edit make other instructions stale?)
- A `package.json` was modified (dependency or config changes)

## Quality Criteria

- [ ] No false positives — only flag genuinely stale documentation
- [ ] No noise — never mention the check if nothing needs updating
- [ ] Correct priority — doc todos always after implementation todos
- [ ] Specific titles — "Update fx-core.instructions.md scaffold section" not "Update docs"
- [ ] Deferred execution — only add to todo list, don't execute updates mid-implementation

## Anti-patterns

- **Updating docs mid-implementation**: The guard only *flags* — actual updates happen at the end via the `session-review` skill or explicit user request
- **Over-flagging**: Don't add a doc todo for every single file edit; batch related changes
- **Vague todos**: "Update instructions" is useless — specify which file and which section
- **Blocking implementation**: Never pause implementation work to update documentation
