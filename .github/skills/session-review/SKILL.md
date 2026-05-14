---
name: session-review
description: "Summarize session changes and update instructions and plans. Use when: ending a work session, reviewing what was done, updating coding conventions after bugfixes, syncing instructions with code changes, updating plan status."
argument-hint: "Optionally describe what to focus on (e.g. 'AI agent changes' or 'all')"
---

# Session Review

## When to Use

- User says "review changes", "summarize session", "update instructions"
- After a batch of bugfixes or feature work
- Before ending a work session to capture lessons learned
- When instructions or plans may be stale after code changes

## Procedure

### Step 1 — Gather Changes

1. Review the conversation history to identify all files modified in this session
2. Categorize changes:
   - **Bugfixes**: What broke, what was the root cause, what was the fix
   - **Features**: What was added, what components are involved
   - **Refactors**: What was restructured and why
3. For each change, note the file path and a one-line summary

### Step 2 — Check Instruction Files

1. Read all instruction files in `.github/instructions/`
2. For each instruction file, check if any documented conventions are now **stale or inaccurate** due to session changes:
   - API parameter names changed?
   - Default behaviors changed (e.g., approval levels)?
   - New patterns established (e.g., routing style, component usage)?
   - New constraints discovered (e.g., model compatibility)?
3. Update instruction files to reflect the current state of the code
4. Add new sections only if a genuinely new convention was established

### Step 3 — Check Plan Files

1. Read plan files in `.dev/plans/` that relate to the session's work
2. Update task checkboxes to reflect completion status
3. Add a **Bugfixes** or **Changes** section if session work was unplanned fixes
4. Update the `Status` field (In Progress → Complete if all tasks done)
5. Add a log entry with today's date and summary

### Step 4 — Present Summary

Output a summary table to the user with:

| Column | Content |
|--------|---------|
| # | Sequential number |
| File | Link to changed file with line number |
| Change | One-line description of what was done |

Then list which instruction/plan files were updated and what was changed in each.

## Quality Checks

- [ ] Every code change is accounted for in the summary
- [ ] No instruction file contains stale information contradicted by current code
- [ ] Plan task lists match actual completion status
- [ ] New conventions are documented if they'll apply to future work
- [ ] Summary is concise — one line per change, no prose paragraphs

## Anti-patterns

- **Over-documenting**: Don't add one-off fixes as permanent conventions. Only document patterns that will recur.
- **Stale sensitivity tables**: If approval behavior changes, update the sensitivity level table — this is frequently missed.
- **Missing routing conventions**: Navigation patterns (absolute vs relative) are easy to forget documenting.
- **Forgetting API compat notes**: Model-specific parameter changes (like `max_tokens` → `max_completion_tokens`) should be in instructions so they aren't re-introduced.

## Related: Per-Turn Documentation Guard

For **continuous** staleness detection during implementation (not just at session end),
use the `docs-guard` skill. It checks after each code-modifying turn whether instructions,
features, or plans need updating and adds flagged items to the todo list. Session-review
then executes those flagged updates at the end of the session.

Workflow: `docs-guard` flags → todo list accumulates → `session-review` executes updates.
