---
name: plan-tracker
description: "Save a structured plan and task list to the workspace before starting implementation. Use when: beginning a new feature, refactoring, bug fix, or multi-step task. Updates plan status during and after work."
argument-hint: "Describe the task or feature to plan"
---

# Plan Tracker

## When to Use

- Starting any multi-step implementation (feature, bug fix, refactor)
- User asks to plan before coding
- User says "track this", "make a plan", "save todos"

## Procedure

### Phase 1 — Create Plan (Before Implementation)

1. Create a new plan file at `.dev/plans/YYYY-MM-DD-short-title.md` (e.g., `.dev/plans/2026-03-20-auth-feature.md`)
2. Use this template:

```markdown
# Plan: [Short Title]

**Status:** In Progress
**Created:** [date]
**Updated:** [date]

## Goal

[One-sentence summary of the objective]

## Tasks

- [ ] Task 1
- [ ] Task 2
- [ ] Task 3

## Notes

- [Key decisions, constraints, or open questions]

## Log

- [date] — Plan created
```

3. Tasks should be feature-level (e.g., "Add user authentication endpoint"), not micro-steps. Only break a task into subtasks when it is genuinely complex:

```markdown
- [ ] Add user authentication endpoint
  - [ ] Define request/response schema
  - [ ] Implement token validation
```

4. Commit to the plan before writing any implementation code.

### Phase 2 — Track Progress (During Implementation)

- After completing each task, update the plan file:
  - Check off the completed task (`- [x]`)
  - Add a log entry with what was done
  - Update the `Updated` date
- If new tasks emerge, add them to the task list
- If the goal shifts, update the Goal section and log the change

### Phase 3 — Close Out (After Completion)

1. Mark all tasks as done in the plan file
2. Set `**Status:**` to `Complete`
3. Add a final log entry summarizing the outcome
4. If any follow-up work was identified, note it under a `## Follow-up` section

## Rules

- One plan per file — never merge plans into a single file
- Keep task descriptions short and actionable (start with a verb)
- Do not delete completed plan files — they serve as project history
- When starting work, check `.dev/plans/` for any existing in-progress plans and ask the user whether to continue or start fresh
