# Code Review Best Practices

Comprehensive guidelines for conducting effective code reviews.

## Review Philosophy

### Goals of Code Review

**Primary Goals:**
- Catch bugs and edge cases before production
- Ensure code maintainability and readability
- Share knowledge across the team
- Enforce coding standards consistently
- Improve design and architecture decisions

**Secondary Goals:**
- Mentor junior developers
- Build team culture and trust
- Document design decisions through discussions

### What Code Review is NOT

- A gatekeeping mechanism to block progress
- An opportunity to show off knowledge
- A place to nitpick formatting (use linters)
- A way to rewrite code to personal preference

## Review Timing

### When to Review

| Trigger | Action |
|---------|--------|
| PR opened | Review within 24 hours, ideally same day |
| Changes requested | Re-review within 4 hours |
| Blocking issue found | Communicate immediately |

### Time Allocation

- **Small PR (<100 lines)**: 10-15 minutes
- **Medium PR (100-400 lines)**: 20-40 minutes
- **Large PR (>400 lines)**: Request to split, or 60+ minutes

## Review Depth Levels

### Level 1: Skim Review (5 minutes)
- Check PR description and linked issues
- Verify CI/CD status
- Look at file changes overview
- Identify if deeper review needed

### Level 2: Standard Review (20-30 minutes)
- Full code walkthrough
- Logic verification
- Test coverage check
- Security scan

### Level 3: Deep Review (60+ minutes)
- Architecture evaluation
- Performance analysis
- Security audit
- Edge case exploration

## Communication Guidelines

### Tone and Language

**Use collaborative language:**
- "What do you think about..." instead of "You should..."
- "Could we consider..." instead of "This is wrong"
- "I'm curious about..." instead of "Why didn't you..."

**Be specific and actionable:**
- Include code examples when suggesting changes
- Link to documentation or past discussions
- Explain the "why" behind suggestions

### Handling Disagreements

1. **Seek to understand**: Ask clarifying questions
2. **Acknowledge valid points**: Show you've considered their perspective
3. **Provide data**: Use benchmarks, docs, or examples
4. **Escalate if needed**: Involve senior dev or architect
5. **Know when to let go**: Not every hill is worth dying on

## Review Prioritization

### Must Fix (Blocking)
- Security vulnerabilities
- Data corruption risks
- Breaking changes without migration
- Critical performance issues
- Missing error handling for user-facing features

### Should Fix (Important)
- Test coverage gaps
- Moderate performance concerns
- Code duplication
- Unclear naming or structure
- Missing documentation for complex logic

### Nice to Have (Non-blocking)
- Style preferences beyond linting
- Minor optimizations
- Additional test cases
- Documentation improvements

## Anti-Patterns to Avoid

### Reviewer Anti-Patterns
- **Rubber stamping**: Approving without actually reviewing
- **Bike shedding**: Debating trivial details extensively
- **Scope creep**: "While you're at it, can you also..."
- **Ghosting**: Requesting changes then disappearing
- **Perfectionism**: Blocking for minor style preferences

### Author Anti-Patterns
- **Mega PRs**: Submitting 1000+ line changes
- **No context**: Missing PR description or linked issues
- **Defensive responses**: Arguing every suggestion
- **Silent updates**: Making changes without responding to comments

## Metrics and Improvement

### Track These Metrics
- Time to first review
- Review cycle time
- Number of review rounds
- Defect escape rate
- Review coverage percentage

### Continuous Improvement
- Hold retrospectives on review process
- Share learnings from escaped bugs
- Update checklists based on common issues
- Celebrate good reviews and catches
