# PR Review Template

Copy and use this template for your code reviews.

---

## Summary

[Brief overview of what was reviewed - 1-2 sentences]

**PR Size:** [Small/Medium/Large] (~X lines)
**Review Time:** [X minutes]

## Strengths

- [What was done well]
- [Good patterns or approaches used]
- [Improvements from previous code]

## Required Changes

🔴 **[blocking]** [Issue description]
> [Code location or example]
> [Suggested fix or explanation]

🔴 **[blocking]** [Issue description]
> [Details]

## Important Suggestions

🟡 **[important]** [Issue description]
> [Why this matters]
> [Suggested approach]

## Minor Suggestions

🟢 **[nit]** [Minor improvement suggestion]

💡 **[suggestion]** [Alternative approach to consider]

## Questions

❓ [Clarification needed about X]

❓ [Question about design decision Y]

## Security Considerations

- [ ] No hardcoded secrets
- [ ] Input validation present
- [ ] Authorization checks in place
- [ ] No SQL/XSS injection risks

## Test Coverage

- [ ] Unit tests added/updated
- [ ] Edge cases covered
- [ ] Error cases tested

## Verdict

**[ ] ✅ Approve** - Ready to merge
**[ ] 💬 Comment** - Minor suggestions, can merge
**[ ] 🔄 Request Changes** - Must address blocking issues

---

## Quick Copy Templates

### Blocking Issue
```
🔴 **[blocking]** [Title]

[Description of the issue]

**Location:** `file.ts:123`

**Suggested fix:**
\`\`\`typescript
// Your suggested code
\`\`\`
```

### Important Suggestion
```
🟡 **[important]** [Title]

[Why this is important]

**Consider:**
- Option A: [description]
- Option B: [description]
```

### Minor Suggestion
```
🟢 **[nit]** [Suggestion]

Not blocking, but consider [improvement].
```

### Praise
```
🎉 **[praise]** Great work on [specific thing]!

[Why this is good]
```

### Question
```
❓ **[question]** [Your question]

I'm curious about the decision to [X]. Could you explain [Y]?
```
