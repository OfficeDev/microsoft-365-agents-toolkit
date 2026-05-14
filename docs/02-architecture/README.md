# 02 — Architecture

This section is structured along the **arc42** template. Each numbered file maps directly to an arc42 section so contributors familiar with the template feel at home.

## Pages

| arc42 # | File | Topic |
|---------|------|-------|
| 1 | [01-introduction-and-goals.md](01-introduction-and-goals.md) | Goals, top quality goals, stakeholders |
| 2 | [02-constraints.md](02-constraints.md) | Tech, organisational, compatibility constraints |
| 3 | [03-context-and-scope.md](03-context-and-scope.md) | External systems we depend on |
| 4 | [04-solution-strategy.md](04-solution-strategy.md) | Headline architectural decisions |
| 5 | [05-building-blocks.md](05-building-blocks.md) | Package decomposition (C4 L1–L2) |
| 6 | [06-runtime-views.md](06-runtime-views.md) | Provision / deploy / publish sequences |
| 7 | [07-deployment-view.md](07-deployment-view.md) | What infra the toolkit *itself* runs on (none — it's local), what it provisions |
| 8 | [08-crosscutting-concepts.md](08-crosscutting-concepts.md) | Result pattern, errors, telemetry, l10n, security, secret masking |
| 9 | [09-architecture-decisions/](09-architecture-decisions/README.md) | ADR log |
| 10 | [10-quality.md](10-quality.md) | Quality goals and scenarios |
| 11 | [11-risks.md](11-risks.md) | Known risks and technical debt |

For the *infrastructure the toolkit provisions on behalf of users*, see [03-infrastructure/](../03-infrastructure/README.md). This section is about the **toolkit's own** architecture.
