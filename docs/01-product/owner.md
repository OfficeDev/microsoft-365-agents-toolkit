# Product owner registry

## Purpose

- Reference this file from PRD and scenario metadata when assigning PM and Engineer owners.
- Keep human contacts discoverable without embedding ownership rules in every PRD or scenario.
- Use GitHub handles, ADO aliases, Teams aliases, or team aliases that humans can resolve.

Role values: `PM`, `Engineer`, `Design`, `Docs`, `Review`.

## Owners

| Owner ID | Role | Area | Contact | Notes |
| --- | --- | --- | --- | --- |
| summzhan | PM | All product areas | GitHub: @summzhan | Default PM owner for every PRD and scenario. |
| HuihuiWu-Microsoft | Engineer | Declarative Agent (DA) | GitHub: @HuihuiWu-Microsoft | Default Engineer owner for DA PRDs and scenarios. |
| Alive-Fish | Engineer | Declarative Agent (DA) | GitHub: @Alive-Fish | Default Engineer owner for DA PRDs and scenarios. |
| TBD-Engineer | Engineer | Unassigned engineering area | TBD | Replace before engineering handoff for non-DA areas. |

## Status guidance

- Use `TBD-Engineer` only when the engineering owner for a non-DA area is genuinely unknown; replace it before engineering handoff.
- For DA work, default Engineer owner is `HuihuiWu-Microsoft` and/or `Alive-Fish`.
- For PM, default to `summzhan` until a more specific PM owner is identified.
- Engineering ownership can be cross-checked with [`../../.github/CODEOWNERS`](../../.github/CODEOWNERS) for package-level code areas.