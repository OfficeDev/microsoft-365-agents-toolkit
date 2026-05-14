# Docs style guide

How to write for this site.

## Audience

Engineers who will read code immediately after reading the doc. Not end users — those go to <https://aka.ms/teamsfx-docs>.

## One purpose per page

Diátaxis quadrant:

| Quadrant | Purpose | Example |
|----------|---------|---------|
| **Tutorial** | Take a beginner from zero to a result | (rare here; most belong on learn.microsoft.com) |
| **How-to** | Solve a specific problem | [adding-a-driver.md](adding-a-driver.md) |
| **Reference** | Describe the shape of something | [package-reference/core-next.md](../05-engineering/package-reference/core-next.md) |
| **Explanation** | Build understanding | [02-architecture/04-solution-strategy.md](../02-architecture/04-solution-strategy.md), ADRs |

If a page is doing two of these, split it.

## Structure

- Start with a one-paragraph summary.
- Show the data / code / sequence as early as possible.
- End with pointers to deeper material.

## Linking conventions

- Workspace-relative paths only — no `file://` or `vscode://`.
- Use `[display](path)` form. Display text matches the path or describes the destination.
- Code symbols in backticks: `FxCore`, `defineOperation()`.
- File references **not** in backticks — use markdown links (this site's policy mirrors VS Code chat's policy).

## Diagrams

Use **Mermaid**. Keep diagrams short — if it doesn't fit on one screen, split it. Pair every diagram with a one-line text description.

```mermaid
flowchart LR
    A --> B
```

Source `.mmd` for non-trivial diagrams: `docs/assets/diagrams/<name>.mmd`. Inline mermaid in pages is fine for short ones.

## Tables vs prose

Use tables when the structure is regular (config keys, fields, comparisons). Use prose when there's a narrative.

## Voice

- Active. "The driver returns…" not "A driver is returned by…".
- Present tense. "The lifecycle parses YAML…" not "will parse".
- Imperative for instructions. "Add a `TemplateDescriptor`."

## Terminology

Use the [glossary](../00-overview/glossary.md). When a term has v3 / v4 split, mark it.

## Code samples

- Prefer real signatures over invented ones.
- Show types explicitly.
- Don't invent imports — match what the code actually uses.
- Prefer `// good` / `// bad` comments over long explanations.

## Update cadence

This site explains the codebase. When the codebase changes, update relevant pages in the same PR. The [`docs-guard`](../../.github/skills/docs-guard/SKILL.md) skill triggers on implementation turns and adds doc-update tasks automatically.

## What lives elsewhere

- Coding conventions enforced by Copilot — `.github/instructions/*.instructions.md`.
- Workflow skills — `.github/skills/*/SKILL.md` and `.agents/skills/*/SKILL.md`.
- Roadmap — `ROADMAP.md`.
- Release notes — per-package `CHANGELOG.md` and the rolled-up [`CHANGELOG.md`](../../CHANGELOG.md).

This site links to those — it does not duplicate them.
