# Diagrams

Source `.mmd` files for non-trivial Mermaid diagrams used across this site. Inline diagrams (a few lines, used in only one page) live with their page.

Convention:

- One file per diagram.
- Filename = `<topic>-<kind>.mmd`. Kind ∈ `c4-context`, `c4-container`, `sequence`, `flow`, `er`.
- Reference from the page using:

````markdown
```mermaid
%% include: assets/diagrams/<filename>.mmd
```
````

…or paste the contents inline if the docs build doesn't support includes.

## Catalogue

This directory is initially empty. Populate as new diagrams are needed and copy from inline `mermaid` code blocks in the prose pages.

Suggested first additions:

- `c4-context.mmd` — system context (extract from [02-architecture/03-context-and-scope.md](../../02-architecture/03-context-and-scope.md))
- `c4-container-v3.mmd` and `c4-container-v4.mmd` — building blocks per engine
- `provision-sequence.mmd` — provision flow
- `deploy-sequence.mmd` — deploy flow
- `publish-sequence.mmd` — publish flow
- `data-model-er.mmd` — data-model overview ER
