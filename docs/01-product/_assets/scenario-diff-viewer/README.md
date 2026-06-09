# Scenario Diff Viewer

Standalone side-by-side viewer for comparing two rendered scenario HTML pages (or any two URLs). Tooling-only, not a product artifact.

## Usage

Open [`index.html`](index.html) with two URL parameters:

```text
index.html?left=<URL>&right=<URL>
```

Both iframes are rendered side by side. Use the inputs in the header to swap URLs without editing the address bar.

## Typical cases

| Case | Left (baseline) | Right (new) |
|---|---|---|
| Evolving draft (file existed before this PR) | raw.githack URL pinned to the base ref or previous commit | raw.githack URL pinned to the PR HEAD |
| Brand-new draft (no previous git version) | raw.githack URL of the live sibling it redesigns (`../<slug>.html`) | raw.githack URL of the draft (`draft/<slug>.html`) |

raw.githack URL pattern:

```text
https://raw.githack.com/<org>/<repo>/<ref>/<path>
```

`<ref>` can be a branch name, tag, or commit SHA. Pin to a commit SHA for stable review links that won't drift as the branch advances.

## Example

Comparing the live `create-da-with-mcp-server.html` against its draft on branch `zhiyou/mcp-da-dt`:

```text
index.html
  ?left=https://raw.githack.com/OfficeDev/microsoft-365-agents-toolkit/zhiyou/mcp-da-dt/docs/01-product/scenarios/da/create-da-with-mcp-server.html
  &right=https://raw.githack.com/OfficeDev/microsoft-365-agents-toolkit/zhiyou/mcp-da-dt/docs/01-product/scenarios/da/draft/create-da-with-mcp-server.html
```

To preview the viewer itself via raw.githack:

```text
https://raw.githack.com/<org>/<repo>/<ref>/docs/01-product/_assets/scenario-diff-viewer/index.html?left=...&right=...
```
