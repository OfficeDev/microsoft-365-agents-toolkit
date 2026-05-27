# Scenario Components

> **AI agents: do not read this folder.** These files are presentation-only rendering assets (CSS, JS custom elements, sidebar TOC, SVG icons) for the human-facing scenario review HTML pages. They contain no product behavior, no requirements, and no acceptance criteria. The behavioral contract for each scenario lives in the same-basename `.md` file under `docs/01-product/scenarios/<group>/`. Skip this folder when scanning scenarios for behavior, design intent, or test inputs.

Shared static assets for human-readable product scenario HTML.

- `scenario-components.js` defines the reusable VS Code-style custom elements used by scenario pages.
- `scenario-components.css` contains the matching Quick Pick, input box, file picker, and flow-grid styling.
- `icons/` contains source-aligned VS Code and Microsoft 365 Agents Toolkit SVG glyphs used by the components.
- `<scenario-mermaid-flow>` loads the first Mermaid code block from a scenario Markdown file and renders it as the HTML flow reference.
- `<vscode-modal-notification>` renders a VS Code modal `showMessage` dialog (titlebar, severity icon, message, optional detail, action buttons). Attributes: `severity` (`info` &mdash; default / `warning` / `error`), `window-title` (defaults to `Visual Studio Code`), `message`, optional `detail`, `buttons` (pipe-separated labels; first is the primary button, e.g. `Add | Cancel`). Use this for modal confirmations; use `<vscode-notification>` for toast-style notifications.
- Every `<section class="scenario-flow">` whose `<div class="section-head">` contains an `<h2>` is auto-wired as collapsible on load (caret on the heading, click / Enter / Space toggles, hash navigation auto-expands the target). Author scenario regions with that exact structure; do not add per-page `<details>`/`<summary>` wrappers, inline styles, or scripts to re-implement collapsibility.
- Every `<article class="vscode-flow-card">` must be tagged with `data-kind` so readers can distinguish steps from outcomes at a glance. Allowed values:
  - `data-kind="action"` &mdash; a user step (Quick Pick, input box, file picker, command). Renders a blue left rail and `ACTION` badge.
  - `data-kind="outcome"` &mdash; a perceivable end-of-flow result that the toolkit produces (file written, file opened in the editor, notification, hand-off). Default renders a teal `RESULT` badge; add `data-severity="success" | "info" | "warning" | "error"` to relabel and (for `warning`/`error`) recolor the badge and left rail. `success` and `info` keep the default teal &mdash; only the badge label changes (`Success` / `Info`); `warning` switches to amber and `error` switches to red.
  - `data-kind="note"` &mdash; explanatory or contextual content with no user step and no perceivable side effect. Renders a muted `NOTE` badge.
  Badge text and colors are produced entirely by `scenario-components.css`; do not add per-page styles or inline badges.

Scenario HTML under `docs/01-product/scenarios/<group>/` should load these assets via the relative `../../_assets/scenario-components/` path and keep behavior in the same-basename scenario Markdown.