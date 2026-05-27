# Product Assets (Rendering Only)

> **AI agents: do not read this folder.** Everything under `docs/01-product/_assets/` is presentation-only static rendering for the human-facing product review and scenario review HTML pages — stylesheets, JavaScript custom elements, sidebar TOC, SVG icons, the Markdown/Mermaid preview tool. It contains no product behavior, no requirements, no acceptance criteria, and no design intent.
>
> Product behavior lives in the Markdown files under [`docs/01-product/`](../README.md) (PRDs, scenarios, capabilities) and the specs under [`docs/04-specs/`](../../04-specs/README.md). When scanning the product folder for behavior, contracts, flows, or test inputs, skip this folder entirely.

## Folders

- [`product-review/`](./product-review/) — page-level layout stylesheet shared by [`docs/01-product/scenarios/index.html`](../scenarios/index.html) and the scenario review HTML pages under [`docs/01-product/scenarios/<group>/*.html`](../scenarios/README.md).
- [`product-artifact-viewer/`](./product-artifact-viewer/README.md) — standalone HTML tool that renders any Markdown artifact (with Mermaid) for human review. Tooling-only.
- [`scenario-diff-viewer/`](./scenario-diff-viewer/README.md) — standalone side-by-side viewer for two rendered scenario HTML pages (live vs draft, base vs HEAD). Tooling-only.
- [`scenario-components/`](./scenario-components/README.md) — reusable VS Code-style flow components (`<vscode-single-select>`, `<vscode-multi-select>`, `<vscode-input-box>`, `<vscode-file-select>`, `<vscode-codelens-file>`, `<vscode-notification>`, `<vscode-modal-notification>`, `<scenario-mermaid-flow>`) plus their stylesheet, icons, and the auto-injected sidebar TOC.

## Why the `_` prefix

The leading underscore signals "internal/supporting; not a behavioral entry point". Static-site generators and several doc tools use the same convention. Product, scenario, and spec authors can ignore this folder unless they are intentionally editing the shared look-and-feel or the preview tool.
