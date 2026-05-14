# Accessibility

The toolkit must be usable with assistive technologies on every surface.

## VS Code

- Tree view items have aria-labels via VS Code's standard `TreeItem` API — no extra work needed in handlers.
- Webviews built with `@fluentui/react` v8: keep the default focus rings; do not override `aria-*` attributes of Fluent components.
- Custom React controls in `packages/vscode-extension/src/controls/` follow [Fluent UI accessibility guidelines](https://react.fluentui.dev/?path=/docs/concepts-developer-accessibility-wai-aria--docs).
- Icons in tree views must come paired with text labels — never icon-only.

## CLI

- All inquirer prompts (`@inquirer/prompts`) are screen-reader-aware out of the box.
- Custom prompts in `packages/cli/src/userInteraction/customizedListPrompt.ts` and `customizedCheckboxPrompt.ts` preserve the same keyboard and aria contract — do not replace them with raw `readline`.
- Color is informational, never the sole signal: errors are bold red with the word "Error", success has a `✓` glyph plus green.
- Respect `NO_COLOR` env var (chalk does this automatically).

## Visual Studio

- Use VS-provided dialog primitives (e.g. `IVsThreadedWaitDialog`) — they are accessible by default.
- High-contrast themes are honoured automatically by VS XAML controls.

## Documentation surfaces

- All Mermaid diagrams in this docs site must be paired with a textual summary above or below them — Mermaid renders to SVG without semantic structure.
- Tables are preferred over multi-column layouts for technical reference.

## What we do not yet do

- Automated a11y testing in CI (Pa11y / axe). Tracked as future work; not on the current roadmap.
