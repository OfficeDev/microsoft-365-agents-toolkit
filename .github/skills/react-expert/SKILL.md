---
name: react-expert
description: Use when building React applications in .jsx or .tsx files within the VS Code extension's webview controls (src/controls/). Creates components, implements custom hooks, debugs rendering issues, and optimizes performance. This project uses @fluentui/react v8 (not v9), react-router-dom MemoryRouter, react-intl for i18n, and Vite for the webview build.
license: MIT
metadata:
  author: https://github.com/Jeffallan
  version: "1.1.0"
  domain: frontend
  triggers: React, JSX, hooks, useState, useEffect, useContext, Suspense, component, frontend, webview, FluentUI, fluent
  role: specialist
  scope: implementation
  output-format: code
  related-skills: typescript-pro, test-master
---

# React Expert

Senior React specialist focused on the VS Code extension's React webview controls.

> **Project context:** React is used exclusively in `packages/vscode-extension/src/controls/`. The UI library is **`@fluentui/react` v8** (classic, not v9 `@fluentui/react-components`). Routing uses `react-router-dom` `MemoryRouter`. Internationalization uses `react-intl` `IntlProvider`. Build tool is **Vite** (separate from the main esbuild extension bundle).

## When to Use This Skill

- Building new webview panel components
- Implementing custom hooks for VS Code state
- Debugging rendering issues in webviews
- Optimizing performance of webview controls
- Migrating class components to functional hooks

## Core Workflow

1. **Analyze requirements** — Identify component hierarchy, state needs, data flow
2. **Choose patterns** — FluentUI v8 components, local state vs. context
3. **Implement** — Write TypeScript components with proper types
4. **Validate** — Run `tsc --noEmit`; fix all type issues before proceeding
5. **Optimize** — Apply memoization where needed, ensure accessibility
6. **Test** — Verify component renders correctly (visual inspection via `npm run dev` in vscode-extension)

## Key Patterns

### FluentUI v8 Component
```tsx
import { PrimaryButton, TextField, Stack } from "@fluentui/react";
import { useIntl } from "react-intl";

interface Props {
  onSubmit: (value: string) => void;
}

export const MyPanel: React.FC<Props> = ({ onSubmit }) => {
  const intl = useIntl();
  const [value, setValue] = React.useState("");

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <TextField
        label={intl.formatMessage({ id: "myPanel.label" })}
        value={value}
        onChange={(_, v) => setValue(v ?? "")}
      />
      <PrimaryButton
        text={intl.formatMessage({ id: "myPanel.submit" })}
        onClick={() => onSubmit(value)}
        disabled={!value}
      />
    </Stack>
  );
};
```

### Custom Hook with Cleanup
```tsx
function useWindowWidth(): number {
  const [width, setWidth] = React.useState(() => window.innerWidth);

  React.useEffect(() => {
    const handler = () => setWidth(window.innerWidth);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler); // cleanup
  }, []);

  return width;
}
```

## Constraints

### MUST DO
- Use TypeScript with strict mode
- Use `@fluentui/react` v8 components (not v9)
- Use `react-intl` for all user-facing strings — no raw string literals in JSX
- Clean up effects (return cleanup function)
- Use semantic HTML and ARIA for accessibility
- Memoize when passing callbacks/objects to memoized children

### MUST NOT DO
- Import from `@fluentui/react-components` (v9) — wrong package for this project
- Use `vscode.window.showErrorMessage()` directly from React — go through the extension message bridge
- Mutate state directly
- Use array index as key for dynamic lists
- Forget useEffect cleanup (memory leaks)

## Knowledge Reference

React 18, TypeScript, @fluentui/react v8, react-router-dom MemoryRouter, react-intl, Vite, custom hooks, useEffect, useState, useCallback, useMemo, React.memo
