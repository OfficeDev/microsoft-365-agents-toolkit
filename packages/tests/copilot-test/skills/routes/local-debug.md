# Route: Local Debug

**Preconditions** (must be satisfied before any debug step):
1. A scaffolded project is already open in VS Code (run scaffold first, or pre-scaffold and point VS Code at the folder).
2. Auth has been completed (see `routes/auth.md`).

---

## Pattern

```typescript
// 1. Confirm the workspace file exists
const yml = path.join(workspaceDir, "m365agents.yml");
assert.ok(fs.existsSync(yml), `Workspace must contain m365agents.yml: ${yml}`);
await takeElementScreenshot("debug-00-explorer", ".explorer-viewlet");

// 2. Open the Run & Debug panel
await vscode.commands.executeCommand("workbench.view.debug");
await wait(1000);
await takeElementScreenshot("debug-01-debug-panel",
  "[id='workbench.view.debug'] .pane-body");

// 3. Select the correct debug configuration
//    Configuration names come from .vscode/launch.json in the scaffolded project.
//    Screenshot the dropdown before selecting so reviewers can confirm the right option.
await takeElementScreenshot("debug-02-config-picker",
  ".debug-configuration-select-container, .debug-toolbar, .quick-input-widget");
await sendSignal("clickText:Debug (Edge)", 10000);  // replace with actual config name
await wait(500);

// 4. Start debugging (F5 equivalent)
await vscode.commands.executeCommand("workbench.action.debug.start");
await wait(5000);
await takeElementScreenshot("debug-03-start",
  ".debug-toolbar, .notifications-toasts");

// 5. Wait for the debug session to attach
//    "Running on port" or similar text signals readiness — read from the scenario doc.
await sendSignal("waitForTextThenScreenshot:Listening on:120000:debug-04-running", 130000);
await takeElementScreenshot("debug-04-running", ".debug-toolbar");

// 6. Stop debug
await vscode.commands.executeCommand("workbench.action.debug.stop");
await wait(3000);
await takeElementScreenshot("debug-05-stopped", ".debug-toolbar");
```

---

## Success condition

Screenshot `debug-04-running` must show the debug toolbar in its active state and
(where applicable) the "Listening on port …" text in the terminal or notification.

---

## Notes

- The debug configuration name in step 3 changes per template. Check the generated
  `.vscode/launch.json` and use the exact label text.
- For bot/agent templates, the debug session typically starts Microsoft 365 Agents Playground.
  Screenshot the Playground URL or app-up notification as the success evidence.
