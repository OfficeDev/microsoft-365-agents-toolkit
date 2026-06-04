# Route: Remote Provision & Deploy

**Preconditions** (must be satisfied before provision):
1. A scaffolded project is open in VS Code.
2. M365 auth is complete (see `routes/auth.md`).
3. Azure auth is complete (see Azure auth section below).

---

## Azure auth

```typescript
// Open ATK sidebar and click "Sign in to Azure"
await sendSignal("click:[id='workbench.view.extension.teamsfx']", 5000);
await wait(1000);
await sendSignal("clickText:Sign in to Azure", 10000);
await wait(2000);

// Screenshot — the Accounts pane should show the Azure subscription after sign-in
await takeElementScreenshot("provision-00-azure-signed-in",
  ".view-pane[aria-label*='Accounts'], [id='workbench.view.extension.teamsfx'] .pane-body");
```

---

## Remote Provision

```typescript
// 1. Run provision
//    Using the command directly (no QuickPick needed for most projects)
await vscode.commands.executeCommand("fx-extension.provision");
await wait(3000);
await takeElementScreenshot("provision-01-start", ".notifications-toasts");

// 2. Environment selection (if prompted)
await sendSignal("waitForTextThenScreenshot:Select an environment:30000:provision-02-env-select", 38000);
await takeElementScreenshot("provision-02-env-select", ".quick-input-widget");
await sendSignal("clickText:dev", 8000);
await wait(500);

// 3. Wait for completion
//    Typical provision takes 3–10 minutes. Use a long timeout.
await sendSignal(
  "waitForTextThenScreenshot:Provision completed:600000:provision-03-done",
  610000,
);
await takeElementScreenshot("provision-03-done", ".notifications-toasts");
```

---

## Remote Deploy

```typescript
// Deploy runs after a successful provision
await vscode.commands.executeCommand("fx-extension.deploy");
await wait(3000);
await takeElementScreenshot("deploy-01-start", ".notifications-toasts");

// Environment selection (if prompted)
await sendSignal("waitForTextThenScreenshot:Select an environment:30000:deploy-02-env-select", 38000);
await takeElementScreenshot("deploy-02-env-select", ".quick-input-widget");
await sendSignal("clickText:dev", 8000);
await wait(500);

// Wait for completion
await sendSignal(
  "waitForTextThenScreenshot:Deploy completed:600000:deploy-03-done",
  610000,
);
await takeElementScreenshot("deploy-03-done", ".notifications-toasts");
```

---

## Success condition

- `provision-03-done` screenshot shows the "Provision completed" notification.
- `deploy-03-done` screenshot shows the "Deploy completed" notification.

---

## Notes

- Both `fx-extension.provision` and `fx-extension.deploy` are registered command IDs in the extension. They do not have Command Palette titles in all builds, so use `executeCommand` rather than typing in the palette.
- Long-running operations (provision especially) may hit the 10-minute VS Code extension host timeout. Set `"extensions.experimental.affinity"` or increase the signal wait timeout if needed.
