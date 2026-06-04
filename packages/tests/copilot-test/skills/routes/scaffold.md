# Route: Scaffold

No auth required. Creates a new project through the ATK wizard and verifies the output files.

---

## Wizard option labels

Option labels come from `templates/src/ui/` — **read the relevant file before writing the click sequence**:

| File | Content |
|---|---|
| `templates/src/ui/root.ts` | First page: project-type options |
| `templates/src/ui/da.ts` | Declarative Agent sub-tree |
| `templates/src/ui/cea.ts` | Custom Engine Agent sub-tree |
| `templates/src/ui/teams.ts` | Teams Apps sub-tree |

---

## Pattern

```typescript
// 1. Random app name — prevents directory collisions between runs
const appName = `da-${Date.now()}`;

// 2. Fire wizard WITHOUT await — the wizard blocks the thread until the user finishes.
//    Omitting await lets sendSignal() run so Playwright can interact with the QuickPick.
vscode.commands.executeCommand("fx-extension.create").catch(() => {});
await wait(500); // yield event loop so the command dispatch reaches the extension host

// 3. Each wizard step:
//    a) Wait for the option text to appear (proves this step's QuickPick is open)
//    b) Screenshot the QuickPick widget before clicking
//    c) Click the option
await sendSignal("waitForTextThenScreenshot:Declarative Agent:60000:02-step1-project-type", 68000);
await takeElementScreenshot("02-step1-project-type", ".quick-input-widget");
await sendSignal("clickText:Declarative Agent", 10000);
await wait(800);

// ... repeat for each step using labels from templates/src/ui/*.ts ...

// 4. App name input step — type the random name
await sendSignal("waitForTextThenScreenshot:Application Name:15000:07-step-appname", 23000);
await takeElementScreenshot("07-step-appname", ".quick-input-widget");
await sendSignal(`type:${appName}`, 8000);
await wait(300);
await sendSignal("pressKey:Enter", 5000);

// 5. Scaffold in progress — ATK copies template files and opens a NEW VS Code window.
//    Screenshot the progress notification, then close the new window.
await wait(30000);
await takeElementScreenshot("08-scaffold-progress", ".notifications-toasts");
await wait(60000); // total ~90s for scaffold
await vscode.commands.executeCommand("workbench.action.closeWindow");

// 6. Open the scaffolded project in the current window
const projectDir = findScaffoldedDir(appName); // search ~/AgentsToolkitProjects then homedir
await vscode.commands.executeCommand("vscode.openFolder", vscode.Uri.file(projectDir));
await wait(5000);

// 7. Verify output files
//    Poll for sentinel file first (ATK writes m365agents.yml early in copy)
const sentinel = path.join(projectDir, "m365agents.yml");
for (let i = 0; i < 60 && !fs.existsSync(sentinel); i++) await wait(1000);

for (const relPath of EXPECTED_FILES) {
  const exists = fs.existsSync(path.join(projectDir, relPath));
  step(`File: ${relPath}`, exists, exists ? "✓" : `not found in ${projectDir}`);
  // Screenshot explorer to show actual directory state for each check
  await takeElementScreenshot(`verify-${relPath.replace(/\//g, "-")}`, ".explorer-viewlet");
}
```

---

## findScaffoldedDir helper

```typescript
function findScaffoldedDir(appName: string): string {
  const roots = [
    path.join(os.homedir(), "AgentsToolkitProjects"),
    os.homedir(),
    os.tmpdir(),
    process.cwd(),
  ];
  for (const root of roots) {
    if (!fs.existsSync(root)) continue;
    const direct = path.join(root, appName);
    if (fs.existsSync(direct)) return direct;
    try {
      for (const entry of fs.readdirSync(root)) {
        if (entry.includes(appName)) {
          const full = path.join(root, entry);
          try { if (fs.statSync(full).isDirectory()) return full; } catch {}
        }
      }
    } catch {}
  }
  return "";
}
```

---

## Always-expected files (every template)

- `m365agents.yml`
- `appPackage/manifest.json`

Additional files depend on the specific template — list them explicitly in the test plan.
