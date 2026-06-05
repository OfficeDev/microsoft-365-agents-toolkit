# Route: Scaffold

No auth required. Creates a new project through the ATK wizard and verifies the output files.

---

## Two-window design (important)

When ATK finishes scaffolding, it **automatically opens a second VS Code window** with the new project as the workspace root. The first window remains open but has no project.

**The correct test pattern is:**
- **Window 1** (wizard window): run the wizard steps → write a `scaffold-done.json` marker with the `projectDir` → let the window exit naturally
- **Window 2** (project window): read the marker → verify files + do sign-in + local debug → clean up marker

Do **not** try to:
- Reopen the project in Window 1 via `vscode.openFolder` (that reloads the window and breaks the extension host)
- Use `updateWorkspaceFolders` to add the project folder (async, unreliable timing)
- Close Window 2 and continue in Window 1

Window 2 is the correct surface for everything after scaffold because:
- The project is already the workspace root — `launch.json` configs are immediately visible
- ATK is already activated (triggered by `workspaceContains:m365agents*.yml`)
- No folder manipulation needed

### Marker file pattern

```typescript
// In Window 1 — after pressing Enter on app name:
const projectDir = path.join(os.homedir(), "AgentsToolkitProjects", appName);
// Poll briefly to confirm scaffold started
for (let i = 0; i < 30; i++) {
  if (fs.existsSync(projectDir)) break;
  await wait(1000);
}
// Write marker for Window 2
fs.writeFileSync(
  path.join(OUTPUT_DIR, "scaffold-done.json"),
  JSON.stringify({ appName, projectDir, createdAt: Date.now() }),
  "utf8",
);
// Window 1's work is done — it will exit when ATK finishes and the test suite completes
```

```typescript
// In test file top-level (both windows read this at load time):
const SCAFFOLD_MARKER = path.join(OUTPUT_DIR, "scaffold-done.json");
const isWindow2 = fs.existsSync(SCAFFOLD_MARKER);
let scaffoldedProjectDir = "";
if (isWindow2) {
  try {
    const marker = JSON.parse(fs.readFileSync(SCAFFOLD_MARKER, "utf8"));
    scaffoldedProjectDir = marker.projectDir;
  } catch {}
}

// Then gate test cases:
if (!isWindow2) {
  test("Phase 1: wizard", async () => { /* ... wizard steps ... */ });
} else {
  test("Phase 2: verify files", async () => { /* ... */ });
  test("Sign in + debug", async () => { /* ... */ });
}
```

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

## Window 1 pattern (wizard only)

```typescript
const appName = `da-${Date.now()}`;

// Fire wizard WITHOUT await
vscode.commands.executeCommand("fx-extension.create").then(undefined, () => {});
await wait(500);

// Each wizard step: wait for text → screenshot → click/select
await sendSignal("waitForTextThenScreenshot:Declarative Agent:60000:02-step1-project-type", 68000);
await sendSignal("pressKey:Enter", 5000); // first item is pre-highlighted; Enter selects it
await wait(800);

// ... repeat for each step using labels from templates/src/ui/*.ts ...

// App name step
await sendSignal("waitForTextThenScreenshot:Application Name:15000:05-wizard-appname", 23000);
await sendSignal(`type:${appName}`, 8000);
await wait(300);
await sendSignal("pressKey:Enter", 5000);

// ATK opens a new VS Code window with the project — Window 1's job is done.
// Poll briefly to confirm the project dir was created, then write the marker.
const projectDir = path.join(os.homedir(), "AgentsToolkitProjects", appName);
for (let i = 0; i < 30; i++) {
  if (fs.existsSync(projectDir)) break;
  await wait(1000);
}
fs.writeFileSync(
  path.join(OUTPUT_DIR, "scaffold-done.json"),
  JSON.stringify({ appName, projectDir, createdAt: Date.now() }),
  "utf8",
);
await takeScreenshot("06-scaffold-started");
```

---

## Window 2 pattern (verify files)

```typescript
// projectDir comes from scaffold-done.json read at top of file
const sentinel = path.join(projectDir, "m365agents.yml");
for (let i = 0; i < 90 && !fs.existsSync(sentinel); i++) await wait(1000);
await takeScreenshot("06-scaffold-explorer");

for (const relPath of EXPECTED_FILES) {
  const exists = fs.existsSync(path.join(projectDir, relPath));
  step(`${relPath} exists`, exists, exists ? "✓" : `missing in ${projectDir}`);
}
// Clean up marker when all work in Window 2 is complete
try { fs.unlinkSync(path.join(OUTPUT_DIR, "scaffold-done.json")); } catch {}
```

---

## Always-expected files (every template)

- `m365agents.yml`
- `appPackage/manifest.json`

Additional files depend on the specific template — list them explicitly in the test plan.
