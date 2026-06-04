# Route: Auth

Auth is required for Local Debug, Remote Provision, and Remote Deploy.
Scaffold does **not** require auth — skip this section for scaffold-only tests.

---

## Source patches (local dev build only)

Two patches are applied to the extension source to enable non-interactive auth:

### 1. `doesUserConfirmLogin` bypass — `m365Login.ts`

```typescript
// packages/vscode-extension/src/commonlib/m365Login.ts
async doesUserConfirmLogin(...args: unknown[]): Promise<boolean> {
  if (process.env.TEAMSFX_AUTO_CONFIRM_LOGIN === "true") return true;
  // ... original dialog code
}
```

**Effect**: the "Do you want to sign in?" modal is skipped.

### 2. Auth URL capture — `codeFlowLogin.ts`

```typescript
// packages/vscode-extension/src/commonlib/codeFlowLogin.ts
// After getAuthCodeUrl() resolves, write the URL to disk so the test can open it:
const authFile = path.join(os.tmpdir(), "atk-auth-url.txt");
fs.writeFileSync(authFile, authUrl, "utf8");
```

**Effect**: the test can read the auth URL and navigate a browser to it automatically.

After editing either file, rebuild the extension:
```bash
cd packages/vscode-extension
node esbuild.mjs --skip-type-check
```

---

## Environment variables

| Variable | Value | Effect |
|---|---|---|
| `TEAMSFX_AUTO_CONFIRM_LOGIN` | `true` | Bypasses the confirmation dialog (`doesUserConfirmLogin` patch above) |
| `TEAMSFX_BROKER_AUTH` | `false` | Forces MSAL browser code flow instead of WAM (Windows Authentication Manager) broker — required so a headless auth URL is generated |

Set both before VS Code launches:
```typescript
// In the Playwright launcher (e.g. scripts/src/verify.ts)
process.env.TEAMSFX_AUTO_CONFIRM_LOGIN = "true";
process.env.TEAMSFX_BROKER_AUTH = "false";
```

---

## Auth flow (step-by-step)

```typescript
// A. Open ATK sidebar, click "Sign in to Microsoft 365"
await sendSignal("click:[id='workbench.view.extension.teamsfx']", 5000);
await wait(1000);
await takeElementScreenshot("auth-01-atk-sidebar",
  "[id='workbench.view.extension.teamsfx'] .pane-body");
await sendSignal("clickText:Sign in to Microsoft 365", 10000);
await wait(2000);

// B. Read the auth URL written by the codeFlowLogin.ts patch
const authUrlFile = path.join(os.tmpdir(), "atk-auth-url.txt");
const deadline = Date.now() + 30000;
while (!fs.existsSync(authUrlFile) && Date.now() < deadline) await wait(500);
assert.ok(fs.existsSync(authUrlFile), "Auth URL file not written — check env vars and rebuild");
const authUrl = fs.readFileSync(authUrlFile, "utf8").trim();

// C. Open the auth URL in the system browser
await sendSignal(`openUrl:${authUrl}`, 5000);

// D. (Manual) complete the M365 sign-in in the browser.
//    When done, delete the auth URL file to signal completion:
while (fs.existsSync(authUrlFile)) await wait(1000);

// E. Confirm auth succeeded — screenshot the Accounts pane
//    The "Microsoft 365" account tile should show a user email.
await wait(3000);
await vscode.commands.executeCommand("workbench.view.extension.teamsfx");
await wait(1000);
await sendSignal(
  "waitForTextThenScreenshot:microsoft 365:30000:auth-02-signed-in",
  38000,
);
await takeElementScreenshot("auth-02-signed-in",
  ".view-pane[aria-label*='Accounts'], [id='workbench.view.extension.teamsfx'] .pane-body");
```

---

## Success condition

The screenshot `auth-02-signed-in` must show the signed-in email address under
the **Microsoft 365** section of the ATK Accounts pane.
