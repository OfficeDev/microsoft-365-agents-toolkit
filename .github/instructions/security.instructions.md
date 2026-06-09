---
description: Security rules for toolkit TypeScript code — webview XSS, shell/path injection, untrusted JSON parsing, cryptography. Complements the cross-cutting rules in copilot-instructions.md (maskSecret, EAFP, no `as` casts).
applyTo: 'packages/**/*.ts,packages/**/*.tsx'
---

# Security Rules

Scope: the toolkit's TypeScript engine, CLI, and VS Code extension. The toolkit
does not host a public HTTP API, does not issue JWTs, does not talk to SQL
databases, and delegates user authentication and authorization to Microsoft
Identity (MSAL / `@azure/identity`). Categories below are limited to threats
that actually apply to that surface.

Cross-cutting basics (`maskSecret()` before logging, EAFP for filesystem, no
`as` casts, `getLocalizedString` for user-facing strings) live in
[copilot-instructions.md](../copilot-instructions.md) — those still apply; this
file adds the concrete patterns reviewers and the AI need to see.

## Webview XSS — `packages/vscode-extension/src/controls/**`

React escapes children by default. The only XSS vectors are `innerHTML` and
`dangerouslySetInnerHTML`. Anything user- or project-supplied that reaches
those must be sanitized or rejected.

```tsx
// Bad
element.innerHTML = userInput;
return <div dangerouslySetInnerHTML={{ __html: userInput }} />;

// Good
element.textContent = userInput;
return <div>{userInput}</div>;
```

Verify the webview is created with a Content Security Policy that disables
inline scripts and external sources.

## Command injection — `child_process` callers

The CLI and extension shell out to `func`, `azd`, `npm`, etc. Never build a
shell string with template literals.

```typescript
// Bad
import { exec } from "child_process";
exec(`func start --port ${port}`);

// Good
import { spawn } from "child_process";
spawn("func", ["start", "--port", String(port)], { shell: false });
```

If a shell is unavoidable (piping, env expansion), quote arguments explicitly
and never pass user-controlled values raw.

## Path traversal — template / project file I/O

Resolve user-supplied paths against a fixed root and verify containment before
read or write.

```typescript
// Bad
const filePath = `./templates/${userInput}`;

// Good
import * as path from "path";
const root = path.resolve("./templates");
const filePath = path.resolve(root, path.normalize(userInput));
if (!filePath.startsWith(root + path.sep)) {
  return err(new UserError({
    source: "fx-core",
    name: "PathEscapesRoot",
    message: getLocalizedString("error.path.escapes-root"),
  }));
}
```

## Untrusted JSON — manifests, plugin specs

Teams / Declarative Agent / API Plugin manifests come from the user's project.
Parse via the validating converters in `packages/manifest`; do not `as`-cast
`JSON.parse` results.

```typescript
// Bad
const manifest = JSON.parse(text) as TeamsManifest;

// Good
import { AppManifestUtils } from "@microsoft/app-manifest";
const result = AppManifestUtils.parse(text);
if (result.isErr()) return err(result.error);
const manifest = result.value;
```

## Cryptography

The engine and CLI should derive as little as possible — push secret material
into Key Vault / VS Code SecretStorage and reference it. When derivation is
unavoidable:

```typescript
// Bad
const token = Math.random().toString(36);
import { createHash } from "crypto";
const hash = createHash("md5").update(secret).digest("hex");

// Good
import { randomBytes, scrypt } from "crypto";
const token = randomBytes(32).toString("hex");
// For password-derived material, use scrypt / a vetted KDF.
```

Never use `Math.random` for tokens, IDs that need unpredictability, or
anything security-sensitive. Never use MD5 or SHA1 for credentials.

## Error messages to the user

Do not leak internal paths, stack frames, or remote URLs into the user-facing
string. Keep full context in telemetry; surface a localized friendly message.

```typescript
// Bad
return err(new SystemError({
  source: "fx-core",
  name: "ProvisionFailed",
  message: `${error.stack} for ${innerUrl}`,
}));

// Good
logger.error("Provision failed", { error, correlationId });
return err(new SystemError({
  source: "fx-core",
  name: "ProvisionFailed",
  message: getLocalizedString("error.provision.generic"),
}));
```

## Dependencies

- Use the official npm registry only.
- `pnpm audit` clean at merge time (no high / critical).
- `pnpm-lock.yaml` is committed; never hand-edit it — run `pnpm install`.
- Prefer Node stdlib or an existing helper over pulling a new dependency.
