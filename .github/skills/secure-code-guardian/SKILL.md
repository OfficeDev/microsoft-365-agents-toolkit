---
name: secure-code-guardian
description: Use when implementing authentication/authorization, securing user input, or preventing OWASP Top 10 vulnerabilities. Invoke for MSAL token handling, Azure credential management, secret masking, input validation with Zod, JWT validation, rate limiting, and security hardening of CLI and extension code. This project uses MSAL (msal-node, msal-node-extensions) for auth and keytar for secret storage.
license: MIT
metadata:
  author: https://github.com/Jeffallan
  version: "1.1.0"
  domain: security
  triggers: security, authentication, authorization, encryption, OWASP, vulnerability, secure coding, MSAL, JWT, OAuth, keytar, secret, token, credential, maskSecret
  role: specialist
  scope: implementation
  output-format: code
  related-skills: security-reviewer, architecture-designer
---

# Secure Code Guardian

## Core Workflow

1. **Threat model** — Identify attack surface and threats
2. **Design** — Plan security controls
3. **Implement** — Write secure code with defense in depth
4. **Validate** — Test security controls
5. **Document** — Record security decisions

### Validation Checkpoints

After each implementation step, verify:

- **Authentication**: Test token expiration, invalid-credential error messages (must not leak user existence), MSAL token refresh paths
- **Authorization**: Verify privilege escalation paths are blocked; test with tokens belonging to different roles/scopes
- **Input handling**: Confirm injection payloads are rejected or escaped
- **Secret handling**: Confirm `maskSecret()` is called before any credential appears in logs or telemetry

## Constraints

### MUST DO
- Use `maskSecret()` from `packages/core-next/src/secretMasker/` before logging any credential, token, or key
- Store secrets via `keytar` — never in plaintext files or environment variables committed to repo
- Validate all external input at system boundaries (CLI args, API responses, file contents)
- Use `validator` library for URL/string validation — no hand-rolled regex for security checks
- Use parameterized queries — never string-interpolated SQL or command injection paths
- Rate-limit authentication endpoints
- Set security headers for any HTTP surface

### MUST NOT DO
- Log tokens, credentials, or secrets — always `maskSecret()` first
- Store passwords or tokens in plaintext
- Trust user input without validation
- Expose sensitive data in error responses or telemetry
- Use weak or deprecated algorithms (MD5, SHA-1, DES, ECB mode)
- Hardcode credentials or client secrets in source code

## Code Examples

### Secret Masking (project-specific)
```typescript
import { maskSecret } from "../secretMasker";

// BAD — logs raw token
logger.info(`Token: ${accessToken}`);

// GOOD — masks before logging
logger.info(`Token: ${maskSecret(accessToken)}`);
```

### Input Validation with Zod
```typescript
import { z } from "zod";

const ProvisionInputSchema = z.object({
  subscriptionId: z.string().uuid(),
  resourceGroupName: z.string().min(1).max(90).regex(/^[a-zA-Z0-9._-]+$/),
  location: z.string().min(1),
});

export function validateProvisionInput(raw: unknown) {
  const result = ProvisionInputSchema.safeParse(raw);
  if (!result.success) {
    return err(new UserError({
      source: "ProvisionOp",
      name: "InvalidProvisionInput",
      message: getLocalizedString("error.invalidInput"),
    }));
  }
  return ok(result.data);
}
```

### JWT Validation
```typescript
import jwt from "jsonwebtoken";

export function verifyToken(token: string): jwt.JwtPayload {
  const payload = jwt.verify(token, process.env.JWT_SECRET!, {
    algorithms: ["HS256"],
    issuer: "your-app",
    audience: "your-app",
  });
  if (typeof payload === "string") throw new Error("Invalid token payload");
  return payload;
}
```

## Output Templates

When implementing security features, provide:
1. Secure implementation code
2. Security considerations noted
3. Configuration requirements (env vars, headers)
4. Testing recommendations

## Knowledge Reference

OWASP Top 10, MSAL (msal-node, msal-node-extensions), keytar, maskSecret, JWT, OAuth 2.0, OIDC, Zod validation, rate limiting, input validation, output encoding, TLS, security headers
