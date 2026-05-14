# Security

The toolkit handles user credentials, downloads remote ZIP templates, executes user code, and ships software to wide audiences. The security baseline is non-negotiable.

## Filesystem — EAFP

Never check existence before operating. CodeQL flags TOCTOU race conditions (`js/toctou-race-condition`).

```ts
// BAD
if (await fs.access(path).then(() => true).catch(() => false)) {
  await fs.readFile(path);
}

// GOOD
try {
  const data = await fs.readFile(path);
} catch (e: any) {
  if (e.code === "ENOENT") { /* handle */ }
  else throw e;
}
```

## ZIP extraction — Zip Slip

Validate **every** entry name before writing:

1. Reject names containing `..` segments (`name.indexOf("..") !== -1`).
2. After joining with target dir, `path.resolve()` must remain a prefix of the target dir.

Source: `packages/core-next/src/templates/scaffold/download.ts` (and v3 equivalents in `fx-core`).

## ZIP upload — magic-byte validation

Before uploading any ZIP to a service (TDP, Kudu, etc.), verify the first 4 bytes match `PK\x03\x04`:

```ts
const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
if (!buffer.subarray(0, 4).equals(ZIP_MAGIC)) {
  return err(new UserError({ source, name: "InvalidZipPackage", message: ... }));
}
```

Source: `packages/core-next/src/clients/teamsDevPortal/client.ts`, `packages/core-next/src/drivers/builtin/azureAppService/zipDeploy.ts`.

## Secret masking

Every log line, error message, and telemetry property passes through the secret masker before transport.

| Engine | Mechanism |
|--------|-----------|
| v3 | `common/secretmasker/` — keyword + SVM (string variable matcher) + BloomFilter |
| v4 | `core-next/src/secretMasker/` — keyword regex only (100+ credential suffixes); pure functions |

URLs are sanitised in HTTP client interceptors — query-string credentials are stripped.

## Token cache

`~/.fx/account/` holds MSAL caches encrypted with **AES-256-GCM**. Symmetric key resolution:

1. If `keytar` is installed and an OS keychain is reachable → key in keychain.
2. Otherwise → derived from a machine-stable secret (less strong, still encrypted at rest).

Shared between v3 and v4. Logout (`atk account logout m365|azure`) deletes the cache files.

## CI auth secrets

Never accept secrets via CLI args (they leak in shell history). Only `process.env`. The CLI's auto-detect of CI mode (env vars: `CI`, `TF_BUILD`, `GITHUB_ACTIONS`) routes through `azureLoginCI.ts` which expects `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` (or OIDC for federated workload identity — preferred).

## Telemetry

- `error.message` (English) is sent; `error.displayMessage` (may contain user data) is **not**.
- Property values pass through the masker.
- No `error.stack` in production builds — only `error.name` + `error.message`.
- Correlation IDs are random UUIDs — not derived from any user-identifying source.

## Code execution

Drivers run user-supplied scripts (`script`, `cli/runNpmCommand`, `cli/runDotnetCommand`). The toolkit:

- Does **not** sandbox these — they are user code on the user's machine.
- **Does** scrub `process.env` after each driver to avoid leaking earlier-step secrets to later steps that don't need them.
- **Does not** auto-elevate privileges — runs at the toolkit's own privilege level.

## Dependency hygiene

- Renovate auto-PRs upstream updates.
- `pnpm audit` run in CI; high-severity vulnerabilities block merge.
- Native node addons (`keytar`, `msal-node-extensions`) are optional deps with graceful fallback paths.
- Pinned `swagger2openapi 7.0.8` — newer versions break `RemoteMCPServer`; tracked as risk.

## Reporting

Security issues go to MSRC: <secure@microsoft.com>. See [`SECURITY.md`](../../SECURITY.md). Public disclosure follows MSRC disclosure timelines.
