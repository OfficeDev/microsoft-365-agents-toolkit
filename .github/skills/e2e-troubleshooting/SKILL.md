---
name: e2e-troubleshooting
description: "Diagnose and fix E2E lifecycle test failures in core-next and cli-next. Use when: CI E2E tests fail, lifecycle test errors, provision/deploy failures, driver registration issues, template scaffolding errors, ARM deployment failures, permission errors, debugging GitHub Actions run logs."
argument-hint: "Paste the CI run URL or describe the test failure"
---

# E2E Lifecycle Test Troubleshooting

## When to Use

- Investigating failed E2E tests from `.github/workflows/e2e-test-next.yml`
- Diagnosing `lifecycle.test.ts` failures (scaffold / provision / deploy / validate)
- Fixing driver registration, template rendering, or ARM deployment issues
- Analyzing GitHub Actions run logs for `core-next` / `cli-next`

## Architecture Overview

For the full test flow diagram, key files, CI matrix structure, and template language matrix, see **[`docs/06-operations/e2e-strategy.md`](../../docs/06-operations/e2e-strategy.md)**.

Quick reference — key files for diagnosis:

| File | Purpose |
|------|---------|
| `packages/cli-next/tests/e2e/lifecycle.test.ts` | Data-driven E2E test generator |
| `packages/core-next/src/drivers/builtin/index.ts` | Driver registry (all builtin drivers) |
| `packages/core-next/src/templates/descriptors/` | Template descriptors (id, languages, tags, testable) |
| `.github/workflows/e2e-test-next.yml` | CI workflow (matrix of test files) |

## Failure Diagnosis Workflow

### Step 1: Get the Logs

```bash
# List failed jobs
gh run view <run-id> --repo OfficeDev/microsoft-365-agents-toolkit

# Get failed job logs (can be large)
gh run view <run-id> --log-failed --repo OfficeDev/microsoft-365-agents-toolkit 2>&1 \
  | Select-String "<template-name>" | Select-String "Error|fail|assert" | Select-Object -First 30
```

> **Tip:** `--log-failed` output is huge. Always pipe through `Select-String` (PowerShell)
> or `grep` (bash) to filter for the failing template name first, then for error patterns.

### Step 2: Categorize the Failure

| Error Pattern | Category | Root Cause | Fix Location |
|---------------|----------|------------|--------------|
| `Driver "<name>" not registered` | Missing driver | Driver not implemented or not in registry | `drivers/builtin/index.ts` + new driver file |
| `Invalid configuration` / zod validation | Driver input | Schema mismatch or missing optional field | Driver's `inputSchema` (zod) |
| `ENOENT` on file path | Path resolution | Relative path not resolved against projectPath | Driver's `execute()` function |
| `Cannot convert 'X' to Edm.Guid` | Permission resolution | AAD manifest has friendly names, not GUIDs | `aadApp/update.ts` permission mapping |
| `DeploymentFailed` (ARM) | ARM deployment | Password complexity, SKU, naming, quota | Bicep template or dummy env values |
| `npm ERR!` / `dotnet build` failure | Build failure | Mustache rendering issue, missing env vars | Template `.tpl` files or scaffold options |
| `authorizationUrl is required` | OAuth config | Template scaffolds wrong OAuth variant | `TEMPLATE_TEST_OPTIONS` in lifecycle.test.ts |
| `maxLength` / `minLength` ARM error | Name too long | `resourceBaseName` exceeds Bicep constraint | `getUniqueAppName()` or `RESOURCE_SUFFIX` |
| `400 /appCatalogs/teamsApps/…/appDefinitions?requiresReview=true` | Publish update (sideloaded) | Shared-scope sideloaded app doesn't support admin-review update | `publishTeamsAppUpdate` fallback in `graphApi/client.ts` |
| `404 /appCatalogs/teamsApps/…/appDefinitions` | Publish update (phantom) | Sideloaded app has catalog entry but no real REST resource | `publishTeamsAppUpdate` 400→404 fallback returns existing ID |
| `412 /appCatalogs/teamsApps?requiresReview=true` | Publish first-time (transient) | Graph API etag/propagation race; retryable | `sendWithRetry` exempts 412 from no-retry rule |
| `404 /appCatalogs/teamsApps/<id>` during cleanup | Unpublish (expected) | App already removed by tenant policies or transient | Cleanup uses `Promise.allSettled`; safe to ignore |

### Step 3: Fix by Category

See **Common Fix Patterns** below.

## Common Fix Patterns

### 1. Missing Driver

**Symptom:** `Error: Driver "cli/runDotnetCommand" is not registered`

**Diagnosis:** A template's `m365agents.yml` references a driver that doesn't exist in `core-next`.
Check which zip (ts, js, csharp, python) contains the template and inspect its YAML.

```bash
# List entries in a zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead("packages/core-next/templates/fallback/csharp.zip")
$zip.Entries | Where-Object { $_.FullName -like "*<template-name>*" } | Select FullName
$zip.Dispose()
```

**Fix:**
1. Create the driver in `packages/core-next/src/drivers/builtin/<category>/<name>.ts`
2. Follow the existing driver pattern: `createDriver()` with zod `inputSchema` and `execute` fn
3. Register it in `packages/core-next/src/drivers/builtin/index.ts`
4. Update all test files that assert driver count (search for the old count number)

**Driver count test files** (grep for the current count, e.g. `22`):
- `tests/unit/drivers/builtin/teamsAppPlatform.test.ts`
- `tests/unit/drivers/builtin/teamsAppDrivers.test.ts` (2 assertions)
- `tests/unit/drivers/builtin/registration.test.ts` (2 assertions)
- `tests/unit/drivers/builtin/entraAndBotDrivers.test.ts`
- `tests/unit/drivers/builtin/azureInfraDrivers.test.ts`
- `tests/unit/drivers/builtin/authPluginDrivers.test.ts`
- `tests/integration/lifecycleExecution.test.ts`

### 2. Driver Schema Mismatch (Optional Fields)

**Symptom:** `Invalid configuration` error with zod validation details

**Diagnosis:** The template's YAML passes fewer fields than the driver schema requires.
Compare the YAML `with:` block against the driver's `inputSchema`.

**Common case:** A field like `baseUrl` is required in the schema but the template only
provides `apiSpecPath`. The fx-core (v3) driver may have handled this, but core-next
doesn't yet.

**Fix:** Make the field optional in the zod schema and add fallback logic in `execute()`.
For example, extract domains from `apiSpecPath` (OpenAPI spec) when `baseUrl` is missing:

```typescript
// Parse OpenAPI spec to extract server URLs
const specContent = fs.readFileSync(resolvedSpecPath, "utf-8");
const spec = specPath.endsWith(".yaml") ? yaml.load(specContent) : JSON.parse(specContent);
const servers = spec?.servers ?? [];
const domains = servers.map((s: { url: string }) => new URL(s.url).hostname).filter(Boolean);
```

### 3. AAD Permission Name → GUID Resolution

**Symptom:** `Cannot convert 'ExternalConnection.ReadWrite.OwnedBy' to Edm.Guid`

**Diagnosis:** The AAD app manifest in the template uses human-readable permission
names (e.g. `ExternalConnection.ReadWrite.OwnedBy`) instead of GUIDs. The Graph API
`PATCH /applications/{id}` endpoint requires GUIDs in `requiredResourceAccess`.

**Fix:** Add permission resolution in `aadApp/update.ts` using `permissions.json`
(copied from `packages/fx-core/src/component/driver/aad/permissions/permissions.json`).
The mapping structure:

```json
{ "value": [{ "appId": "...", "appRoles": [{"id": "guid", "value": "Name"}], "oauth2PermissionScopes": [...] }] }
```

Resolution logic: for each `resourceAccess` entry, if `id` is not a valid GUID,
search permissions.json for a matching `value` field and replace with the real GUID.

### 4. ARM Deployment Failures

**Symptom:** `DeploymentFailed` — "At least one resource deployment operation failed"

**Common sub-causes:**

| Sub-cause | How to Identify | Fix |
|-----------|----------------|-----|
| **Password complexity** | Template has SQL Server resource with `@secure() param sqlAdminPassword` | In `collectUnresolvedParameterVars()`, generate complex passwords for vars matching `/password/i` |
| **Name too long** | Bicep has `@maxLength(20) param resourceBaseName` | Keep `getUniqueAppName()` ≤ 17 chars; pre-set `RESOURCE_SUFFIX` to 6 chars before collecting unresolved vars |
| **Missing env vars** | `${{VAR}}` placeholders in `parameters.json.tpl` not populated | `collectUnresolvedParameterVars()` fills these with dummy values |
| **SKU unavailable** | Region doesn't support the requested SKU | Change region in `createResourceGroup()` or template |

**Key function:** `collectUnresolvedParameterVars()` in lifecycle.test.ts scans
`infra/*.parameters.json` for `${{VAR}}` placeholders not yet defined in env,
then populates them with dummy values:

```typescript
// Password vars get complex values for Azure SQL etc.
if (/password/i.test(varName)) {
  vars.set(varName, "Tst!Pa55w0rd#E2E");
} else {
  vars.set(varName, `test-placeholder-${varName.toLowerCase()}`);
}
```

### 5. Template Scaffold Option Mismatch

**Symptom:** Template scaffolds a code path requiring runtime secrets/endpoints
not available in CI (e.g. Custom OAuth `authorizationUrl`).

**Fix:** Add the template to `TEMPLATE_TEST_OPTIONS` in lifecycle.test.ts to override
scaffold options and select a CI-friendly variant:

```typescript
const TEMPLATE_TEST_OPTIONS: Record<string, Record<string, string>> = {
  "da/api-plugin-oauth": { authType: "microsoft-entra" },
  // Add new overrides here as needed
};
```

These options are spread into the `createProjectOp` call:
```typescript
options: { llmService: "azure-openai", ...TEMPLATE_TEST_OPTIONS[template.id] },
```

### 6. Mustache Rendering Issues

**Symptom:** Template renders with literal `{{#condition}}...{{/condition}}` blocks
instead of being evaluated. Or `${{VAR}}` env placeholders get eaten by Mustache.

**Diagnosis:** Mustache uses `{{...}}` syntax which conflicts with the toolkit's
`${{VAR}}` env placeholder syntax. If the template has both Mustache conditionals
and env placeholders, collisions occur.

**Fix:** The render engine (`render.ts`) uses sentinel substitution:
1. Before Mustache: replace `${{VAR}}` with `\0VAR\0` (NUL-byte sentinels)
2. Run Mustache rendering
3. After Mustache: restore `\0VAR\0` back to `${{VAR}}`

### 7. Lifecycle Phase Detection

**Symptom:** Test creates an Azure resource group when none is needed (wasting CI time),
or skips provision/deploy that should run.

**Key logic in lifecycle.test.ts:**

```typescript
// Phase detection from m365agents.yml
const hasProvisionLifecycle = yamlExists && yamlContent.includes("provision:");
const hasDeployLifecycle = yamlExists && yamlContent.includes("deploy:");

// Azure need detection — checks for specific driver prefixes
const AZURE_DRIVER_PREFIXES = ["arm/", "azureFunctions/", "azureAppService/", "azureStorage/"];
const needsAzure = yamlExists && yamlNeedsAzure(yamlContent);
```

Templates with only `teamsApp/*` or `cli/*` drivers skip RG creation entirely.

## Template Language Matrix

See [`docs/06-operations/e2e-strategy.md` §Template Language Matrix](../../docs/06-operations/e2e-strategy.md#template-language-matrix).

## Debugging Checklist

When a new E2E failure appears:

1. **Get logs** — `gh run view <id> --log-failed` filtered for the failing template
2. **Categorize** — Match the error pattern to the table in Step 2 above
3. **Check template YAML** — Extract from the appropriate zip, read `m365agents.yml.tpl`
4. **Check driver exists** — Is every `uses:` driver registered in `builtin/index.ts`?
5. **Check driver schema** — Does the YAML `with:` block satisfy the driver's zod schema?
6. **Check env vars** — Are all `${{VAR}}` placeholders in parameter files resolved?
7. **Check names** — Does `resourceBaseName` fit Bicep `@maxLength` constraints?
8. **Check credentials** — Are AAD manifests using GUIDs (not friendly permission names)?
9. **Build & unit test** — `pnpm --filter ./packages/core-next build && npm run test:unit`
10. **Update driver count** — If you added/removed a driver, grep for the old count in test files

## Making a Template Non-Testable

If a template can't be fixed for CI (requires infrastructure not available in E2E),
mark it as `testable: false` in its descriptor:

```typescript
// packages/core-next/src/templates/descriptors/<category>.ts
{ id: "da/mcp-remote", ..., testable: false }
```

This excludes it from `templateRegistry.list().filter(t => t.testable !== false)`.
Use sparingly — prefer fixing the root cause.

## Graph API Publish Failures

The `teamsApp/publishAppPackage` step uses the Graph beta endpoint to publish or update
apps in the tenant catalog. Several failure modes arise from how the Graph API handles
sideloaded vs admin-published apps.

### Publish Flow

```
publishAppPackage driver
  │
  ├─ First time? → publishTeamsApp()
  │    POST /appCatalogs/teamsApps?requiresReview=true
  │    ├─ 200 with id → OK
  │    ├─ 200 with error body (BadGateway) → fallback to getStagedApp
  │    ├─ 409 (AppDefinitionAlreadyExists) → publishTeamsAppUpdate
  │    └─ 412 (PreconditionFailed) → retried by sendWithRetry
  │
  └─ Already published? → publishTeamsAppUpdate()
       GET /appCatalogs/teamsApps?$filter=externalId eq '...'
       POST /appCatalogs/teamsApps/{id}/appDefinitions?requiresReview=true
       ├─ 200 → OK
       ├─ 400 → retry without ?requiresReview=true
       │    ├─ 200 → OK
       │    └─ 404 → return existing catalog ID (phantom entry)
       └─ other error → fail
```

### Scope-Dependent Behavior

| Sideload Scope | Catalog Entry | requiresReview Update | Without requiresReview |
|---------------|---------------|----------------------|----------------------|
| **Personal** | No | N/A (first-time publish) | N/A |
| **Shared** | Yes (phantom) | 400 BadRequest | 404 Not Found |
| **Admin-published** | Yes (real) | 200 OK | 200 OK |

**Personal scope:** `extendToM365` with `scope: Personal` does NOT create a catalog
entry. The publish step does a first-time `POST /appCatalogs/teamsApps`.

**Shared scope:** `extendToM365` with `scope: Shared` creates a catalog entry visible
to `GET /appCatalogs/teamsApps?$filter=externalId eq '...'`, but this entry doesn't
support the REST update endpoints. Both `?requiresReview=true` (400) and without (404)
fail. The fix returns the existing catalog ID since the app IS already published via
sideloading.

### sendWithRetry Exemptions

The retry helper does NOT retry most 4xx errors (client bugs won't self-heal), but
exempts two status codes that are known to be transient:

| Status | Meaning | Why Retryable |
|--------|---------|---------------|
| **429** | Too Many Requests | Rate limiting; will succeed after backoff |
| **412** | Precondition Failed | Graph API etag/propagation race; resolves on retry |

All other 4xx (400, 401, 403, 404, 409, etc.) throw immediately without retry.

**Key file:** `packages/core-next/src/http/retry.ts`

### Cleanup Failures (Safe to Ignore)

After a successful test run, cleanup attempts to:
1. Delete the Azure resource group (`az group delete`)
2. Unpublish the Teams app (`DELETE /appCatalogs/teamsApps/{id}`)

Both can return 404 if the resource was already removed. These are wrapped in
`Promise.allSettled` and logged but don't fail the test. Common patterns:

- `Resource group '...' could not be found` — DA templates have no Azure resources
- `App doesn't exist in the tenant` — tenant policies auto-removed the app
