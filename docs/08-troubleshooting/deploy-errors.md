# Deploy errors

Errors surfaced during the `deploy` lifecycle.

## `NpmBuildFailed` / `DotnetBuildFailed`

**Trigger.** `cli/runNpmCommand` or `cli/runDotnetCommand` returned non-zero exit code.

**Mitigation.** Application-level. Read the captured stdout/stderr in the toolkit log to see the build error. Fix the app source. Re-run deploy.

## `ZipDeployFailed` (Kudu 4xx)

**Trigger.** Zip-deploy POST to Kudu returned 4xx.

**Mitigation.** Common causes:

- Wrong `resourceId` in the YAML — the App Service / Function App was renamed or deleted. Re-run provision.
- Locked deployment slot. Use `az webapp deployment list --name <name>` to inspect.
- File too large or filename invalid (Windows path length, special characters).

## `ZipDeployTimeout`

**Trigger.** Kudu accepted the upload but deployment didn't complete within the timeout.

**Mitigation.** Re-run with `--debug` to see the polling URL. Inspect deployment status in Azure portal. Large bundles benefit from `.deployment` exclusions.

## `InvalidZipPackage`

**Trigger.** The zipped artefact's first 4 bytes are not `PK\x03\x04` (magic-byte validation).

**Mitigation.** Something corrupted the build output. Clean and rebuild (`rm -rf node_modules build && npm install && npm run build`).

## App deployed but doesn't respond

Out of toolkit scope — check:

- App Service / Function App diagnostic logs.
- Application Insights for runtime errors.
- `BOT_ENDPOINT` env var matches the deployed hostname.
- Bot Framework channel registration points at the right endpoint.

## Deploy is *not* idempotent

Each deploy is a fresh push. Failed deploy may leave the previous revision running (App Service slot semantics). To roll back, redeploy the previous source.
