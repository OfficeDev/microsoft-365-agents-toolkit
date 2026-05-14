# Publish flow

```
User → "Publish to Tenant" (tree view / palette / atk publish --env dev)
  ↓
Engine entry: FxCore.publishApplication (v3) OR publishOp (v4)
  ↓
loadEnv(envName) → parseProjectYaml (publish section) → resolveLifecycle
  ↓
executeLifecycle(publishSteps, envMap):
  teamsApp/validateManifest        # schema check 2.4
  teamsApp/zipAppPackage           # bundle manifest + icons
  teamsApp/validateAppPackage      # full package validation via TDP
  teamsApp/publishAppPackage       # POST /beta/appCatalogs/teamsApps via MS Graph
  optional: teamsApp/extendToM365  # sideload V2 DA into M365 ecosystem
  ↓
persistEnv(envName, envMap)
  ↓
PostAction[]: "Open admin center to approve", helpLink
```

## What "publish" actually does

- For a Teams app: posts the app package to the org's app catalog. The app shows up in the admin center for approval. It does **not** automatically publish to the Microsoft Teams Store.
- For a Declarative Agent: same code path; the DA shows up in M365 Copilot once admin approves.

## Re-publish

`teamsApp/publishAppPackage` is **idempotent** — if the app already exists in the catalog (matched by `Teams App ID`), it updates the existing entry via Graph PUT instead of creating a new one.

## Failure modes

| Failure | Likely cause | Mitigation |
|---------|-------------|-----------|
| Manifest validation fails | Schema drift / missing field | `teamsApp/validateManifest` returns Zod-style issue path |
| Package validation fails | Icon size, manifest cross-check | TDP returns specific error; surfaced verbatim |
| 403 on Graph publish | User lacks `AppCatalog.ReadWrite.All` | Surfaced as `UserError` with help link to admin permissions doc |
| `extendToM365` failure (V2 DA) | Sideloading disabled in tenant | `UserError` with admin-policy help link |
