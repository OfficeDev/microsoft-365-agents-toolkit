# Deploy flow

```
User → "Deploy" (tree view / palette / atk deploy --env dev)
  ↓
Engine entry: FxCore.deployArtifacts (v3) OR deployOp (v4)
  ↓
loadEnv(envName) → parseProjectYaml (deploy section) → resolveLifecycle
  ↓
confirmDeploy(envName)   # skipped for local / testtool / playground / sandbox
  ↓
executeLifecycle(deploySteps, envMap):
  Typical chain for a TS bot on App Service:
    cli/runNpmCommand   { script: "build" }
    azureAppService/zipDeploy  { resourceId: ${{BOT_AZURE_APP_SERVICE_RESOURCE_ID}}, artifactFolder: "./" }
  Typical chain for a Python bot on Functions:
    azureFunctions/zipDeploy   { resourceId: ${{FUNCTION_APP_RESOURCE_ID}}, artifactFolder: "./" }
  Typical chain for a C# bot:
    cli/runDotnetCommand { script: "publish -c Release" }
    azureAppService/zipDeploy  { ... }
  ↓
persistEnv(envName, envMap)
  ↓
PostAction[]: "App deployed", "Open browser to {url}"
```

## Speed considerations

- Builds are user-controlled (`npm run build`, `dotnet publish`) — speed depends on the app, not the toolkit.
- Zip-deploy uploads via Kudu over HTTPS; large bundles may benefit from `.deployment` or `.funcignore` exclusions in the user's project.

## Failure modes

| Failure | Cause | Note |
|---------|-------|------|
| Build failure | App-level | `cli/runNpmCommand` surfaces stdout / stderr |
| 4xx from Kudu | Wrong `resourceId` (provision skipped?) | Re-run provision |
| 5xx from Kudu | Service throttling / outage | `sendWithRetry` retries 5xx with exponential backoff |
| Deploy succeeds, app crashes | App-level | `kudu` logs, App Insights — out of toolkit scope |
