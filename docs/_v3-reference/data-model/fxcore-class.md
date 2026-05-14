# `FxCore` class — v3 entry point

> **FORBIDDEN AS v4 DESIGN INPUT.** This is precisely the kind of god-class shape v4 was created to escape. See [`../README.md`](../README.md).

Source: [`packages/fx-core/src/core/FxCore.ts`](../../../packages/fx-core/src/core/FxCore.ts). **57 public methods** on a single class (56 if you exclude the `on()` event handler).

The list below is exhaustive at the time of extraction. It will drift as v3 evolves — the source file is the truth.

## Public methods (signatures)

```typescript
on(event: CoreCallbackEvent, callback: CoreCallbackFunc): void

createProject(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>>
createProjectFromTdp(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>>
createProjectByCustomizedGenerator(...): Promise<Result<CreateProjectResult, FxError>>
createSampleProject(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>>

provisionResources(inputs: Inputs): Promise<Result<undefined, FxError>>
provisionResourcesOnce(...): Promise<Result<undefined, FxError>>
deployArtifacts(...): Promise<Result<undefined, FxError>>
publishApplication(...): Promise<Result<undefined, FxError>>
localDebug(inputs: Inputs): Promise<Result<undefined, FxError>>

deployAadManifest(inputs: Inputs): Promise<Result<undefined, FxError>>
buildAadManifest(inputs: Inputs): Promise<Result<undefined, FxError>>
convertAadToNewSchema(inputs: Inputs): Promise<Result<undefined, FxError>>

deployTeamsManifest(...): Promise<Result<undefined, FxError>>
updateTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>>
validateTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>>
packageTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>>
publishTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>>

validateApplication(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>>
validateManifest(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>>
validateAppPackage(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>>
validateWithTestCases(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>>

syncManifest(inputs: SyncManifestInputs): Promise<Result<any, FxError>>
createAppPackage(inputs: Inputs): Promise<Result<any, FxError>>
previewWithManifest(inputs: Inputs): Promise<Result<string, FxError>>

uninstall(inputs: UninstallInputs): Promise<Result<undefined, FxError>>
uninstallByManifestId(inputs: UninstallInputs): Promise<Result<undefined, FxError>>
uninstallByEnv(...): Promise<Result<undefined, FxError>>
uninstallByTitleId(inputs: UninstallInputs): Promise<Result<undefined, FxError>>
uninstallM365App(...): Promise<Result<undefined, FxError>>
uninstallAppRegistration(manifestId: string): Promise<Result<undefined, FxError>>
uninstallBotFrameworRegistration(...): Promise<Result<undefined, FxError>>

shareApplication(...): Promise<Result<undefined, FxError>>
removeSharedAccess(...): Promise<Result<undefined, FxError>>

executeUserTask(func: Func, inputs: Inputs): Promise<Result<any, FxError>>
addWebpart(inputs: Inputs): Promise<Result<undefined, FxError>>

getDotEnvs(...): Promise<Result<Map<string, string>, FxError>>
isEnvFile(projectPath: string, inputFile: string): Promise<Result<boolean, FxError>>
getProjectId(projectPath: string): Promise<Result<string, FxError>>
getProjectMetadata(...): Promise<Result<ProjectMetadata, FxError>>
getTeamsAppName(projectPath: string): Promise<Result<string, FxError>>
getProjectInfo(...): Promise<Result<Map<string, string>, FxError>>

grantPermission(inputs: Inputs): Promise<Result<PermissionsResult, FxError>>
checkPermission(inputs: Inputs): Promise<Result<PermissionsResult, FxError>>
listCollaborator(inputs: Inputs): Promise<Result<ListCollaboratorResult, FxError>>

createLocalCrypto(projectPath: string): Promise<Result<CryptoProvider, FxError>>
encrypt(plaintext: string, inputs: Inputs): Promise<Result<string, FxError>>
decrypt(ciphertext: string, inputs: Inputs): Promise<Result<string, FxError>>

createEnv(inputs: Inputs): Promise<Result<undefined, FxError>>
createEnvCopyV3(...): Promise<Result<undefined, FxError>>
phantomMigrationV3(inputs: Inputs): Promise<Result<undefined, FxError>>
```

## Notes

- All methods take and return `Inputs`-shaped bags rather than typed inputs per operation.
- All async operations route through the `Coordinator` and middleware chain.
- The class also holds a `tools` reference (`TOOLS` global singleton) — see [api-package.md](api-package.md).

## Why this is forbidden as v4 design input

This shape is the v4 anti-target. v4 replaces it with `Operation` records produced by `defineOperation(name, schema, fn)` and executed via `runOperation(op, input, ctx)` — typed per operation, no class state, no `TOOLS` singleton, composable.
