# v3 error catalogue

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).
>
> **Exception:** individual error *names* may be reused for back-compat (telemetry partition keys are stable across v3/v4). The catalogue's organising structure (file grouping, module boundaries) **must not** be reused.

Source: [`packages/fx-core/src/error/`](../../../packages/fx-core/src/error/). The list below is current at extraction time — the source files are the truth (this catalogue grows over time).

## Catalogue

| File | Errors exported |
|------|-----------------|
| `common.ts` | `FileNotFoundError`, `MissingEnvironmentVariablesError`, `InvalidActionInputError`, `InvalidProjectError`, `MultipleAuthError`, `MultipleServerError`, `InjectAPIKeyActionFailedError`, `InjectOAuthActionFailedError`, `DeclarativeAgentPathNotFoundError`, `ActionNotFoundError`, `SpecNotFoundError`, `OriginalSpecNotFoundError`, `JSONSyntaxError`, `ReadFileError`, `WriteFileError`, `FilePermissionError`, `UnhandledError`, `UnhandledUserError`, `InstallSoftwareError`, `MissingRequiredInputError`, `InputValidationError`, `NoEnvFilesError`, `MissingRequiredFileError`, `NetworkError`, `HttpClientError`, `HttpServerError`, `AccessGithubError`, `UserCancelError`, `NeedRedoError`, `EmptyOptionError`, `NotImplementedError`, `MFARequiredError`, `ConcurrentError`, `InternalError`, `NoProjectOpenedError`, `MigrationError`, `NotAllowedMigrationError`, `FailedToLoadManifestId`, `VideoFilterAppRemoteNotSupportedError`, `UpgradeV3CanceledError`, `IncompatibleProjectError`, `AbandonedProjectError`, `FailedToParseResourceIdError`, `NpmInstallError`, `FileNotSupportError`, `matchDnsError()` (~45 exports total) |
| `azure.ts` | `InvalidAzureCredentialError`, `InvalidAzureSubscriptionError`, `SelectSubscriptionError`, `ResourceGroupConflictError`, `ResourceGroupNotExistError`, `CreateResourceGroupError`, `CheckResourceGroupExistenceError`, `ListResourceGroupsError`, `GetResourceGroupError`, `ListResourceGroupLocationsError` |
| `teamsApp.ts` | `DeveloperPortalAPIFailedSystemError`, `DeveloperPortalAPIFailedUserError`, `CheckSideloadingPermissionFailedError`, `InvalidFileOutsideOfTheDirectotryError` (sic), `AppIdNotExist` (5 errors) |
| `deploy.ts` | `DeployEmptyFolderError`, `CheckDeploymentStatusTimeoutError`, `ZipFileError`, `CacheFileInUse`, `GetPublishingCredentialsError`, `DeployZipPackageError`, `CheckDeploymentStatusError`, `AzureStorageClearBlobsError`, `AzureStorageUploadFilesError`, `AzureStorageGetContainerError`, `AzureStorageGetContainerPropertiesError`, `AzureStorageSetContainerPropertiesError` |
| `yml.ts` | `InvalidYamlSchemaError`, `YamlFieldTypeError`, `YamlFieldMissingError`, `InvalidYmlActionNameError`, `LifeCycleUndefinedError` |
| `arm.ts` | `CompileBicepError`, `DeployArmError`, `GetArmDeploymentError`, `ConvertArmOutputError`, `DownloadBicepCliError` |
| `depCheck.ts` | `PortsConflictError`, `SideloadingDisabledError`, `CopilotDisabledError`, `NodejsNotLtsError`, `NodejsNotFoundError`, `NodejsNotRecommendedError`, `VxTestAppInvalidInstallOptionsError`, `VxTestAppValidationError`, `DepsCheckerError`, `FindProcessError`, `InstallNodeJSError` (11 errors) |
| `kiota.ts` | `KiotaGeneratePluginError` |
| `script.ts` | `ScriptTimeoutError`, `ScriptExecutionError` |
| `m365.ts` | `M365TokenJSONNotFoundError`, `M365TenantIdNotFoundInTokenError`, `M365TenantIdNotMatchError` |
| `upgrade.ts` | `NoNeedUpgradeError` |
| `types.ts` | `ErrorCategory` (enum) |

## Notes

- All extend `UserError` or `SystemError` from `@microsoft/teamsfx-api`.
- The split between `User`/`System` is per error name; some files mix both.
- Each error's `name` field is the telemetry partition key — they are stable across releases.

## Why the catalogue's structure is forbidden as v4 design input

The v3 grouping (one file per loosely-defined "domain") evolved organically. v4 does not need the same shape; what matters is:

- The error type discriminator (`UserError` / `SystemError` or v4's `AtkError` with `kind`).
- Stable `name` strings (v4 may reuse v3 names where the failure mode is genuinely the same).
- A shared `displayMessage` / `helpLink` / `innerError` shape.

v4 is free to organise errors per package, per module, or per operation — whatever fits the v4 architecture. Don't import the file grouping.
