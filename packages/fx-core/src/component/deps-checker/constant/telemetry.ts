// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export enum DepsCheckerEvent {
  nodeVersion = "node-version",
  nodeNotFound = "node-not-found",
  nodeNotSupportedForProject = "node-not-supported-for-project",
  nodeNotLts = "node-not-lts",

  dotnetCheckSkipped = "dotnet-check-skipped",
  dotnetAlreadyInstalled = "dotnet-already-installed",
  dotnetInstallCompleted = "dotnet-install-completed",
  dotnetInstallError = "dotnet-install-error",
  dotnetInstallScriptCompleted = "dotnet-install-script-completed",
  dotnetInstallScriptError = "dotnet-install-script-error",
  dotnetValidationError = "dotnet-validation-error",
  dotnetSearchDotnetSdks = "dotnet-search-dotnet-sdks",
}

export enum TelemtryMessages {
  failedToExecDotnetScript = "failed to exec dotnet script.",
  failedToValidateDotnet = "failed to validate dotnet.",
  failedToSearchDotnetSdks = "failed to search dotnet sdks.",
}

export enum TelemetryProperties {
  SymlinkFuncVersion = "symlink-func-version",
  SelectedPortableFuncVersion = "selected-portable-func-version",
  HistoryFuncVersion = "history-func-version",
  VersioningFuncVersions = "versioning-func-versions",
  GlobalFuncVersion = "global-func-version",
  InstalledFuncVersion = "installed-func-version",
  SymlinkFuncVersionError = "symlink-func-version-error",
  HistoryFuncVersionError = "history-func-version-error",
  VersioningFuncVersionError = "versioning-func-version-error",
  GlobalFuncVersionError = "global-func-version-error",
  InstallFuncError = "install-func-error",

  InstalledTestToolVersion = "installed-test-tool-version",
  InstallTestToolError = "install-test-tool-error",
  InstallTestToolReleaseType = "install-test-tool-release-type",
  InstallTestToolVersionFallback = "install-test-tool-version-fallback",
  InstallTestToolVersionError = "install-test-tool-version-error",
  SymlinkTestToolVersion = "symlink-test-tool-version",
  SymlinkTestToolVersionError = "symlink-test-tool-version-error",
  SelectedPortableTestToolVersion = "selected-test-tool-version",
  VersioningTestToolVersionError = "versioning-test-tool-version-error",
  VersioningTestToolVersionFallback = "versioning-test-tool-version-fallback",
  GlobalTestToolVersion = "global-test-tool-version",
  GlobalTestToolVersionError = "global-test-tool-version-error",
  GetFallbackVersionError = "get-fallback-version-error",
  TestToolLastUpdateTimestamp = "test-tool-last-update-timestamp",
  TestToolUpdatePreviousVersion = "test-tool-update-previous-version",
  TestToolUpdateError = "test-tool-update-error",
}

export enum TelemetryMessurement {
  completionTime = "completion-time",
}
