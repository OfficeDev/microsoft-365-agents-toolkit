// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

export enum SummaryConstant {
  Succeeded = "(√) Done:",
  Failed = "(×) Error:",
  NotExecuted = "(!) Warning:",
  Warning = "(!) Warning:",
}

export const component = "ConfigManager";

export const lifecycleExecutionEvent = "lifecycle-execution";

export enum TelemetryProperty {
  Lifecycle = "lifecycle",
  Actions = "actions",
  ResolvedPlaceholders = "resolved",
  UnresolvedPlaceholders = "unresolved",
  FailedAction = "failed-action",
}

export const MicrosoftEntraAuthType = "MicrosoftEntra";
export const OAuthAuthType = "OAuthPluginVault";
export const APIKeyAuthType = "ApiKeyPluginVault";
