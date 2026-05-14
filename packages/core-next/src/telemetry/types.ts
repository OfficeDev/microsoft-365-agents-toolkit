// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Telemetry event names for implemented core-next operations.
 * Each operation emits "{name}-start" and "{name}-end" events via runOperation().
 * These constants cover additional cross-cutting events.
 */
export enum TelemetryEvent {
  // Project
  CreateProject = "create-project",
  Scaffold = "scaffold",

  // Lifecycle
  LifecycleProvision = "lifecycle-provision",
  LifecycleExecute = "lifecycle-execute",

  // Declarative Agent
  AddKnowledge = "da-add-knowledge",
  AddExistingPlugin = "da-add-existing-plugin",
  AddMCPAction = "da-add-mcp-action",
  RemoveAction = "da-remove-action",
  SetSensitivityLabel = "da-set-sensitivity-label",
  SetConversationStarters = "da-set-conversation-starters",
  InjectOAuthAction = "da-inject-oauth-action",
  InjectApiKeyAction = "da-inject-api-key-action",

  // Teams App
  ValidateManifest = "validate-manifest",
  PackageApp = "package-app",
  PublishApp = "publish-app",
}

/**
 * Well-known telemetry property keys.
 */
export enum TelemetryProperty {
  CorrelationId = "correlation-id",
  Success = "success",
  Duration = "duration",
  ErrorCode = "error-code",
  ErrorKind = "error-kind",
  ErrorMessage = "error-message",
  Component = "component",
  OperationName = "operation-name",
}

/**
 * Standard success values.
 */
export enum TelemetrySuccess {
  Yes = "yes",
  No = "no",
}
