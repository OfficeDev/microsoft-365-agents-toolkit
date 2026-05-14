// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Types
export {
  ManifestType,
  type EnvResolutionResult,
  type EnvPlaceholder,
  type ManifestTelemetryProperties,
} from "./types";

// Resolution engine
export {
  resolveEnvPlaceholders,
  getEnvVariables,
  expandFunctionExpressions,
  resolveManifest,
} from "./resolve";

// Manifest I/O
export { getManifestPath, readTeamsManifest, readAndResolveTeamsManifest } from "./readManifest";

// Validation
export { validateManifestSchema } from "./validate";

// Telemetry
export { parseCommonTelemetryProperties } from "./telemetry";
