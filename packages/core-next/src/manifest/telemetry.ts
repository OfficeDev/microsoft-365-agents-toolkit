// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TeamsManifestWrapper } from "@microsoft/app-manifest";
import type { ManifestTelemetryProperties } from "./types";

/**
 * Extract common telemetry properties from a parsed Teams app manifest.
 *
 * Port of fx-core ManifestUtils.parseCommonTelemetryProperties(), adapted to
 * work with `TeamsManifestWrapper` instead of the raw JSON type.
 */
export function parseCommonTelemetryProperties(
  manifest: TeamsManifestWrapper
): ManifestTelemetryProperties {
  const capabilities: string[] = [];

  if (manifest.staticTabs.length > 0) capabilities.push("staticTab");
  if (manifest.configurableTabs.length > 0) capabilities.push("configurableTab");
  if (manifest.bots.length > 0) capabilities.push("Bot");
  if (manifest.composeExtensions.length > 0) capabilities.push("MessageExtension");
  if (manifest.declarativeAgents.length > 0) capabilities.push("copilotGpt");

  // Detect API-based compose extension
  const isApiME = manifest.composeExtensions.some(
    (ext) => (ext as Record<string, unknown>).type === "apiBased"
  );

  // Detect SPFx (web application info with resource that is NOT an api:// URI)
  const webAppInfo = manifest.webApplicationInfo;
  const isSPFx =
    webAppInfo !== undefined &&
    typeof webAppInfo.resource === "string" &&
    !webAppInfo.resource.startsWith("api://");

  // Detect API ME with AAD auth
  const isApiMeAAD = isApiME && webAppInfo !== undefined && typeof webAppInfo.id === "string";

  // Check for plugins referenced through copilotAgents
  const copilotAgents = manifest.copilotAgents;
  if (copilotAgents && (copilotAgents as Record<string, unknown>).plugins) {
    capabilities.push("plugin");
  }

  return {
    id: manifest.id,
    version: manifest.version,
    capabilities: capabilities.join(";"),
    manifestVersion: manifest.manifestVersion,
    isApiME: String(isApiME),
    isSPFx: String(isSPFx),
    isApiMeAAD: String(isApiMeAAD),
  };
}
