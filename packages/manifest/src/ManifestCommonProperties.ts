// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @deprecated This interface is used internally for manifest property extraction.
 * Consider using the auto-generated types from `./generated-types` for manifest definitions.
 */
export interface ManifestCommonProperties {
  /**
   * Capabilities, e.g. "staticTab" | "configurableTab" | "MessageExtension" | "WebApplicationInfo" | "plugin" | "copilotGpt"
   */
  capabilities: string[];
  /**
   * Teams app id
   */
  id: string;
  /**
   * Teams app version, e.g.1.0.0
   */
  version: string;
  /**
   * manifest version, e.g. 1.16
   */
  manifestVersion: string;
  /**
   * Whether it's an API ME
   */
  isApiME: boolean;
  /**
   * Whether it's SPFx Teams app
   */
  isSPFx: boolean;
  /**
   * Whether it's an API ME with AAD auth
   */
  isApiMeAAD: boolean;
}
