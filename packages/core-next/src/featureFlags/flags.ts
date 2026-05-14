// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FeatureFlag } from "./types";
import { FeatureFlagRegistry } from "./registry";

/**
 * Built-in feature flags relevant to core-next operations.
 * Only flags that gate core-next behaviour are included here;
 * UI-only or CLI-only flags stay in their respective packages.
 */
export const builtinFlags: FeatureFlag[] = [
  {
    name: "TEAMSFX_V4_CORE",
    defaultValue: false,
    description: "Enable v4 core engine (core-next)",
  },
  {
    name: "TEAMSFX_MCP_FOR_DA",
    defaultValue: true,
    description: "Enable MCP action support for Declarative Agents",
  },
  {
    name: "TEAMSFX_SENSITIVITY_LABEL",
    defaultValue: false,
    description: "Enable sensitivity label support for DA manifests",
  },
  {
    name: "TEAMSFX_DA_METAOS",
    defaultValue: false,
    description: "Enable MetaOS variant for Declarative Agents",
  },
  {
    name: "TEAMSFX_CEA_ENABLED",
    defaultValue: false,
    description: "Enable Custom Engine Agent features",
  },
  {
    name: "TEAMSFX_GENERATE_CONFIG_FILES",
    defaultValue: false,
    description: "Generate config files during scaffolding",
  },
  {
    name: "TEAMSFX_TELEMETRY_TEST",
    defaultValue: false,
    description: "Route telemetry to test endpoint",
  },
];

/**
 * Create a FeatureFlagRegistry pre-populated with built-in flags.
 */
export function createDefaultRegistry(): FeatureFlagRegistry {
  const registry = new FeatureFlagRegistry();
  registry.registerAll(builtinFlags);
  return registry;
}
