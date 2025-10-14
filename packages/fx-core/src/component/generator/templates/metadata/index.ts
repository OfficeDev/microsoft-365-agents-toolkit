// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { customEngineAgentTemplates } from "./customEngineAgent";
import { declarativeAgentTemplates } from "./da";
import { graphConnectorTemplates } from "./graphConnector";
import { Template } from "./interface";
import { messagingExtensionTemplates } from "./me";
import { specialTemplates } from "./special";
import { teamsAgentsAndAppsTemplates } from "./teams";
import { vsOnlyTemplates } from "./vs";
import { wxpTemplates } from "./wxp";

const allTemplates: Template[] = [
  ...declarativeAgentTemplates,
  ...customEngineAgentTemplates,
  ...teamsAgentsAndAppsTemplates,
  ...messagingExtensionTemplates,
  ...specialTemplates,
  ...vsOnlyTemplates,
  ...wxpTemplates,
  ...graphConnectorTemplates,
];

const defaultGeneratorTemplates: Template[] = [
  ...customEngineAgentTemplates,
  ...teamsAgentsAndAppsTemplates,
  ...messagingExtensionTemplates,
  ...vsOnlyTemplates,
  ...graphConnectorTemplates,
];

// used by programming language question options filter
export function getAllTemplatesOnPlatform(platform: Platform): Template[] {
  switch (platform) {
    case Platform.VSCode:
      return allTemplates.filter((t) => t.language !== "csharp");
    case Platform.VS:
      return allTemplates.filter((t) => t.language === "csharp");
    case Platform.CLI:
      return allTemplates;
    default:
      return [];
  }
}

// used by default generator
export function getDefaultTemplatesOnPlatform(platform: Platform): Template[] {
  switch (platform) {
    case Platform.VSCode:
      return defaultGeneratorTemplates.filter((t) => t.language !== "csharp");
    case Platform.VS:
      return defaultGeneratorTemplates.filter((t) => t.language === "csharp");
    case Platform.CLI:
      return defaultGeneratorTemplates;
    default:
      return [];
  }
}
