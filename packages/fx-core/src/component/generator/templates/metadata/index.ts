// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { basicBotTemplates } from "./bot";
import { customEngineAgentTemplates } from "./cea";
import { Template } from "./interface";
import { messagingExtensionTemplates } from "./me";
import { specialTemplates } from "./special";
import { tabTemplates } from "./tab";
import { tdpTemplates } from "./tdp";
import { vsOnlyTemplates } from "./vs";

const allTemplates: Template[] = [
  ...tabTemplates,
  ...basicBotTemplates,
  ...messagingExtensionTemplates,
  ...customEngineAgentTemplates,
  ...tdpTemplates,
  ...specialTemplates,
  ...vsOnlyTemplates,
];

const defaultGeneratorTemplates: Template[] = [
  ...tabTemplates,
  ...basicBotTemplates,
  ...messagingExtensionTemplates,
  ...customEngineAgentTemplates,
  ...tdpTemplates,
  ...vsOnlyTemplates,
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
