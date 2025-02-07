// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { basicBotTemplates, notificationBotTemplates } from "./bot";
import { customEngineAgentTemplates } from "./cea";
import { copilotPluginTemplates } from "./copilotPlugin";
import { Template } from "./interface";
import { messagingExtensionTemplates } from "./me";
import { specialTemplates } from "./special";
import { tabTemplates } from "./tab";
import { tdpTemplates } from "./tdp";
import { vsOnlyTemplates } from "./vs";

// used by programming language question options filter
export const allTemplates: Template[] = [
  ...tabTemplates,
  ...basicBotTemplates,
  ...notificationBotTemplates,
  ...messagingExtensionTemplates,
  ...copilotPluginTemplates,
  ...customEngineAgentTemplates,
  ...tdpTemplates,
  ...specialTemplates,
];

const defaultGeneratorTemplates: Template[] = [
  ...tabTemplates,
  ...basicBotTemplates,
  ...notificationBotTemplates,
  ...messagingExtensionTemplates,
  ...copilotPluginTemplates,
  ...customEngineAgentTemplates,
  ...tdpTemplates,
  ...vsOnlyTemplates,
];
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
