// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { getTemplatesFolder } from "../../../../folder";
import { Template } from "./interface";

// const allTemplates: Template[] = [
//   ...declarativeAgentTemplates,
//   ...customEngineAgentTemplates,
//   ...teamsAgentsAndAppsTemplates,
//   ...messagingExtensionTemplates,
//   ...specialTemplates,
//   ...vsOnlyTemplates,
//   ...wxpTemplates,
//   ...graphConnectorTemplates,
// ];

// const defaultGeneratorTemplates: Template[] = [
//   ...customEngineAgentTemplates,
//   ...teamsAgentsAndAppsTemplates,
//   ...messagingExtensionTemplates,
//   ...vsOnlyTemplates,
//   ...graphConnectorTemplates,
// ];

let allTemplates: Template[] = [];
let defaultGeneratorTemplates: Template[] = [];

// used by programming language question options filter
export function getAllTemplatesOnPlatform(platform: Platform): Template[] {
  if (allTemplates.length == 0) {
    const data = fs.readFileSync(
      path.join(getTemplatesFolder(), "metadata", "allTemplates.json"),
      "utf-8"
    );
    allTemplates = JSON.parse(data) as Template[];
  }
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
  if (defaultGeneratorTemplates.length == 0) {
    const data = fs.readFileSync(
      path.join(getTemplatesFolder(), "metadata", "defaultGeneratorTemplates.json"),
      "utf-8"
    );
    defaultGeneratorTemplates = JSON.parse(data) as Template[];
  }
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
