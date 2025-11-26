// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, Platform } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import os from "os";
import path from "path";
import templateConfig from "../../../../common/templates-config.json";
import { getTemplatesFolder } from "../../../../folder";
import { Template } from "./interface";

function getTemplateMetadataConfig(configName: string): Template[] {
  let jsonPath: string;

  const cachedJsonPath = path.join(
    os.homedir(),
    `.${String(ConfigFolderName)}`,
    "metadata",
    configName
  );

  // Check if cached JSON exists, otherwise fallback to bundled templates folder
  if (!templateConfig.useLocalTemplate && cachedJsonPath && fs.pathExistsSync(cachedJsonPath)) {
    jsonPath = cachedJsonPath;
  } else {
    jsonPath = path.join(getTemplatesFolder(), "metadata", configName);
  }

  const content = fs.readFileSync(jsonPath, "utf-8");
  return JSON.parse(content) as Template[];
}

// used by programming language question options filter
export function getAllTemplatesOnPlatform(platform: Platform): Template[] {
  const allTemplates = getTemplateMetadataConfig("allTemplates.json");
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
  const defaultGeneratorTemplates = getTemplateMetadataConfig("defaultGeneratorTemplates.json");
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
