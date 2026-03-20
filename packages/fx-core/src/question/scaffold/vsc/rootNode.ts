// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { useLocalTemplate } from "../../../component/generator/templateHelper";
import { getTemplatesFolder } from "../../../folder";
import { constructNode } from "../constructNode";

/**
 * Load the root question node from rootNode.json.
 * This defines the project type options (DA, CEA, Teams, etc.)
 * on the first page of the "Create New Agent/App" wizard.
 */
export function getRootProjectTypeNode(platform: Platform = Platform.VSCode): IQTreeNode {
  let jsonPath: string;

  const cachedJsonPath = path.join(
    os.homedir(),
    `.${String(ConfigFolderName)}`,
    "ui",
    "rootNode.json"
  );

  // Check if cached JSON exists, otherwise fallback to bundled templates folder
  if (!useLocalTemplate() && fs.pathExistsSync(cachedJsonPath)) {
    jsonPath = cachedJsonPath;
  } else {
    jsonPath = path.join(getTemplatesFolder(), "ui", "rootNode.json");
  }

  const content = fs.readFileSync(jsonPath, "utf-8");
  return constructNode(content, platform);
}
