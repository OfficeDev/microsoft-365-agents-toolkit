// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, IQTreeNode } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { useLocalTemplate } from "../../../component/generator/templateHelper";
import { getTemplatesFolder } from "../../../folder";
import { constructNode } from "../constructNode";

export function getCustomEngineAgentNode(): IQTreeNode {
  let jsonPath: string;

  const cachedJsonPath = path.join(
    os.homedir(),
    `.${String(ConfigFolderName)}`,
    "ui",
    "ceaNode.json"
  );

  // Check if cached JSON exists, otherwise fallback to bundledtemplates folder
  if (!useLocalTemplate() && fs.pathExistsSync(cachedJsonPath)) {
    jsonPath = cachedJsonPath;
  } else {
    jsonPath = path.join(getTemplatesFolder(), "ui", "ceaNode.json");
  }

  const content = fs.readFileSync(jsonPath, "utf-8");
  return constructNode(content);
}
