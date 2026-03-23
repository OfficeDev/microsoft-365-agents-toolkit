// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { TOOLS } from "../../../common/globalVars";
import { useLocalTemplate } from "../../../component/generator/templateHelper";
import { getTemplatesFolder } from "../../../folder";
import { constructNode } from "../constructNode";

/**
 * Load the wizard question tree from wizardNode.json.
 * This is a combined JSON containing root project type options and all inlined sub-trees
 * (DA, graphConnector, officeAddin). CEA and Teams nodes are still referenced via node names
 * and loaded from their own JSON files.
 */
export function getRootProjectTypeNode(platform: Platform = Platform.VSCode): IQTreeNode {
  const fileName = "wizardNode.json";
  const cachedJsonPath = path.join(os.homedir(), `.${String(ConfigFolderName)}`, "ui", fileName);

  let jsonPath: string;
  let source: string;
  if (!useLocalTemplate() && fs.pathExistsSync(cachedJsonPath)) {
    jsonPath = cachedJsonPath;
    source = "cache";
  } else {
    jsonPath = path.join(getTemplatesFolder(), "ui", fileName);
    source = "bundled";
  }

  const content = fs.readFileSync(jsonPath, "utf-8");
  TOOLS?.logProvider?.info(`[Dynamic Template] Loaded ${fileName} from ${source}: ${jsonPath}`);
  return constructNode(content, platform);
}
