// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { TOOLS } from "../../../common/globalVars";
import * as templateHelper from "../../../component/generator/templateHelper";
import * as folder from "../../../folder";
import { constructNode } from "../constructNode";

/**
 * Load the wizard question tree from wizardNode.json.
 * Combined JSON with all sub-trees inlined.
 */
export function getRootProjectTypeNode(platform: Platform = Platform.VSCode): IQTreeNode {
  return loadUiNode("wizardNode.json", platform);
}

/**
 * Load the TDP wizard question tree from tdpNode.json.
 * Subset of wizard options for Teams Developer Portal import flow.
 */
export function getTdpProjectTypeNode(platform: Platform = Platform.VSCode): IQTreeNode {
  return loadUiNode("tdpNode.json", platform);
}

function loadUiNode(fileName: string, platform: Platform): IQTreeNode {
  const cachedJsonPath = path.join(os.homedir(), `.${String(ConfigFolderName)}`, "ui", fileName);

  let jsonPath: string;
  let source: string;
  if (
    !templateHelper.useLocalTemplate() &&
    !templateHelper.useBundledMetadataForV4() &&
    fs.pathExistsSync(cachedJsonPath)
  ) {
    jsonPath = cachedJsonPath;
    source = "cache";
  } else {
    jsonPath = path.join(folder.getTemplatesFolder(), "ui", fileName);
    source = "bundled";
  }

  const content = fs.readFileSync(jsonPath, "utf-8");
  TOOLS?.logProvider?.info(`[Dynamic Template] Loaded ${fileName} from ${source}: ${jsonPath}`);
  return constructNode(content, platform);
}
