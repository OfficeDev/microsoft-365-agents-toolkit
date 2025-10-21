// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { getTemplatesFolder } from "../../../folder";
import { constructNode } from "../constructNode";

export function getCustomEngineAgentNode(): IQTreeNode {
  const content = fs.readFileSync(path.join(getTemplatesFolder(), "ui", "ceaNode.json"), "utf-8");
  return constructNode(content);
}
