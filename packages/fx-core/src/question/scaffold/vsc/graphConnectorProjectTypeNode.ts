// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { GCNameQuestion } from "../../create";

export function graphConnectorProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Custom Engine Agent
    condition: { equals: ProjectTypeOptions.graphConnectorOptionId },
    data: GCNameQuestion(),
  };
}
