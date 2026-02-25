// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { llmServiceNode } from "../commonNodes";
import { QuestionNames } from "../../questionNames";
import { Agent365Options, setTemplateName } from "./CapabilityOptions";

export function getAgent365TypeNode(): IQTreeNode {
  return {
    condition: { equals: ProjectTypeOptions.agent365OptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: "Agent 365",
      placeholder: "Please select your Agent 365 capability",
      type: "singleSelect",
      staticOptions: [Agent365Options.basic()],
      skipSingleOption: true,
      onDidSelection: setTemplateName,
    },
    children: [llmServiceNode()],
  };
}
