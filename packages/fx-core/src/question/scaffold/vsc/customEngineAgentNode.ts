// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import { CustomEngineAgentOptions, setTemplateName } from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { llmServiceNode } from "./agentForTeamsNode";

export function customEngineAgentNode(): IQTreeNode {
  return {
    // project-type = Custom Engine Agent
    condition: { equals: ProjectTypeOptions.customEngineAgentOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.customCopilot.title"),
      type: "singleSelect",
      staticOptions: [CustomEngineAgentOptions.basicCustomEngineAgent()],
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.customCopilot.placeholder"
      ),
      onDidSelection: setTemplateName,
    },
    children: [llmServiceNode()],
  };
}
