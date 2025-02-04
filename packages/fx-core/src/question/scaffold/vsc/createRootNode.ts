// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode, OptionItem } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { ProgrammingLanguage, QuestionNames } from "../../constants";
import { appNameQuestion, folderQuestion } from "../../create";
import { Templates } from "../../templates";
import { ApiPluginStartOptions } from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { customEngineAgentProjectTypeNode } from "./customAgentProjectTypeNode";
import { daProjectTypeNode } from "./daProjectTypeNode";
import { botProjectTypeNode, meProjectTypeNode, tabProjectTypeNode } from "./m365ProjectTypeNode";
import { officeAddinProjectTypeNode } from "./officeAddinProjectTypeNode";

export const LanguageOptionMap = new Map<string, OptionItem>([
  [ProgrammingLanguage.JS, { id: ProgrammingLanguage.JS, label: "JavaScript" }],
  [ProgrammingLanguage.TS, { id: ProgrammingLanguage.TS, label: "TypeScript" }],
  [ProgrammingLanguage.CSharp, { id: ProgrammingLanguage.CSharp, label: "C#" }],
  [ProgrammingLanguage.PY, { id: ProgrammingLanguage.PY, label: "Python" }],
]);

export function languageNode(): IQTreeNode {
  return {
    condition: (inputs: Inputs) => {
      const templateName = inputs[QuestionNames.TemplateName];
      const languages = Templates.filter((t) => t.name === templateName)
        .map((t) => t.language)
        .filter((lang) => lang !== "none" && lang !== undefined);
      return languages.length > 0;
    },
    data: {
      type: "singleSelect",
      title: getLocalizedString("core.ProgrammingLanguageQuestion.title"),
      name: QuestionNames.ProgrammingLanguage,
      staticOptions: [
        { id: ProgrammingLanguage.JS, label: "JavaScript" },
        { id: ProgrammingLanguage.TS, label: "TypeScript" },
        { id: ProgrammingLanguage.CSharp, label: "C#" },
        { id: ProgrammingLanguage.PY, label: "Python" },
      ],
      dynamicOptions: (inputs: Inputs) => {
        const templateName = inputs[QuestionNames.TemplateName];
        const languages = Templates.filter((t) => t.name === templateName)
          .map((t) => t.language)
          .filter((lang) => lang !== "none" && lang !== undefined);
        return languages.map((lang) => LanguageOptionMap.get(lang) as OptionItem);
      },
      skipSingleOption: true,
    },
  };
}

export function languageCondition(inputs: Inputs): boolean {
  // Only skip this project when need to rediect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec 3. No plugin manifest path
  return !(
    featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
    inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
    inputs[QuestionNames.ProjectType] === ProjectTypeOptions.copilotAgentOptionId &&
    !inputs[QuestionNames.ApiPluginManifestPath]
  );
}

/**
 *
 * FxCore API for scaffold: scaffold(questionModel: IQTreeNode, generators: DefaultTemplateGenerator[]): Promise<Result<any, FxError>>
 * Dedicated for VS Code platform
 */

export function scaffoldQuestionForVSCode(): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        data: {
          name: QuestionNames.ProjectType,
          title: getLocalizedString("core.createProjectQuestion.title"),
          type: "singleSelect",
          staticOptions: [
            ProjectTypeOptions.Agent(),
            ProjectTypeOptions.customCopilot(),
            ProjectTypeOptions.bot(),
            ProjectTypeOptions.tab(),
            ProjectTypeOptions.me(),
            ProjectTypeOptions.officeAddin(),
            ...(featureFlagManager.getBooleanValue(FeatureFlags.ChatParticipantUIEntries)
              ? [ProjectTypeOptions.startWithGithubCopilot()]
              : []),
          ],
        },
        children: [
          daProjectTypeNode(),
          customEngineAgentProjectTypeNode(),
          botProjectTypeNode(),
          tabProjectTypeNode(),
          meProjectTypeNode(),
          officeAddinProjectTypeNode(),
        ],
      },
      languageNode(),
      {
        condition: languageCondition,
        data: {
          type: "group",
        },
        children: [
          {
            data: folderQuestion(),
          },
          {
            data: appNameQuestion(),
          },
        ],
      },
    ],
  };
  return node;
}
