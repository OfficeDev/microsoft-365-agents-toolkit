// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode, OptionItem } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getAllTemplatesOnPlatform } from "../../../component/generator/templates/metadata";
import { ProgrammingLanguage, QuestionNames } from "../../constants";
import { appNameQuestion, folderQuestion } from "../../create";
import {
  ApiPluginStartOptions,
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  DACapabilityOptions,
  MeCapabilityOptions,
  OfficeAddinCapabilityOptions,
  TabCapabilityOptions,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { customEngineAgentProjectTypeNode } from "./customAgentProjectTypeNode";
import { daProjectTypeNode } from "./daProjectTypeNode";
import { officeAddinProjectTypeNode } from "./officeAddinProjectTypeNode";
import { botProjectTypeNode, meProjectTypeNode, tabProjectTypeNode } from "./teamsProjectTypeNode";

export const LanguageOptionMap = new Map<string, OptionItem>([
  [ProgrammingLanguage.JS, { id: ProgrammingLanguage.JS, label: "JavaScript" }],
  [ProgrammingLanguage.TS, { id: ProgrammingLanguage.TS, label: "TypeScript" }],
  [ProgrammingLanguage.CSharp, { id: ProgrammingLanguage.CSharp, label: "C#" }],
  [ProgrammingLanguage.PY, { id: ProgrammingLanguage.PY, label: "Python" }],
  [ProgrammingLanguage.Common, { id: ProgrammingLanguage.Common, label: "None" }],
  [ProgrammingLanguage.None, { id: ProgrammingLanguage.None, label: "None" }],
]);

export function getLanguageOptions(inputs: Inputs): OptionItem[] {
  const templateName = inputs[QuestionNames.TemplateName];
  const languages = getAllTemplatesOnPlatform(inputs.platform)
    .filter((t) => t.name === templateName)
    .map((t) => t.language)
    .filter((lang) => lang !== "none" && lang !== undefined);
  const languageOptions = languages.map(
    (lang) =>
      (LanguageOptionMap.get(lang) as OptionItem) || {
        id: ProgrammingLanguage.Common,
        label: "None",
      }
  );
  return languageOptions;
}

export function getDefaultLanguage(inputs: Inputs): string | undefined {
  const options = getLanguageOptions(inputs);
  return options[0]?.id;
}

export function languageNode(): IQTreeNode {
  return {
    condition: (inputs: Inputs) => {
      const templateName = inputs[QuestionNames.TemplateName];
      const languages = getAllTemplatesOnPlatform(inputs.platform)
        .filter((t) => t.name === templateName)
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
      dynamicOptions: getLanguageOptions,
      default: getDefaultLanguage,
      skipSingleOption: true,
    },
  };
}

export function folderAndAppNameCondition(inputs: Inputs): boolean {
  // Only skip this project when need to rediect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec 3. No plugin manifest path
  return !(
    featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
    inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
    (inputs[QuestionNames.ProjectType] === ProjectTypeOptions.copilotAgentOptionId ||
      inputs[QuestionNames.Capabilities] === DACapabilityOptions.declarativeAgent().id) &&
    !inputs[QuestionNames.ApiPluginManifestPath]
  );
}

/**
 * Scaffold question model dedicated for VS Code platform
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
            ProjectTypeOptions.declarativeAgent(),
            ProjectTypeOptions.customEngineAgent(),
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
        condition: folderAndAppNameCondition,
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

/**
 * CLI non-interactive mode has no "project-type" input, to make it compatible to the question model,
 * we need to convert capability to project type
 */
export function getProjectTypeByCapability(capability: string): string {
  if ([DACapabilityOptions.declarativeAgent().id].includes(capability)) {
    return ProjectTypeOptions.copilotAgentOptionId;
  }
  if (
    [
      CustomCopilotCapabilityOptions.basicChatbot().id,
      CustomCopilotCapabilityOptions.customCopilotRag().id,
      CustomCopilotCapabilityOptions.aiAgent().id,
    ].includes(capability)
  ) {
    return ProjectTypeOptions.customCopilotOptionId;
  }
  if (
    [
      BotCapabilityOptions.basicBot().id,
      BotCapabilityOptions.notificationBot().id,
      BotCapabilityOptions.commandBot().id,
      BotCapabilityOptions.workflowBot().id,
    ].includes(capability)
  ) {
    return ProjectTypeOptions.botOptionId;
  }
  if (
    [
      TabCapabilityOptions.nonSsoTab().id,
      TabCapabilityOptions.m365SsoLaunchPage().id,
      TabCapabilityOptions.dashboardTab().id,
      TabCapabilityOptions.SPFxTab().id,
    ].includes(capability)
  ) {
    return ProjectTypeOptions.tabOptionId;
  }
  if (
    [
      MeCapabilityOptions.m365SearchMe().id,
      MeCapabilityOptions.collectFormMe().id,
      MeCapabilityOptions.linkUnfurling().id,
    ].includes(capability)
  ) {
    return ProjectTypeOptions.meOptionId;
  }
  if (
    [
      OfficeAddinCapabilityOptions.wxpTaskPane().id,
      OfficeAddinCapabilityOptions.officeAddinImport().id,
    ].includes(capability)
  ) {
    return ProjectTypeOptions.officeMetaOSOptionId;
  }

  if (
    [
      OfficeAddinCapabilityOptions.outlookTaskPane().id,
      OfficeAddinCapabilityOptions.outlookAddinImport().id,
    ].includes(capability)
  ) {
    return ProjectTypeOptions.outlookAddinOptionId;
  }

  return "";
}
