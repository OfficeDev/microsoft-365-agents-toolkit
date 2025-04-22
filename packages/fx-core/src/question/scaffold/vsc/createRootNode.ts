// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getAllTemplatesOnPlatform } from "../../../component/generator/templates/metadata";
import { ProgrammingLanguage, QuestionNames } from "../../constants";
import {
  appNameQuestion,
  folderQuestion,
  GCConnectionIdQuestion,
  GCNameQuestion,
} from "../../create";
import {
  ActionStartOptions,
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
import { graphConnectorProjectTypeNode } from "./graphConnectorProjectTypeNode";

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
    .filter((lang) => lang !== undefined);
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
  // skip this project when need to redirect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec 3. No plugin manifest path
  // or start with github copilot
  return (
    !(
      featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
      inputs[QuestionNames.ActionType] === ActionStartOptions.apiSpec().id &&
      (inputs[QuestionNames.ProjectType] === ProjectTypeOptions.copilotAgentOptionId ||
        inputs[QuestionNames.Capabilities] === DACapabilityOptions.declarativeAgent().id) &&
      !inputs[QuestionNames.ActionManifestPath]
    ) && inputs[QuestionNames.ProjectType] !== ProjectTypeOptions.startWithGithubCopilotOptionId
  );
}

/**
 * Scaffold question model dedicated for VS Code platform
 */
export function scaffoldQuestionForVSCode(platform: Platform = Platform.VSCode): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        data: {
          name: QuestionNames.ProjectType,
          title: getLocalizedString("core.createProjectQuestion.title"),
          type: "singleSelect",
          staticOptions: [
            ProjectTypeOptions.declarativeAgent(platform),
            ProjectTypeOptions.customEngineAgent(platform),
            ProjectTypeOptions.graphConnector(platform),
            ProjectTypeOptions.bot(platform),
            ProjectTypeOptions.tab(platform),
            ProjectTypeOptions.me(platform),
            ProjectTypeOptions.officeAddin(platform),
            ...(featureFlagManager.getBooleanValue(FeatureFlags.ChatParticipantUIEntries)
              ? [ProjectTypeOptions.startWithGithubCopilot()]
              : []),
          ],
        },
        children: [
          daProjectTypeNode(),
          customEngineAgentProjectTypeNode(),
          graphConnectorProjectTypeNode(),
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
          {
            condition: (inputs: Inputs) => {
              return inputs[QuestionNames.WithPlugin] === DACapabilityOptions.withGC().id;
            },
            data: {
              type: "group",
            },
            children: [
              {
                data: GCNameQuestion(),
              },
              {
                data: GCConnectionIdQuestion(),
              },
            ],
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
