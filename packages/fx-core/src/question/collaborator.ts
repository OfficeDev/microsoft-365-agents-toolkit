// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DynamicPlatforms,
  Inputs,
  IQTreeNode,
  MultiSelectQuestion,
  OptionItem,
  Platform,
} from "@microsoft/teamsfx-api";
import * as path from "path";
import { featureFlagManager, FeatureFlags } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";
import { CollaborationConstants, CollaborationUtil } from "../core/collaborator";
import { QuestionNames } from "./constants";
import {
  confirmManifestQuestion,
  inputUserEmailQuestion,
  selectAadManifestQuestion,
  selectTargetEnvQuestion,
  selectTeamsAppManifestQuestionNode,
} from "./other";

export function grantPermissionQuestionNode(): IQTreeNode {
  const selectAppManifestNode = selectTeamsAppManifestQuestionNode();
  selectAppManifestNode.condition = {
    contains: CollaborationConstants.TeamsAppQuestionId,
  };
  selectAppManifestNode.children?.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  const selectAadAppNode = selectAadAppManifestQuestionNode();
  selectAadAppNode.condition = { contains: CollaborationConstants.AadAppQuestionId };
  selectAadAppNode.children?.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  return {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) => DynamicPlatforms.includes(inputs.platform),
        data: selectAppTypeQuestion(),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
        children: [
          selectAppManifestNode,
          selectAadAppNode,
          {
            data: inputUserEmailQuestion(
              getLocalizedString("core.getUserEmailQuestion.title"),
              "Email address of the collaborator.",
              true
            ),
          },
        ],
      },
    ],
  };
}

export function listCollaboratorQuestionNode(): IQTreeNode {
  const selectAppManifestNode = selectTeamsAppManifestQuestionNode();
  selectAppManifestNode.condition = {
    contains: CollaborationConstants.TeamsAppQuestionId,
  };
  selectAppManifestNode.children?.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });

  const selectAadAppNode = selectAadAppManifestQuestionNode();
  selectAadAppNode.condition = { contains: CollaborationConstants.AadAppQuestionId };
  selectAadAppNode.children?.push({
    condition: envQuestionCondition,
    data: selectTargetEnvQuestion(QuestionNames.Env, false, false, ""),
  });
  return {
    data: { type: "group" },
    children: [
      {
        condition: (inputs: Inputs) => DynamicPlatforms.includes(inputs.platform),
        data: selectAppTypeQuestion(),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
        children: [selectAppManifestNode, selectAadAppNode],
      },
    ],
  };
}

export async function envQuestionCondition(inputs: Inputs): Promise<boolean> {
  const appType = inputs[CollaborationConstants.AppType] as string[];
  const requireAad = appType?.includes(CollaborationConstants.AadAppQuestionId);
  const requireTeams = appType?.includes(CollaborationConstants.TeamsAppQuestionId);
  const aadManifestPath = inputs[QuestionNames.AadAppManifestFilePath];
  const teamsManifestPath = inputs[QuestionNames.TeamsAppManifestFilePath];

  // When both is selected, only show the question once at the end
  if ((requireAad && !aadManifestPath) || (requireTeams && !teamsManifestPath)) {
    return false;
  }

  // Only show env question when manifest id is referencing value from .env file
  let requireEnv = false;
  if (requireTeams && teamsManifestPath) {
    const teamsAppIdRes = await CollaborationUtil.loadManifestId(teamsManifestPath);
    if (teamsAppIdRes.isOk()) {
      requireEnv = CollaborationUtil.requireEnvQuestion(teamsAppIdRes.value);
      if (requireEnv) {
        return true;
      }
    } else {
      return false;
    }
  }

  if (requireAad && aadManifestPath) {
    const aadAppIdRes = await CollaborationUtil.loadManifestId(aadManifestPath);
    if (aadAppIdRes.isOk()) {
      requireEnv = CollaborationUtil.requireEnvQuestion(aadAppIdRes.value);
      if (requireEnv) {
        return true;
      }
    } else {
      return false;
    }
  }

  return false;
}

export function selectAadAppManifestQuestionNode(): IQTreeNode {
  return {
    data: selectAadManifestQuestion(),
    children: [
      {
        condition: (inputs: Inputs) =>
          inputs.platform === Platform.VSCode && // confirm question only works for VSC
          inputs.projectPath &&
          inputs[QuestionNames.AadAppManifestFilePath] &&
          path.resolve(inputs[QuestionNames.AadAppManifestFilePath]) !==
            path.join(inputs.projectPath, "aad.manifest.json"),
        data: confirmManifestQuestion(false, false),
        cliOptionDisabled: "self",
        inputsDisabled: "self",
      },
    ],
  };
}

function selectAppTypeQuestion(): MultiSelectQuestion {
  const options: MultiSelectQuestion = {
    name: QuestionNames.collaborationAppType,
    title: getLocalizedString("core.selectCollaborationAppTypeQuestion.title"),
    type: "multiSelect",
    staticOptions: [
      {
        id: CollaborationConstants.AadAppQuestionId,
        label: getLocalizedString("core.aadAppQuestion.label"),
        description: getLocalizedString("core.aadAppQuestion.description"),
      },
      {
        id: CollaborationConstants.TeamsAppQuestionId,
        label: getLocalizedString("core.teamsAppQuestion.label"),
        description: getLocalizedString("core.teamsAppQuestion.description"),
      },
    ],
    validation: { minItems: 1 },
    validationHelp: "Please select at least one app type.",
  };
  if (featureFlagManager.getBooleanValue(FeatureFlags.ShareEnabled)) {
    (options.staticOptions as OptionItem[]).push({
      id: CollaborationConstants.AgentOptionId,
      label: getLocalizedString("core.collaboration.agent.label"),
      description: getLocalizedString("core.collaboration.agent.description"),
    });
  }
  return options;
}
