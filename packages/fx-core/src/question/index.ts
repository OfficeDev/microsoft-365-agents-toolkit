// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import { createSampleProjectQuestionNode } from "./create";
import {
  addAuthActionQuestion,
  addKnowledgeQuestionNode,
  addPluginQuestionNode,
  addWebPartQuestionNode,
  apiSpecApiKeyQuestion,
  convertAadToNewSchemaQuestionNode,
  copilotPluginAddAPIQuestionNode,
  createNewEnvQuestionNode,
  deployAadManifestQuestionNode,
  grantPermissionQuestionNode,
  kiotaRegenerateQuestion,
  listCollaboratorQuestionNode,
  oauthQuestion,
  previewWithTeamsAppManifestQuestionNode,
  removeSharedAccessNode,
  selectTeamsAppManifestQuestionNode,
  setSensitivityLabelNode,
  shareNode,
  syncManifestQuestionNode,
  uninstallQuestionNode,
  validateTeamsAppQuestionNode,
  regeneratePluginNode,
} from "./other";
import { scaffoldQuestionForVS } from "./scaffold/vs/createRootNode";
import { createFromTdpNode } from "./scaffold/vsc/createFromTdpNode";
import { scaffoldQuestionForVSCode } from "./scaffold/vsc/createRootNode";
export * from "./constants";
export * from "./create";
export * from "./inputs";
export * from "./options";

export class QuestionNodes {
  createProject(platform: Platform): IQTreeNode {
    if (platform === Platform.VS) return scaffoldQuestionForVS();
    return scaffoldQuestionForVSCode(platform);
  }
  createFromTdp(platform: Platform): IQTreeNode {
    return createFromTdpNode(platform);
  }
  createSampleProject(): IQTreeNode {
    return createSampleProjectQuestionNode();
  }
  addWebpart(): IQTreeNode {
    return addWebPartQuestionNode();
  }
  selectTeamsAppManifest(): IQTreeNode {
    return selectTeamsAppManifestQuestionNode();
  }
  validateTeamsApp(): IQTreeNode {
    return validateTeamsAppQuestionNode();
  }
  previewWithTeamsAppManifest(): IQTreeNode {
    return previewWithTeamsAppManifestQuestionNode();
  }
  listCollaborator(): IQTreeNode {
    return listCollaboratorQuestionNode();
  }
  grantPermission(): IQTreeNode {
    return grantPermissionQuestionNode();
  }
  deployAadManifest(): IQTreeNode {
    return deployAadManifestQuestionNode();
  }
  convertAadToNewSchema(): IQTreeNode {
    return convertAadToNewSchemaQuestionNode();
  }
  createNewEnv(): IQTreeNode {
    return createNewEnvQuestionNode();
  }
  copilotPluginAddAPI(): IQTreeNode {
    return copilotPluginAddAPIQuestionNode();
  }
  apiKey(): IQTreeNode {
    return apiSpecApiKeyQuestion();
  }
  oauth(): IQTreeNode {
    return oauthQuestion();
  }
  addPlugin(): IQTreeNode {
    return addPluginQuestionNode();
  }
  regeneratePlugin(): IQTreeNode {
    return regeneratePluginNode();
  }
  uninstall(): IQTreeNode {
    return uninstallQuestionNode();
  }
  syncManifest(): IQTreeNode {
    return syncManifestQuestionNode();
  }
  kiotaRegenerate(): IQTreeNode {
    return kiotaRegenerateQuestion();
  }
  addAuthAction(): IQTreeNode {
    return addAuthActionQuestion();
  }
  addKnowledge(): IQTreeNode {
    return addKnowledgeQuestionNode();
  }
  setSensitivityLabel(): IQTreeNode {
    return setSensitivityLabelNode();
  }
  share(): IQTreeNode {
    return shareNode();
  }
  removeSharedAccess(): IQTreeNode {
    return removeSharedAccessNode();
  }
}

export const questionNodes = new QuestionNodes();
