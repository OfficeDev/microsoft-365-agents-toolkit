// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import { createProjectCliHelpNode, createSampleProjectQuestionNode } from "./create";
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
  selectTeamsAppManifestQuestionNode,
  syncManifestQuestionNode,
  uninstallQuestionNode,
  validateTeamsAppQuestionNode,
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
    // return createProjectQuestionNode();
    if (platform === Platform.VS) return scaffoldQuestionForVS();
    return scaffoldQuestionForVSCode();
  }
  createFromTdp(): IQTreeNode {
    return createFromTdpNode();
  }
  createSampleProject(): IQTreeNode {
    return createSampleProjectQuestionNode();
  }
  createProjectCliHelp(): IQTreeNode {
    return createProjectCliHelpNode();
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
}

export const questionNodes = new QuestionNodes();
