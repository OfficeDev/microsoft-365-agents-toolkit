// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { createProjectCliHelpNode, createSampleProjectQuestionNode } from "./create";
import {
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
import { scaffoldQuestionForVSCode } from "./refactor/vsc/createRootNode";
export * from "./constants";
export * from "./create";
export * from "./inputs";
export * from "./options";

export class QuestionNodes {
  createProject(): IQTreeNode {
    // return createProjectQuestionNode();
    return scaffoldQuestionForVSCode();
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
}

export const questionNodes = new QuestionNodes();
