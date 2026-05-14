// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Canonical question names used across the question model.
 * Maps to the `name` property of Question objects.
 *
 * OLD fx-core names → NEW core-next names:
 *  AppName → projectName
 *  Folder → destinationFolder
 *  ProgrammingLanguage → language
 *  ProjectType → projectType
 *  Capabilities → templateId
 *  ApiSpecLocation → apiSpecPath
 *  ApiOperation → apiOperations
 *  LLMService → llmProvider
 *  CustomCopilotRag → ragSource
 */
export const QuestionNames = {
  projectName: "projectName",
  destinationFolder: "destinationFolder",
  language: "language",
  projectType: "projectType",
  templateId: "templateId",
  apiSpecPath: "apiSpecPath",
  apiOperations: "apiOperations",
  llmProvider: "llmProvider",
  azureOpenAiKey: "azureOpenAiKey",
  azureOpenAiEndpoint: "azureOpenAiEndpoint",
  openAiKey: "openAiKey",
  oauthClientId: "oauthClientId",
  oauthClientSecret: "oauthClientSecret",
  ragSource: "ragSource",
  mcpType: "mcpType",
  mcpServerUrl: "mcpServerUrl",
  graphConnectorName: "graphConnectorName",
  graphConnectorConnectionId: "graphConnectorConnectionId",
  knowledgeSource: "knowledgeSource",
  foundryEndpoint: "foundryEndpoint",
  foundryAgentId: "foundryAgentId",
  officeAddinFolder: "officeAddinFolder",
} as const;

export type QuestionName = (typeof QuestionNames)[keyof typeof QuestionNames];
