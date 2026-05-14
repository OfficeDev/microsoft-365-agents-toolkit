// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export {
  TemplateDescriptor,
  TemplateCategory,
  Language,
  ScaffoldResult,
  ProvisionResult,
  DeployResult,
  TemplateActionOptions,
  QuestionSpec,
} from "./types";

export { TemplateRegistry, templateRegistry } from "./registry";

export * from "./scaffold";

export {
  registerBuiltinTemplates,
  daTemplateDescriptors,
  DATemplateNames,
  botTemplateDescriptors,
  BotTemplateNames,
  tabTemplateDescriptors,
  TabTemplateNames,
  aiAgentTemplateDescriptors,
  AiAgentTemplateNames,
  engineAgentTemplateDescriptors,
  EngineAgentTemplateNames,
  connectorTemplateDescriptors,
  ConnectorTemplateNames,
  messageExtensionTemplateDescriptors,
  MessageExtensionTemplateNames,
  openApiTemplateDescriptors,
  OpenApiTemplateNames,
} from "./descriptors";

export * from "./openApi";
