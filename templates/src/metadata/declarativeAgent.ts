// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const declarativeAgentTemplates: Template[] = [
  {
    id: "declarative-agent-basic",
    name: TemplateNames.DeclarativeAgentBasic,
    language: "common",
    description: "Basic Declarative Agent",
  },
  {
    id: "declarative-agent-basic-csharp",
    name: TemplateNames.DeclarativeAgentBasic,
    language: "csharp",
    description: "Basic Declarative Agent",
  },
  {
    id: "declarative-agent-basic",
    name: TemplateNames.DeclarativeAgentWithExistingAction,
    language: "common",
    description: "Declarative Agent With Existing Action",
  },
  {
    id: "declarative-agent-basic-csharp",
    name: TemplateNames.DeclarativeAgentWithExistingAction,
    language: "csharp",
    description: "Declarative Agent With Existing Action",
  },
  {
    id: "declarative-agent-with-action-from-scratch-ts",
    name: TemplateNames.DeclarativeAgentWithActionFromScratch,
    language: "typescript",
    description: "Declarative Agent with a new action built from scratch",
  },
  {
    id: "declarative-agent-with-action-from-scratch-js",
    name: TemplateNames.DeclarativeAgentWithActionFromScratch,
    language: "javascript",
    description: "Declarative Agent with a new action built from scratch",
  },
  {
    id: "declarative-agent-with-action-from-scratch-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromScratch,
    language: "csharp",
    description: "Declarative Agent with a new action built from scratch",
  },
  {
    id: "declarative-agent-with-action-from-scratch-bearer-ts",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
    language: "typescript",
    description:
      "Declarative Agent with a new action built from scratch using Bearer Token authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-bearer-js",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
    language: "javascript",
    description:
      "Declarative Agent with a new action built from scratch using Bearer Token authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-bearer-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
    language: "csharp",
    description:
      "Declarative Agent with a new action built from scratch using Bearer Token authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-oauth-ts",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
    language: "typescript",
    description:
      "Declarative Agent with a new action built from scratch using OAuth authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-oauth-js",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
    language: "javascript",
    description:
      "Declarative Agent with a new action built from scratch using OAuth authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-oauth-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
    language: "csharp",
    description:
      "Declarative Agent with a new action built from scratch using OAuth authentication",
  },
  {
    id: "declarative-agent-typespec",
    name: TemplateNames.DeclarativeAgentWithTypeSpec,
    language: "common",
    description: "Declarative Agent with TypeSpec for API definition",
  },
  {
    id: "declarative-agent-with-graph-connector-ts",
    name: TemplateNames.DeclarativeAgentWithGraphConnector,
    language: "typescript",
    description: "Declarative Agent with Microsoft Graph Connector integration",
  },
  {
    id: "declarative-agent-meta-os-new-project",
    name: TemplateNames.DeclarativeAgentMetaOSNewProject,
    language: "common",
    description: "Declarative Agent for MetaOS - new project",
  },
  {
    id: "declarative-agent-meta-os-upgrade-project",
    name: TemplateNames.DeclarativeAgentMetaOSUpgradeProject,
    language: "common",
    description: "Declarative Agent for MetaOS - upgrade existing project",
  },
  {
    id: "declarative-agent-with-action-from-mcp",
    name: TemplateNames.DeclarativeAgentWithActionFromMCP,
    language: "common",
    description: "Declarative Agent with action from Model Context Protocol (MCP)",
  },
];
