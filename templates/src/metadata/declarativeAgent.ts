// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const declarativeAgentTemplates: Template[] = [
  {
    id: "declarative-agent-basic",
    name: TemplateNames.DeclarativeAgentBasic,
    language: "common",
    displayName: "Basic Declarative Agent",
    description: "Basic Declarative Agent",
  },
  {
    id: "declarative-agent-basic-csharp",
    name: TemplateNames.DeclarativeAgentBasic,
    language: "csharp",
    displayName: "Basic Declarative Agent",
    description: "Basic Declarative Agent",
  },
  {
    id: "declarative-agent-basic",
    name: TemplateNames.DeclarativeAgentWithExistingAction,
    language: "common",
    displayName: "Declarative Agent With Existing Action",
    description: "Declarative Agent With Existing Action",
  },
  {
    id: "declarative-agent-basic-csharp",
    name: TemplateNames.DeclarativeAgentWithExistingAction,
    language: "csharp",
    displayName: "Declarative Agent With Existing Action",
    description: "Declarative Agent With Existing Action",
  },
  {
    id: "declarative-agent-with-action-from-scratch-ts",
    name: TemplateNames.DeclarativeAgentWithActionFromScratch,
    language: "typescript",
    displayName: "Declarative Agent with Action from Scratch",
    description: "Declarative Agent with a new action built from scratch",
  },
  {
    id: "declarative-agent-with-action-from-scratch-js",
    name: TemplateNames.DeclarativeAgentWithActionFromScratch,
    language: "javascript",
    displayName: "Declarative Agent with Action from Scratch",
    description: "Declarative Agent with a new action built from scratch",
  },
  {
    id: "declarative-agent-with-action-from-scratch-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromScratch,
    language: "csharp",
    displayName: "Declarative Agent with Action from Scratch",
    description: "Declarative Agent with a new action built from scratch",
  },
  {
    id: "declarative-agent-with-action-from-scratch-bearer-ts",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
    language: "typescript",
    displayName: "Declarative Agent with Action from Scratch (Bearer Token)",
    description:
      "Declarative Agent with a new action built from scratch using Bearer Token authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-bearer-js",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
    language: "javascript",
    displayName: "Declarative Agent with Action from Scratch (Bearer Token)",
    description:
      "Declarative Agent with a new action built from scratch using Bearer Token authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-bearer-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchBearer,
    language: "csharp",
    displayName: "Declarative Agent with Action from Scratch (Bearer Token)",
    description:
      "Declarative Agent with a new action built from scratch using Bearer Token authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-oauth-ts",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
    language: "typescript",
    displayName: "Declarative Agent with Action from Scratch (OAuth)",
    description:
      "Declarative Agent with a new action built from scratch using OAuth authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-oauth-js",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
    language: "javascript",
    displayName: "Declarative Agent with Action from Scratch (OAuth)",
    description:
      "Declarative Agent with a new action built from scratch using OAuth authentication",
  },
  {
    id: "declarative-agent-with-action-from-scratch-oauth-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromScratchOAuth,
    language: "csharp",
    displayName: "Declarative Agent with Action from Scratch (OAuth)",
    description:
      "Declarative Agent with a new action built from scratch using OAuth authentication",
  },
  {
    id: "declarative-agent-typespec",
    name: TemplateNames.DeclarativeAgentWithTypeSpec,
    language: "common",
    displayName: "Declarative Agent with TypeSpec",
    description: "Declarative Agent with TypeSpec for API definition",
  },
  {
    id: "declarative-agent-with-graph-connector-ts",
    name: TemplateNames.DeclarativeAgentWithGraphConnector,
    language: "typescript",
    displayName: "Declarative Agent with Graph Connector",
    description: "Declarative Agent with Microsoft Graph Connector integration",
  },
  {
    id: "declarative-agent-meta-os-new-project",
    name: TemplateNames.DeclarativeAgentMetaOSNewProject,
    language: "common",
    displayName: "Declarative Agent for MetaOS (New Project)",
    description: "Declarative Agent for MetaOS - new project",
  },
  {
    id: "declarative-agent-meta-os-upgrade-project",
    name: TemplateNames.DeclarativeAgentMetaOSUpgradeProject,
    language: "common",
    displayName: "Declarative Agent for MetaOS (Upgrade Project)",
    description: "Declarative Agent for MetaOS - upgrade existing project",
  },
  {
    id: "declarative-agent-with-action-from-mcp",
    name: TemplateNames.DeclarativeAgentWithActionFromMCP,
    language: "common",
    displayName: "Declarative Agent with Action from MCP",
    description: "Declarative Agent with action from Model Context Protocol (MCP)",
  },
];
