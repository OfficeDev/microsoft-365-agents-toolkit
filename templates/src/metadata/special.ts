// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

// these template are not handled by default generator which means they need extra steps during scaffolding
export const specialTemplates: Template[] = [
  {
    id: "non-sso-tab-ssr-cs",
    name: TemplateNames.TabSSR,
    language: "csharp",
    description: "Simple Teams Tab App",
  },
  {
    id: "sso-tab-ssr-cs",
    name: TemplateNames.SsoTabSSR,
    language: "csharp",
    description: "Simple Teams Tab App with SSO",
  },
  {
    id: "declarative-agent-with-action-from-existing-api",
    name: TemplateNames.DeclarativeAgentWithActionFromExistingApiSpec,
    language: "none",
    description: "Declarative Agent with action from an existing API specification",
  },
  {
    id: "declarative-agent-with-action-from-existing-api-csharp",
    name: TemplateNames.DeclarativeAgentWithActionFromExistingApiSpec,
    language: "csharp",
    description: "Declarative Agent with action from an existing API specification",
  },
  {
    id: "custom-copilot-rag-custom-api-ts",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "typescript",
    description: "AI agent that chats with your data from a custom API",
  },
  {
    id: "custom-copilot-rag-custom-api-js",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "javascript",
    description: "AI agent that chats with your data from a custom API",
  },
  {
    id: "teams-agent-with-data-custom-api-v2-csharp",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "csharp",
    description: "AI agent that chats with your data from a custom API",
  },
  {
    id: "teams-agent-with-data-custom-api-v2-python",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "python",
    description: "AI agent that chats with your data from a custom API",
  },
  {
    id: "message-extension-with-existing-api",
    name: TemplateNames.MessageExtensionWithExistingApiSpec,
    language: "common",
    description: "Message extension built from an existing API specification",
  },
  {
    id: "message-extension-with-existing-api-csharp",
    name: TemplateNames.MessageExtensionWithExistingApiSpec,
    language: "csharp",
    description: "Message extension built from an existing API specification",
  },
];
