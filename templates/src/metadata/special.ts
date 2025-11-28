// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { getString } from "../ui/helper";
import { Template } from "./interface";

// these template are not handled by default generator which means they need extra steps during scaffolding
export const specialTemplates: Template[] = [
  {
    id: "non-sso-tab-ssr-cs",
    name: TemplateNames.TabSSR,
    language: "csharp",
    displayName: "Tab (Server-Side Rendering)",
    description: "Simple Teams Tab App",
  },
  {
    id: "sso-tab-ssr-cs",
    name: TemplateNames.SsoTabSSR,
    language: "csharp",
    displayName: "Tab with SSO (Server-Side Rendering)",
    description: "Simple Teams Tab App with SSO",
  },
  {
    id: "custom-copilot-rag-custom-api-ts",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "typescript",
    displayName: "Teams Agent with Data from Custom API using OpenAPI Spec",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-custom-api-js",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "javascript",
    displayName: "Teams Agent with Data from Custom API using OpenAPI Spec",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "teams-agent-with-data-custom-api-v2-csharp",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "csharp",
    displayName: "Teams Agent with Data from Custom API using OpenAPI Spec",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "teams-agent-with-data-custom-api-v2-python",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "python",
    displayName: "Teams Agent with Data from Custom API using OpenAPI Spec",
    description: getString("template.teams.rag.detail"),
  },
  // {
  //   id: "message-extension-with-existing-api",
  //   name: TemplateNames.MessageExtensionWithExistingApiSpec,
  //   language: "common",
  //   displayName: "Message Extension with Existing API",
  //   description: "Message extension built from an existing API specification",
  // },
  // {
  //   id: "message-extension-with-existing-api-csharp",
  //   name: TemplateNames.MessageExtensionWithExistingApiSpec,
  //   language: "csharp",
  //   displayName: "Message Extension with Existing API",
  //   description: "Message extension built from an existing API specification",
  // },
];
