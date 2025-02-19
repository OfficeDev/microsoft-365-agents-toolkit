// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { declarativeAgentTemplates } from "./da";
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
    id: "api-plugin-existing-api-csharp",
    name: TemplateNames.ApiPluginWithExistingApiSpec,
    language: "csharp",
    description: "",
  },
  {
    id: "api-plugin-existing-api",
    name: TemplateNames.ApiPluginWithExistingApiSpec,
    language: "none",
    description: "",
  },
  {
    id: "custom-copilot-rag-custom-api-ts",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-rag-custom-api-js",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-rag-custom-api-csharp",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "csharp",
    description: "",
  },
  {
    id: "custom-copilot-rag-custom-api-python",
    name: TemplateNames.CustomCopilotRagCustomApi,
    language: "python",
    description: "",
  },
  {
    id: "copilot-plugin-existing-api",
    name: TemplateNames.MessageExtensionWithExistingApiSpec,
    language: "none",
    description: "",
  },
  {
    id: "copilot-plugin-existing-api-csharp",
    name: TemplateNames.MessageExtensionWithExistingApiSpec,
    language: "csharp",
    description: "",
  },
  ...declarativeAgentTemplates,
];
