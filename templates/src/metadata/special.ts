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

export const foundryAgentTemplate: Template = {
  id: "foundry-agent-to-m365-ts",
  name: TemplateNames.FoundryAgent,
  language: "typescript",
  displayName: "Foundry Agent",
  description: "An Microsoft 365 Agent that connects to Microsoft AI Foundry Agent.",
};
