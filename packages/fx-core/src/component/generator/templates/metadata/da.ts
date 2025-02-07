// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const declarativeAgentTemplates: Template[] = [
  {
    id: "copilot-gpt-basic",
    name: TemplateNames.BasicGpt,
    language: "common",
    description: "",
  },
  {
    id: "copilot-gpt-basic-csharp",
    name: TemplateNames.BasicGpt,
    language: "csharp",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-ts",
    name: TemplateNames.ApiPluginFromScratch,
    language: "typescript",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-js",
    name: TemplateNames.ApiPluginFromScratch,
    language: "javascript",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-csharp",
    name: TemplateNames.ApiPluginFromScratch,
    language: "csharp",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-bearer-ts",
    name: TemplateNames.ApiPluginFromScratchBearer,
    language: "typescript",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-bearer-js",
    name: TemplateNames.ApiPluginFromScratchBearer,
    language: "javascript",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-bearer-csharp",
    name: TemplateNames.ApiPluginFromScratchBearer,
    language: "csharp",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-oauth-ts",
    name: TemplateNames.ApiPluginFromScratchOAuth,
    language: "typescript",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-oauth-js",
    name: TemplateNames.ApiPluginFromScratchOAuth,
    language: "javascript",
    description: "",
  },
  {
    id: "api-plugin-from-scratch-oauth-csharp",
    name: TemplateNames.ApiPluginFromScratchOAuth,
    language: "csharp",
    description: "",
  },
];
