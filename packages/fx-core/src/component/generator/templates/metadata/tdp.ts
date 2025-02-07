// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const tdpTemplates: Template[] = [
  {
    id: "non-sso-tab-default-bot-ts",
    name: TemplateNames.TabAndDefaultBot,
    language: "typescript",
    description: "",
  },
  {
    id: "non-sso-tab-default-bot-js",
    name: TemplateNames.TabAndDefaultBot,
    language: "javascript",
    description: "",
  },
  {
    id: "default-bot-message-extension-ts",
    name: TemplateNames.BotAndMessageExtension,
    language: "typescript",
    description: "",
  },
  {
    id: "default-bot-message-extension-js",
    name: TemplateNames.BotAndMessageExtension,
    language: "javascript",
    description: "",
  },
  {
    id: "message-extension-ts",
    name: TemplateNames.MessageExtension,
    language: "typescript",
    description: "",
  },
  {
    id: "message-extension-js",
    name: TemplateNames.MessageExtension,
    language: "javascript",
    description: "",
  },
  {
    id: "message-extension-csharp",
    name: TemplateNames.MessageExtension,
    language: "csharp",
    description: "",
  },
];
