// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const commonTemplates: Template[] = [
  {
    id: "api-plugin-existing-api",
    name: TemplateNames.MessageExtensionWithExistingApiSpec,
    language: "common",
    description: "",
  },
  {
    id: "copilot-plugin-existing-api",
    name: TemplateNames.CopilotPluginExistingApi,
    language: "common",
    description: "",
  },
];
