// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

// these template are only used in visual studio
export const vsOnlyTemplates: Template[] = [
  {
    id: "empty-csharp",
    name: TemplateNames.Empty,
    language: "csharp",
    description: "",
  },
  {
    id: "message-extension-search-csharp",
    name: TemplateNames.MessageExtensionSearch,
    language: "csharp",
    description: "",
  },
];
