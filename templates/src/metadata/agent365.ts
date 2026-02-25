// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { getString } from "../ui/helper";
import { Template } from "./interface";

export const agent365Templates: Template[] = [
  {
    id: "agent365-ts",
    name: TemplateNames.Agent365,
    language: "typescript",
    displayName: getString("template.customEngineAgent.basic.label"),
    description: getString("template.customEngineAgent.basic.detail"),
  },
];
