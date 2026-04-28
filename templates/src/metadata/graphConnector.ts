// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateAlias, TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const graphConnectorTemplates: Template[] = [
  {
    id: "graph-connector-ts",
    name: TemplateNames.GraphConnector,
    alias: TemplateAlias.GraphConnector,
    language: "typescript",
    displayName: "Copilot Connector",
    description: "Embed your organization data to make it searchable in Microsoft 365 Copilot",
  },
];
