// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const tabTemplates: Template[] = [
  {
    id: "non-sso-tab-ts",
    name: TemplateNames.Tab,
    language: "typescript",
    description: "Simple Teams Tab App",
  },
  {
    id: "non-sso-tab-js",
    name: TemplateNames.Tab,
    language: "javascript",
    description: "Simple Teams Tab App",
  },
  {
    id: "sso-tab-with-obo-flow-ts",
    name: TemplateNames.SsoTabObo,
    language: "typescript",
    description: "Simple Teams Tab App with OBO Flow",
  },
  {
    id: "sso-tab-with-obo-flow-js",
    name: TemplateNames.SsoTabObo,
    language: "javascript",
    description: "Simple Teams Tab App with OBO Flow",
  },
  {
    id: "dashboard-tab-ts",
    name: TemplateNames.DashboardTab,
    language: "typescript",
    description: "Dashboard Tab App",
    link: "https://aka.ms/teamsfx-dashboard-app",
  },
  {
    id: "dashboard-tab-js",
    name: TemplateNames.DashboardTab,
    language: "javascript",
    description: "Dashboard Tab App",
    link: "https://aka.ms/teamsfx-dashboard-app",
  },
];
