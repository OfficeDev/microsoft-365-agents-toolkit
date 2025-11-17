// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const wxpTemplates: Template[] = [
  {
    id: "office-addin-outlook-taskpane-ts",
    name: TemplateNames.OutlookTaskpane,
    language: "typescript",
    description: "Build a task pane add-in for Outlook",
  },
  {
    id: "office-addin-wxpo-taskpane-ts",
    name: TemplateNames.WXPTaskpane,
    language: "typescript",
    description: "Build a task pane add-in for Word, Excel, or PowerPoint",
  },
  {
    id: "office-addin-excel-cfshortcut-ts",
    name: TemplateNames.ExcelCFShortcut,
    language: "typescript",
    description: "Create custom functions in Excel with keyboard shortcuts",
  },
  {
    id: "office-addin-config-ts",
    name: TemplateNames.OfficeAddinCommon,
    language: "typescript",
    description: "Common configuration for Office Add-ins",
  },
];
