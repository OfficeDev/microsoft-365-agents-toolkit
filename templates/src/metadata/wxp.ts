// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const wxpTemplates: Template[] = [
  {
    id: "office-addin-outlook-taskpane-ts",
    name: TemplateNames.OutlookTaskpane,
    language: "typescript",
    displayName: "Outlook Task Pane Add-in",
    description: "Build a task pane add-in for Outlook",
  },
  {
    id: "office-addin-wxpo-taskpane-ts",
    name: TemplateNames.WXPTaskpane,
    language: "typescript",
    displayName: "Office Task Pane Add-in",
    description: "Build a task pane add-in for Word, Excel, or PowerPoint",
  },
  {
    id: "office-addin-excel-cfshortcut-ts",
    name: TemplateNames.ExcelCFShortcut,
    language: "typescript",
    displayName: "Excel Custom Functions",
    description: "Create custom functions in Excel with keyboard shortcuts",
  },
  {
    id: "office-addin-config-ts",
    name: TemplateNames.OfficeAddinCommon,
    language: "typescript",
    displayName: "Office Add-in Common Configuration",
    description: "Common configuration for Office Add-ins",
  },
];
