// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateName";
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
  {
    id: "basic-tab-csharp",
    name: TemplateNames.Tab,
    language: "csharp",
    description: "Simple Teams Tab App",
  },
  {
    id: "default-bot-csharp",
    name: TemplateNames.DefaultBot,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-http-trigger-csharp",
    name: TemplateNames.NotificationHttpTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-timer-trigger-csharp",
    name: TemplateNames.NotificationTimerTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-csharp",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-webapi-csharp",
    name: TemplateNames.NotificationWebApi,
    language: "csharp",
    description: "",
  },
  {
    id: "command-and-response-csharp",
    name: TemplateNames.CommandAndResponse,
    language: "csharp",
    description: "",
  },
  {
    id: "workflow-csharp",
    name: TemplateNames.Workflow,
    language: "csharp",
    description: "",
  },
  {
    id: "message-extension-v2-csharp",
    name: TemplateNames.DefaultMessageExtension,
    language: "csharp",
    description: "",
  },
];
