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
    description: "Start with a minimal Teams app project",
  },
  {
    id: "message-extension-search-csharp",
    name: TemplateNames.MessageExtensionSearch,
    language: "csharp",
    description: "Search-based message extension for Teams",
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
    description: "A simple implementation of an echo bot that's ready for customization",
  },
  {
    id: "notification-http-trigger-csharp",
    name: TemplateNames.NotificationHttpTrigger,
    language: "csharp",
    description: "Send notifications to Teams using HTTP-triggered Azure Function",
  },
  {
    id: "notification-timer-trigger-csharp",
    name: TemplateNames.NotificationTimerTrigger,
    language: "csharp",
    description: "Send scheduled notifications to Teams using timer-triggered Azure Function",
  },
  {
    id: "notification-http-timer-trigger-csharp",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "csharp",
    description: "Send notifications to Teams using both HTTP and timer triggers",
  },
  {
    id: "notification-webapi-csharp",
    name: TemplateNames.NotificationWebApi,
    language: "csharp",
    description: "Send notifications to Teams using ASP.NET Core Web API",
  },
  {
    id: "command-and-response-csharp",
    name: TemplateNames.CommandAndResponse,
    language: "csharp",
    description: "Respond to simple commands in Teams chat",
  },
  {
    id: "workflow-csharp",
    name: TemplateNames.Workflow,
    language: "csharp",
    description: "Build a sequential workflow bot for Teams",
  },
  {
    id: "message-extension-v2-csharp",
    name: TemplateNames.DefaultMessageExtension,
    language: "csharp",
    description: "Receive user input, process it, and send customized results",
  },
];
