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
    displayName: "Empty Project",
    description: "Start with a minimal Teams app project",
  },
  {
    id: "message-extension-search-csharp",
    name: TemplateNames.MessageExtensionSearch,
    language: "csharp",
    displayName: "Message Extension (Search)",
    description: "Search-based message extension for Teams",
  },
  {
    id: "basic-tab-csharp",
    name: TemplateNames.Tab,
    language: "csharp",
    displayName: "Basic Tab",
    description: "Simple Teams Tab App",
  },
  {
    id: "default-bot-csharp",
    name: TemplateNames.DefaultBot,
    language: "csharp",
    displayName: "Echo Bot",
    description: "A simple implementation of an echo bot that's ready for customization",
  },
  {
    id: "notification-http-trigger-csharp",
    name: TemplateNames.NotificationHttpTrigger,
    language: "csharp",
    displayName: "Notification Bot (HTTP Trigger)",
    description: "Send notifications to Teams using HTTP-triggered Azure Function",
  },
  {
    id: "notification-timer-trigger-csharp",
    name: TemplateNames.NotificationTimerTrigger,
    language: "csharp",
    displayName: "Notification Bot (Timer Trigger)",
    description: "Send scheduled notifications to Teams using timer-triggered Azure Function",
  },
  {
    id: "notification-http-timer-trigger-csharp",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "csharp",
    displayName: "Notification Bot (HTTP & Timer Trigger)",
    description: "Send notifications to Teams using both HTTP and timer triggers",
  },
  {
    id: "notification-webapi-csharp",
    name: TemplateNames.NotificationWebApi,
    language: "csharp",
    displayName: "Notification Bot (Web API)",
    description: "Send notifications to Teams using ASP.NET Core Web API",
  },
  {
    id: "command-and-response-csharp",
    name: TemplateNames.CommandAndResponse,
    language: "csharp",
    displayName: "Command and Response Bot",
    description: "Respond to simple commands in Teams chat",
  },
  {
    id: "workflow-csharp",
    name: TemplateNames.Workflow,
    language: "csharp",
    displayName: "Workflow Bot",
    description: "Build a sequential workflow bot for Teams",
  },
  {
    id: "message-extension-v2-csharp",
    name: TemplateNames.DefaultMessageExtension,
    language: "csharp",
    displayName: "Message Extension",
    description: "Receive user input, process it, and send customized results",
  },
];
