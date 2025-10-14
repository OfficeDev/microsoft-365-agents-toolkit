// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

const notificationBotTemplates: Template[] = [
  {
    id: "notification-http-trigger-ts",
    name: TemplateNames.NotificationHttpTrigger,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-http-trigger-js",
    name: TemplateNames.NotificationHttpTrigger,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-http-trigger-csharp",
    name: TemplateNames.NotificationHttpTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-timer-trigger-ts",
    name: TemplateNames.NotificationTimerTrigger,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-timer-trigger-js",
    name: TemplateNames.NotificationTimerTrigger,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-timer-trigger-csharp",
    name: TemplateNames.NotificationTimerTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-ts",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-js",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-csharp",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-express-ts",
    name: TemplateNames.NotificationExpress,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-express-js",
    name: TemplateNames.NotificationExpress,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-webapi-csharp",
    name: TemplateNames.NotificationWebApi,
    language: "csharp",
    description: "",
  },
];

export const basicBotTemplates: Template[] = [
  {
    id: "default-bot-ts",
    name: TemplateNames.DefaultBot,
    language: "typescript",
    description: "",
  },
  {
    id: "default-bot-js",
    name: TemplateNames.DefaultBot,
    language: "javascript",
    description: "",
  },
  {
    id: "default-bot-csharp",
    name: TemplateNames.DefaultBot,
    language: "csharp",
    description: "",
  },
  ...notificationBotTemplates,
  {
    id: "command-and-response-ts",
    name: TemplateNames.CommandAndResponse,
    language: "typescript",
    description: "",
    link: "https://aka.ms/teamsfx-create-command",
  },
  {
    id: "command-and-response-js",
    name: TemplateNames.CommandAndResponse,
    language: "javascript",
    description: "",
    link: "https://aka.ms/teamsfx-create-command",
  },
  {
    id: "command-and-response-csharp",
    name: TemplateNames.CommandAndResponse,
    language: "csharp",
    description: "",
  },
  {
    id: "workflow-ts",
    name: TemplateNames.Workflow,
    language: "typescript",
    description: "",
    link: "https://aka.ms/teamsfx-workflow-new",
  },
  {
    id: "workflow-js",
    name: TemplateNames.Workflow,
    language: "javascript",
    description: "",
    link: "https://aka.ms/teamsfx-workflow-new",
  },
  {
    id: "workflow-csharp",
    name: TemplateNames.Workflow,
    language: "csharp",
    description: "",
  },
];
