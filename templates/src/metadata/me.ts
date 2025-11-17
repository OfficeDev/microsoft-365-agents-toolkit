// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

export const messagingExtensionTemplates: Template[] = [
  {
    id: "message-extension-with-api-from-scratch-ts",
    name: TemplateNames.MessageExtensionWithNewApiFromScratch,
    language: "typescript",
    description: "Build a message extension with a new API built from scratch",
  },
  {
    id: "message-extension-with-api-from-scratch-js",
    name: TemplateNames.MessageExtensionWithNewApiFromScratch,
    language: "javascript",
    description: "Build a message extension with a new API built from scratch",
  },
  {
    id: "message-extension-with-api-from-scratch-csharp",
    name: TemplateNames.MessageExtensionWithNewApiFromScratch,
    language: "csharp",
    description: "Build a message extension with a new API built from scratch",
  },
  {
    id: "message-extension-with-api-from-scratch-api-key-ts",
    name: TemplateNames.MessageExtensionWithNewApiFromScratchUsingApiKey,
    language: "typescript",
    description: "Message extension with API using API Key authentication",
  },
  {
    id: "message-extension-with-api-from-scratch-api-key-js",
    name: TemplateNames.MessageExtensionWithNewApiFromScratchUsingApiKey,
    language: "javascript",
    description: "Message extension with API using API Key authentication",
  },
  {
    id: "message-extension-with-api-from-scratch-api-key-csharp",
    name: TemplateNames.MessageExtensionWithNewApiFromScratchUsingApiKey,
    language: "csharp",
    description: "Message extension with API using API Key authentication",
  },
  {
    id: "message-extension-with-api-from-scratch-sso-ts",
    name: TemplateNames.MessageExtensionWithNewApiFromScratchUsingOAuth,
    language: "typescript",
    description: "Message extension with API using OAuth authentication",
  },
  {
    id: "message-extension-with-api-from-scratch-sso-js",
    name: TemplateNames.MessageExtensionWithNewApiFromScratchUsingOAuth,
    language: "javascript",
    description: "Message extension with API using OAuth authentication",
  },
  {
    id: "message-extension-with-api-from-scratch-sso-csharp",
    name: TemplateNames.MessageExtensionWithNewApiFromScratchUsingOAuth,
    language: "csharp",
    description: "Message extension with API using OAuth authentication",
  },
  {
    id: "m365-message-extension-ts",
    name: TemplateNames.MessageExtensionM365,
    language: "typescript",
    description: "Message extension that works across Microsoft 365",
  },
  {
    id: "m365-message-extension-js",
    name: TemplateNames.MessageExtensionM365,
    language: "javascript",
    description: "Message extension that works across Microsoft 365",
  },
  {
    id: "message-extension-action-csharp",
    name: TemplateNames.MessageExtensionAction,
    language: "csharp",
    description: "Collect input from users and process it in your app",
  },
  {
    id: "message-extension-action-ts",
    name: TemplateNames.MessageExtensionAction,
    language: "typescript",
    description: "Collect input from users and process it in your app",
  },
  {
    id: "message-extension-action-js",
    name: TemplateNames.MessageExtensionAction,
    language: "javascript",
    description: "Collect input from users and process it in your app",
  },
  {
    id: "link-unfurling-ts",
    name: TemplateNames.LinkUnfurling,
    language: "typescript",
    description: "Display rich previews when URLs are pasted in Teams conversations",
  },
  {
    id: "link-unfurling-js",
    name: TemplateNames.LinkUnfurling,
    language: "javascript",
    description: "Display rich previews when URLs are pasted in Teams conversations",
  },
  {
    id: "link-unfurling-csharp",
    name: TemplateNames.LinkUnfurling,
    language: "csharp",
    description: "Display rich previews when URLs are pasted in Teams conversations",
  },
];
