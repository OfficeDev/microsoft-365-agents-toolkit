// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { templateRegistry } from "../registry";
import { daTemplateDescriptors } from "./declarativeAgent";
import { botTemplateDescriptors } from "./bot";
import { tabTemplateDescriptors } from "./tab";
import { aiAgentTemplateDescriptors } from "./aiAgent";
import { engineAgentTemplateDescriptors } from "./engineAgent";
import { connectorTemplateDescriptors } from "./connector";
import { messageExtensionTemplateDescriptors } from "./messageExtension";
import { openApiTemplateDescriptors } from "./openApi";
import { foundryTemplateDescriptors } from "./foundry";

/**
 * All built-in template descriptor arrays, collected for registration.
 */
const allBuiltinDescriptors = [
  ...daTemplateDescriptors,
  ...botTemplateDescriptors,
  ...tabTemplateDescriptors,
  ...aiAgentTemplateDescriptors,
  ...engineAgentTemplateDescriptors,
  ...connectorTemplateDescriptors,
  ...messageExtensionTemplateDescriptors,
  ...openApiTemplateDescriptors,
  ...foundryTemplateDescriptors,
];

/**
 * Register all built-in template descriptors.
 * Call this once at application startup (CLI init, VS Code activation, etc.).
 */
export function registerBuiltinTemplates(): void {
  for (const descriptor of allBuiltinDescriptors) {
    if (!templateRegistry.has(descriptor.id)) {
      templateRegistry.register(descriptor);
    }
  }
}

export { daTemplateDescriptors, DATemplateNames } from "./declarativeAgent";
export { botTemplateDescriptors, BotTemplateNames } from "./bot";
export { tabTemplateDescriptors, TabTemplateNames } from "./tab";
export { aiAgentTemplateDescriptors, AiAgentTemplateNames } from "./aiAgent";
export { engineAgentTemplateDescriptors, EngineAgentTemplateNames } from "./engineAgent";
export { connectorTemplateDescriptors, ConnectorTemplateNames } from "./connector";
export {
  messageExtensionTemplateDescriptors,
  MessageExtensionTemplateNames,
} from "./messageExtension";
export { openApiTemplateDescriptors, OpenApiTemplateNames } from "./openApi";
export { foundryTemplateDescriptors, FoundryTemplateNames } from "./foundry";
