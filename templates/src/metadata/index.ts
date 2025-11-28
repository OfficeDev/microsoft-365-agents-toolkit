// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { customEngineAgentTemplates } from "./customEngineAgent";
import { declarativeAgentTemplates } from "./declarativeAgent";
import { graphConnectorTemplates } from "./graphConnector";
import { Template } from "./interface";
import { specialTemplates } from "./special";
import { teamsAgentsAndAppsTemplates } from "./teams";
import { vsOnlyTemplates } from "./vs";
import { wxpTemplates } from "./wxp";

export const allTemplates: Template[] = [
  ...declarativeAgentTemplates,
  ...customEngineAgentTemplates,
  ...graphConnectorTemplates,
  ...teamsAgentsAndAppsTemplates,
  ...wxpTemplates,
  // ...messagingExtensionTemplates,
  ...specialTemplates,
  ...vsOnlyTemplates,
];

export const defaultGeneratorTemplates: Template[] = [
  ...customEngineAgentTemplates,
  ...teamsAgentsAndAppsTemplates,
  // ...messagingExtensionTemplates,
  ...vsOnlyTemplates,
  ...graphConnectorTemplates,
];
