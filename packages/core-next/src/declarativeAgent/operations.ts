// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { defineOperation } from "../core/operation";
import type { AtkContext } from "../core/context";

// Knowledge
import { addKnowledge } from "./knowledge/addKnowledge";
import type { AddKnowledgeInput } from "./knowledge/addKnowledge";

// Actions
import { addExistingPlugin } from "./actions/addAction";
import { addMCPAction } from "./actions/addActionFromMCP";
import { removeAction } from "./actions/removeAction";

// Auth
import { injectOAuthAction, injectApiKeyAction } from "./auth/authInjector";

// Capabilities
import { setSensitivityLabel, setConversationStarters } from "./capabilities/sensitivityLabel";

import type { AddExistingPluginInput, AddMCPActionInput, SetSensitivityLabelInput } from "./types";

// ---------------------------------------------------------------------------
// Zod schemas
// ---------------------------------------------------------------------------

const knowledgeSourceSchema = z.enum([
  "web-search",
  "onedrive-sharepoint",
  "graph-connector",
  "embedded-knowledge",
]);

const addKnowledgeInputSchema = z.object({
  agentManifestPath: z.string().min(1),
  source: knowledgeSourceSchema,
  siteUrl: z.string().optional(),
  odspItem: z
    .object({
      id: z.string(),
      name: z.string(),
      uniqueId: z.string().optional(),
      listId: z.string().optional(),
      webId: z.string().optional(),
      siteId: z.string().optional(),
      webUrl: z.string().optional(),
      itemType: z.enum(["File", "Folder"]).optional(),
    })
    .optional(),
  connectionIds: z.array(z.string()).optional(),
  embeddedFilePaths: z.array(z.string()).optional(),
});

const addExistingPluginInputSchema = z.object({
  agentManifestPath: z.string().min(1),
  pluginManifestPath: z.string().min(1),
  apiSpecPath: z.string().min(1),
  actionId: z.string().min(1),
});

const mcpToolDefinitionSchema = z.object({
  name: z.string(),
  description: z.string().optional(),
  inputSchema: z.record(z.unknown()).optional(),
});

const addMCPActionInputSchema = z.object({
  pluginManifestPath: z.string().min(1),
  serverUrlOrIdentifier: z.string().min(1),
  serverName: z.string().min(1),
  isLocal: z.boolean(),
  auth: z.enum(["oauth", "api-key", "bearer-token", "microsoft-entra", "none"]),
  oauthMeta: z
    .object({
      authorizationUrl: z.string(),
      tokenUrl: z.string(),
      refreshUrl: z.string().optional(),
    })
    .optional(),
  availableTools: z.array(mcpToolDefinitionSchema),
  selectedTools: z.array(z.string()),
});

const removeActionInputSchema = z.object({
  agentManifestPath: z.string().min(1),
  actionId: z.string().min(1),
});

const setSensitivityLabelInputSchema = z.object({
  agentManifestPath: z.string().min(1),
  labelId: z.string().min(1),
});

const setConversationStartersInputSchema = z.object({
  agentManifestPath: z.string().min(1),
  starters: z.array(
    z.object({
      text: z.string(),
      title: z.string().optional(),
    })
  ),
});

const injectOAuthInputSchema = z.object({
  ymlPath: z.string().min(1),
  authName: z.string().min(1),
  specRelativePath: z.string().min(1),
  isMicrosoftEntra: z.boolean(),
  enablePKCE: z.boolean().optional(),
  registrationId: z.string().optional(),
});

const injectApiKeyInputSchema = z.object({
  ymlPath: z.string().min(1),
  authName: z.string().min(1),
  specRelativePath: z.string().min(1),
  registrationId: z.string().optional(),
});

// ---------------------------------------------------------------------------
// Operation definitions
// ---------------------------------------------------------------------------

/**
 * Add knowledge to a declarative agent.
 */
export const addKnowledgeOp = defineOperation(
  "da/addKnowledge",
  addKnowledgeInputSchema,
  async (_ctx: AtkContext, input: AddKnowledgeInput) => {
    return addKnowledge(input);
  }
);

/**
 * Add an existing plugin to a declarative agent.
 */
export const addExistingPluginOp = defineOperation(
  "da/addExistingPlugin",
  addExistingPluginInputSchema,
  async (_ctx: AtkContext, input: AddExistingPluginInput) => {
    return addExistingPlugin(input);
  }
);

/**
 * Add an MCP-backed action to a plugin manifest.
 */
export const addMCPActionOp = defineOperation(
  "da/addMCPAction",
  addMCPActionInputSchema,
  async (_ctx: AtkContext, input: AddMCPActionInput) => {
    return addMCPAction(input);
  }
);

/**
 * Remove an action from a declarative agent manifest.
 */
export const removeActionOp = defineOperation(
  "da/removeAction",
  removeActionInputSchema,
  async (_ctx: AtkContext, input: z.infer<typeof removeActionInputSchema>) => {
    return removeAction(input.agentManifestPath, input.actionId);
  }
);

/**
 * Set a sensitivity label on a declarative agent manifest.
 */
export const setSensitivityLabelOp = defineOperation(
  "da/setSensitivityLabel",
  setSensitivityLabelInputSchema,
  async (_ctx: AtkContext, input: SetSensitivityLabelInput) => {
    return setSensitivityLabel(input);
  }
);

/**
 * Add conversation starters to a declarative agent manifest.
 */
export const setConversationStartersOp = defineOperation(
  "da/setConversationStarters",
  setConversationStartersInputSchema,
  async (_ctx: AtkContext, input: z.infer<typeof setConversationStartersInputSchema>) => {
    return setConversationStarters(input.agentManifestPath, input.starters);
  }
);

/**
 * Inject an OAuth auth action into teamsapp.yml.
 */
export const injectOAuthActionOp = defineOperation(
  "da/injectOAuthAction",
  injectOAuthInputSchema,
  async (_ctx: AtkContext, input: z.infer<typeof injectOAuthInputSchema>) => {
    return injectOAuthAction(
      input.ymlPath,
      input.authName,
      input.specRelativePath,
      input.isMicrosoftEntra,
      input.enablePKCE,
      input.registrationId
    );
  }
);

/**
 * Inject an API-key auth action into teamsapp.yml.
 */
export const injectApiKeyActionOp = defineOperation(
  "da/injectApiKeyAction",
  injectApiKeyInputSchema,
  async (_ctx: AtkContext, input: z.infer<typeof injectApiKeyInputSchema>) => {
    return injectApiKeyAction(
      input.ymlPath,
      input.authName,
      input.specRelativePath,
      input.registrationId
    );
  }
);
