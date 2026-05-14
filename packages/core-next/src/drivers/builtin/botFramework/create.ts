// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import { appStudioScopes } from "../../../clients/teamsDevPortal/types";
import { BotRegistration, BotChannelType } from "../../../clients/graphApi/types";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const botChannelSchema = z.object({
  name: z.nativeEnum(BotChannelType),
  callingWebhook: z.string().optional(),
});

const inputSchema = z.object({
  /** Bot registration ID — must be a valid UUID (typically the Bot AAD app's client ID) */
  botId: z.string().min(1),
  /** Bot display name */
  name: z.string().min(1),
  /** Webhook URL for bot messaging */
  messagingEndpoint: z.string().min(1),
  /** Optional bot description */
  description: z.string().optional().default(""),
  /** Optional bot icon URL */
  iconUrl: z.string().optional().default(""),
  /** Channel configurations */
  channels: z.array(botChannelSchema).optional(),
});

/**
 * Driver: botFramework/create
 *
 * Creates or updates a Bot Framework registration via the Teams Developer Portal API.
 * Idempotent: if the bot already exists, merges and updates.
 *
 * Outputs: (none — bot registration has no env var outputs)
 */
export const createBotFrameworkDriver = createDriver({
  id: "botFramework/create",
  name: "Create Bot Framework Registration",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "botFramework/create";

    // Validate UUID format
    if (!UUID_RE.test(config.botId)) {
      return err(
        userError("InvalidBotId", `botId must be a valid UUID, got: ${config.botId}`, { source })
      );
    }

    // Acquire M365 token (TDP uses AppStudio scopes, not Graph)
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: appStudioScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source,
          inner: tokenRes.error,
        })
      );
    }
    const client = new TeamsDevPortalClient(ctx, tokenRes.value);

    // Build registration
    const configuredChannels = config.channels?.map((c) => c.name) ?? [
      BotChannelType.MicrosoftTeams,
    ];

    const registration: BotRegistration = {
      botId: config.botId,
      name: config.name,
      description: config.description ?? "",
      iconUrl: config.iconUrl ?? "",
      messagingEndpoint: config.messagingEndpoint,
      callingEndpoint: "",
      configuredChannels,
      isSingleTenant: true,
    };

    // Check existing registration
    const existingRes = await client.getBotRegistration(config.botId);
    if (existingRes.isOk() && existingRes.value) {
      // Merge: local config takes precedence over remote
      const merged: BotRegistration = {
        ...existingRes.value,
        ...registration,
      };
      const updateRes = await client.updateBotRegistration(merged);
      if (updateRes.isErr()) return err(updateRes.error);
      ctx.logger.info(`[${source}] Updated bot registration: ${config.botId}`);
    } else {
      // Create new
      const createRes = await client.createBotRegistration(registration);
      if (createRes.isErr()) return err(createRes.error);
      ctx.logger.info(`[${source}] Created bot registration: ${config.botId}`);
    }

    return ok({ outputs: {} });
  },
});
