// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { GraphApiClient } from "../../../clients/graphApi/client";
import { graphScopes, SignInAudience } from "../../../clients/graphApi/types";

const MAX_NAME_LENGTH = 120;

const inputSchema = z.object({
  /** Display name for the bot's Entra app (max 120 chars) */
  name: z.string().min(1).max(MAX_NAME_LENGTH),
  /**
   * Existing bot ID for idempotency.
   * When both botId and botPassword are present, creation is skipped.
   */
  existingBotId: z.string().optional(),
  /** Existing bot password for idempotency */
  existingBotPassword: z.string().optional(),
});

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Driver: botAadApp/create
 *
 * Creates an Entra ID application specifically for a Bot Framework bot.
 * Always uses AzureADMultipleOrgs audience and always generates a client secret.
 *
 * Outputs:
 * - BOT_ID (the application/client ID)
 * - SECRET_BOT_PASSWORD (the client secret)
 */
export const createBotAadAppDriver = createDriver({
  id: "botAadApp/create",
  name: "Create Bot AAD App",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "botAadApp/create";

    // Idempotency: reuse if both ID and password exist
    if (config.existingBotId && UUID_RE.test(config.existingBotId) && config.existingBotPassword) {
      ctx.logger.info(`[${source}] Bot AAD app ${config.existingBotId} already exists — reusing`);
      return ok({
        outputs: {
          BOT_ID: config.existingBotId,
          SECRET_BOT_PASSWORD: config.existingBotPassword,
        },
      });
    }

    // Detect corrupt state: ID present but password missing
    if (config.existingBotId && UUID_RE.test(config.existingBotId) && !config.existingBotPassword) {
      return err(
        userError(
          "UnexpectedEmptyBotPassword",
          "BOT_ID exists but SECRET_BOT_PASSWORD is empty. Please clean up the environment and retry.",
          { source }
        )
      );
    }

    // Acquire Graph token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: graphScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source,
          inner: tokenRes.error,
        })
      );
    }
    const client = new GraphApiClient(ctx, tokenRes.value);

    // Create bot AAD app — always multi-org for cross-tenant bot auth
    const createRes = await client.createAadApp(config.name, SignInAudience.AzureADMultipleOrgs);
    if (createRes.isErr()) return err(createRes.error);

    const objectId = createRes.value.id ?? "";
    const botId = createRes.value.appId ?? "";

    // Immediately generate client secret
    const secretRes = await client.generateClientSecret(objectId);
    if (secretRes.isErr()) return err(secretRes.error);

    ctx.logger.info(`[${source}] Created bot AAD app: botId=${botId}`);

    return ok({
      outputs: {
        BOT_ID: botId,
        SECRET_BOT_PASSWORD: secretRes.value,
      },
    });
  },
});
