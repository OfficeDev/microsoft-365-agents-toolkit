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
  /** Display name for the Entra app (max 120 chars) */
  name: z.string().min(1).max(MAX_NAME_LENGTH),
  /** Whether to auto-generate a client secret */
  generateClientSecret: z.boolean(),
  /** Sign-in audience — defaults to AzureADMyOrg */
  signInAudience: z.nativeEnum(SignInAudience).optional().default(SignInAudience.AzureADMyOrg),
  /** Credential lifetime in days (default 180) */
  clientSecretExpireDays: z.number().int().positive().optional().default(180),
  /** Credential description */
  clientSecretDescription: z.string().optional().default("default"),
  /** Service management reference (service tree ID) */
  serviceManagementReference: z.string().optional(),
  /**
   * Existing client ID for idempotency.
   * When the lifecycle executor resolves ${{AAD_APP_CLIENT_ID}} from the env map,
   * this will be a UUID if the app was already created.
   */
  existingClientId: z.string().optional(),
  /** Existing object ID (needed if skipping creation but generating secret) */
  existingObjectId: z.string().optional(),
  /** Existing client secret for idempotency */
  existingClientSecret: z.string().optional(),
});

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const ENTRA_LOGIN_HOST = "https://login.microsoftonline.com";

/**
 * Driver: aadApp/create
 *
 * Creates an Entra ID (Azure AD) application registration and optionally
 * generates a client secret.
 *
 * Outputs:
 * - AAD_APP_CLIENT_ID
 * - AAD_APP_OBJECT_ID
 * - AAD_APP_TENANT_ID
 * - AAD_APP_OAUTH_AUTHORITY_HOST
 * - AAD_APP_OAUTH_AUTHORITY
 * - SECRET_AAD_APP_CLIENT_SECRET (if generateClientSecret)
 */
export const createAadAppDriver = createDriver({
  id: "aadApp/create",
  name: "Create AAD App",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "aadApp/create";

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

    let clientId = config.existingClientId;
    let objectId = config.existingObjectId;
    let tenantId: string | undefined;

    // Extract tenant ID from the JWT token (Graph token contains tid claim)
    try {
      const payload = JSON.parse(Buffer.from(tokenRes.value.split(".")[1], "base64").toString());
      tenantId = payload.tid;
    } catch {
      // Non-fatal — tenantId will be empty
    }

    // Idempotency: skip creation if we already have a valid client ID
    if (clientId && UUID_RE.test(clientId)) {
      ctx.logger.info(`[${source}] AAD app ${clientId} already exists — skipping creation`);
    } else {
      // Create the app
      const createRes = await client.createAadApp(
        config.name,
        config.signInAudience,
        config.serviceManagementReference
      );
      if (createRes.isErr()) return err(createRes.error);

      clientId = createRes.value.appId ?? "";
      objectId = createRes.value.id ?? "";
      ctx.logger.info(`[${source}] Created AAD app: clientId=${clientId}, objectId=${objectId}`);
    }

    // Generate client secret if requested and not already present
    let clientSecret = config.existingClientSecret;
    if (config.generateClientSecret && !clientSecret) {
      if (!objectId || !UUID_RE.test(objectId)) {
        return err(
          userError(
            "MissingObjectId",
            "Cannot generate client secret without a valid AAD_APP_OBJECT_ID",
            {
              source,
            }
          )
        );
      }
      const secretRes = await client.generateClientSecret(
        objectId,
        config.clientSecretExpireDays,
        config.clientSecretDescription
      );
      if (secretRes.isErr()) return err(secretRes.error);
      clientSecret = secretRes.value;
    }

    const outputs: Record<string, string> = {
      AAD_APP_CLIENT_ID: clientId ?? "",
      AAD_APP_OBJECT_ID: objectId ?? "",
      AAD_APP_TENANT_ID: tenantId ?? "",
      AAD_APP_OAUTH_AUTHORITY_HOST: ENTRA_LOGIN_HOST,
      AAD_APP_OAUTH_AUTHORITY: tenantId ? `${ENTRA_LOGIN_HOST}/${tenantId}` : "",
    };

    if (config.generateClientSecret) {
      outputs["SECRET_AAD_APP_CLIENT_SECRET"] = clientSecret ?? "";
    }

    return ok({ outputs });
  },
});
