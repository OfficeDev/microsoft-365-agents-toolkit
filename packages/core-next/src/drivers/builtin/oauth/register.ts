// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as yaml from "js-yaml";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import {
  appStudioScopes,
  OauthRegistration,
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
  TokenExchangeMethodType,
} from "../../../clients/teamsDevPortal/types";

const MAX_NAME_LENGTH = 128;

const httpsUrl = z
  .string()
  .url()
  .refine((u) => u.startsWith("https://"), {
    message: "URL must use HTTPS",
  });

const inputSchema = z.object({
  /** Display name for the OAuth configuration (max 128 chars) */
  name: z.string().min(1).max(MAX_NAME_LENGTH),
  /** Teams app ID */
  appId: z.string().min(1),
  /** OAuth flow — only "authorizationCode" is supported */
  flow: z.literal("authorizationCode"),
  /** OAuth client ID */
  clientId: z.string().min(1),
  /** OAuth client secret (required for Custom provider unless PKCE is enabled) */
  clientSecret: z.string().optional(),
  /** Whether PKCE is enabled */
  isPKCEEnabled: z.boolean().optional(),
  /** Identity provider: "Custom" or "MicrosoftEntra" */
  identityProvider: z.enum(["Custom", "MicrosoftEntra"]).optional(),
  /** Token exchange method */
  tokenExchangeMethodType: z.nativeEnum(TokenExchangeMethodType).optional(),
  /** Who can use this registration */
  applicableToApps: z.nativeEnum(OauthRegistrationAppType).optional(),
  /** Tenant audience */
  targetAudience: z.nativeEnum(OauthRegistrationTargetAudience).optional(),
  /** Base URL of the API — optional when apiSpecPath is provided */
  baseUrl: httpsUrl.optional(),
  /** Path to OpenAPI spec file — used to derive domain when baseUrl is absent */
  apiSpecPath: z.string().optional(),
  /** Authorization URL — required for Custom provider */
  authorizationUrl: httpsUrl.optional(),
  /** Token URL — required for Custom provider */
  tokenUrl: httpsUrl.optional(),
  /** Refresh URL */
  refreshUrl: httpsUrl.optional(),
  /** OAuth scopes, comma-separated */
  scope: z.string().optional(),
  /** Existing configuration ID for idempotency */
  existingConfigurationId: z.string().optional(),
});

/**
 * Resolve `${{VAR}}` env-var placeholders using process.env.
 * The lifecycle executor syncs envMap into process.env before each driver call.
 */
function resolveEnvPlaceholders(value: string): string {
  return value.replace(/\$\{\{([^}]+)\}\}/g, (_match, name: string) => {
    return process.env[name] ?? "";
  });
}

/**
 * Extract server base URLs from an OpenAPI spec file.
 * Returns an array of resolved, HTTPS-only domain URLs.
 * Matches fx-core behaviour: when baseUrl is absent, derive from apiSpecPath.
 */
async function extractDomainsFromSpec(specPath: string, projectPath: string): Promise<string[]> {
  const absPath = path.isAbsolute(specPath) ? specPath : path.resolve(projectPath, specPath);
  const raw = await fs.readFile(absPath, "utf-8");
  // Support both JSON and YAML specs
  const spec = (absPath.endsWith(".json") ? JSON.parse(raw) : yaml.load(raw)) as {
    servers?: Array<{ url?: string }>;
  };
  if (!spec?.servers?.length) return [];

  return spec.servers
    .map((s) => {
      if (!s.url) return "";
      const resolved = resolveEnvPlaceholders(s.url);
      try {
        const u = new URL(resolved);
        // Return origin only (strip path); fx-core uses base URL sans path
        return u.protocol === "https:" ? u.origin : "";
      } catch {
        return "";
      }
    })
    .filter(Boolean);
}

/**
 * Driver: oauth/register
 *
 * Creates an OAuth configuration in the Teams Developer Portal.
 *
 * Outputs:
 * - configurationId  (mapped via writeToEnvironmentFile to the template's env var name)
 * - applicationIdUri  (MicrosoftEntra only — the resource identifier URI from TDP)
 */
export const oauthRegisterDriver = createDriver({
  id: "oauth/register",
  name: "Register OAuth Configuration",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "oauth/register";

    // Acquire M365 token for TDP
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

    // Idempotency: skip creation if existing config ID is provided and found
    if (config.existingConfigurationId) {
      const existingRes = await client.getOauthRegistration(config.existingConfigurationId);
      if (existingRes.isOk() && existingRes.value) {
        ctx.logger.info(
          `[${source}] OAuth config ${config.existingConfigurationId} already exists — skipping creation`
        );
        return ok({
          outputs: {
            configurationId: config.existingConfigurationId,
          },
        });
      }
    }

    // Validate that Custom provider has authorizationUrl and tokenUrl
    const isCustom = !config.identityProvider || config.identityProvider === "Custom";
    if (isCustom) {
      if (!config.authorizationUrl) {
        return err(
          userError(
            "MissingAuthorizationUrl",
            "authorizationUrl is required for Custom identity provider",
            { source }
          )
        );
      }
      if (!config.tokenUrl) {
        return err(
          userError("MissingTokenUrl", "tokenUrl is required for Custom identity provider", {
            source,
          })
        );
      }
    }

    // Resolve domain: prefer explicit baseUrl, fall back to apiSpecPath parsing
    // This matches fx-core where baseUrl is optional when apiSpecPath is provided.
    let domains: string[] = [];
    if (config.baseUrl) {
      domains = [config.baseUrl];
    } else if (config.apiSpecPath) {
      try {
        domains = await extractDomainsFromSpec(
          config.apiSpecPath,
          ctx.projectPath ?? process.cwd()
        );
      } catch (e: unknown) {
        ctx.logger.warning(
          `[${source}] Failed to extract domains from apiSpecPath: ${e instanceof Error ? e.message : String(e)}`
        );
      }
    }
    if (domains.length === 0) {
      return err(
        userError(
          "MissingBaseUrl",
          "Either baseUrl or apiSpecPath (with resolvable server URLs) is required",
          { source }
        )
      );
    }

    // Build the registration payload
    const applicableToApps = config.applicableToApps ?? OauthRegistrationAppType.AnyApp;
    const targetAudience = config.targetAudience ?? OauthRegistrationTargetAudience.AnyTenant;
    const tokenExchangeMethodType =
      config.tokenExchangeMethodType ?? TokenExchangeMethodType.BasicAuthorizationHeader;

    const registration: OauthRegistration =
      config.identityProvider === "MicrosoftEntra"
        ? {
            description: config.name,
            targetUrlsShouldStartWith: domains,
            applicableToApps,
            m365AppId:
              applicableToApps === OauthRegistrationAppType.SpecificApp ? config.appId : "",
            targetAudience,
            clientId: config.clientId,
            clientSecret: "",
            identityProvider: "MicrosoftEntra",
            tokenExchangeMethodType,
          }
        : {
            description: config.name,
            targetUrlsShouldStartWith: domains,
            applicableToApps,
            m365AppId:
              applicableToApps === OauthRegistrationAppType.SpecificApp ? config.appId : "",
            targetAudience,
            clientId: config.clientId,
            clientSecret: config.clientSecret ?? "",
            isPKCEEnabled: !!config.isPKCEEnabled,
            authorizationEndpoint: config.authorizationUrl,
            tokenExchangeEndpoint: config.tokenUrl,
            tokenRefreshEndpoint: config.refreshUrl,
            scopes: config.scope ? config.scope.split(",").map((s) => s.trim()) : [],
            identityProvider: "Custom",
            tokenExchangeMethodType,
          };

    const createRes = await client.createOauthRegistration(registration);
    if (createRes.isErr()) return err(createRes.error);

    const configId = createRes.value.configurationRegistrationId.oAuthConfigId;
    ctx.logger.info(`[${source}] Created OAuth config: ${configId}`);

    const outputs: Record<string, string> = {
      configurationId: configId,
    };

    // For MicrosoftEntra, the TDP response includes a resourceIdentifierUri
    // that must be propagated so the AAD manifest can reference it.
    if (config.identityProvider === "MicrosoftEntra" && createRes.value.resourceIdentifierUri) {
      outputs.applicationIdUri = createRes.value.resourceIdentifierUri;
    }

    return ok({ outputs });
  },
});
