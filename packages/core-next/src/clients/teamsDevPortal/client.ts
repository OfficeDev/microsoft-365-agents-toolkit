// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance } from "axios";
import { Result, ok, err } from "neverthrow";
import { AtkContext } from "../../core/context";
import { AtkError, userError, systemError } from "../../core/error";
import { createHttpClient } from "../../http/httpClient";
import { sendWithRetry } from "../../http/retry";
import {
  AppDefinition,
  PublishedAppDefinition,
  TDP_BASE_URL,
  OauthRegistration,
  OauthConfigurationId,
  ApiSecretRegistration,
} from "./types";
import { BotRegistration } from "../graphApi/types";

/**
 * Client for the Teams Developer Portal REST API.
 *
 * Unlike fx-core's singleton, this is instantiated per-operation and receives
 * an already-acquired M365 bearer token. No global state.
 */
export class TeamsDevPortalClient {
  private readonly axios: AxiosInstance;

  /** ZIP local-file-header magic bytes: PK\\x03\\x04 */
  private static readonly ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]);

  constructor(ctx: AtkContext, token: string) {
    this.axios = createHttpClient(ctx, { baseURL: TDP_BASE_URL });
    this.axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    this.axios.defaults.headers.common["Client-Source"] = "teamstoolkit";
  }

  /**
   * Import (create or update) a Teams app from a ZIP package.
   *
   * @param file - ZIP buffer containing manifest.json + icons
   * @param overwrite - If true, update an existing app with the same ID
   */
  async importApp(file: Buffer, overwrite = false): Promise<Result<AppDefinition, AtkError>> {
    const magicErr = this.validateZipBuffer(file, "importApp");
    if (magicErr) return err(magicErr);
    try {
      const response = await sendWithRetry(() =>
        this.axios.post("/api/appdefinitions/v2/import", file, {
          headers: { "Content-Type": "application/zip" },
          params: { overwriteIfAppAlreadyExists: overwrite },
        })
      );

      if (response.data) {
        return ok(response.data as AppDefinition);
      }
      return err(
        systemError(
          "TeamsDevPortalImportError",
          "Empty response from Developer Portal import API",
          {
            source: "TeamsDevPortalClient",
          }
        )
      );
    } catch (e: unknown) {
      return err(this.wrapAxiosError("importApp", e));
    }
  }

  /**
   * Get an app definition by its Teams App ID.
   */
  async getApp(teamsAppId: string): Promise<Result<AppDefinition, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/api/appdefinitions/${encodeURIComponent(teamsAppId)}`)
      );

      if (response.data) {
        const app = response.data as AppDefinition;
        if (app.teamsAppId === teamsAppId) {
          return ok(app);
        }
      }
      return err(
        userError("TeamsAppNotFound", `App ${teamsAppId} not found in Developer Portal`, {
          source: "TeamsDevPortalClient",
        })
      );
    } catch (e: unknown) {
      return err(this.wrapAxiosError("getApp", e));
    }
  }

  /**
   * Publish a Teams app to the organization catalog (first publish).
   * Returns the published app ID.
   */
  async publishTeamsApp(teamsAppId: string, file: Buffer): Promise<Result<string, AtkError>> {
    const magicErr = this.validateZipBuffer(file, "publishTeamsApp");
    if (magicErr) return err(magicErr);
    try {
      const response = await sendWithRetry(() =>
        this.axios.post("/api/publishing", file, {
          headers: { "Content-Type": "application/zip" },
        })
      );

      if (response.data?.error) {
        // Conflict — app already published; fall through to update
        if (
          response.data.error.code === "Conflict" &&
          response.data.error.innerError?.code === "AppDefinitionAlreadyExists"
        ) {
          return this.publishTeamsAppUpdate(teamsAppId, file);
        }
        return err(
          systemError(
            "TeamsAppPublishError",
            `Publish API error: ${JSON.stringify(response.data.error)}`,
            { source: "TeamsDevPortalClient" }
          )
        );
      }

      if (response.data?.id) {
        return ok(response.data.id as string);
      }
      return err(
        systemError("TeamsAppPublishError", "Empty response from publish API", {
          source: "TeamsDevPortalClient",
        })
      );
    } catch (e: unknown) {
      return err(this.wrapAxiosError("publishTeamsApp", e));
    }
  }

  /**
   * Update a previously published Teams app in the organization catalog.
   * Returns the updated published app ID.
   */
  async publishTeamsAppUpdate(teamsAppId: string, file: Buffer): Promise<Result<string, AtkError>> {
    const magicErr = this.validateZipBuffer(file, "publishTeamsAppUpdate");
    if (magicErr) return err(magicErr);
    try {
      const stagedRes = await this.getStagedApp(teamsAppId);
      if (stagedRes.isErr()) return err(stagedRes.error);

      const staged = stagedRes.value;
      if (!staged) {
        return err(
          userError(
            "TeamsAppNotPublished",
            `App ${teamsAppId} not found in published catalog — cannot update`,
            { source: "TeamsDevPortalClient" }
          )
        );
      }

      const response = await sendWithRetry(() =>
        this.axios.post(
          `/api/publishing/${encodeURIComponent(staged.teamsAppId)}/appdefinitions`,
          file,
          { headers: { "Content-Type": "application/zip" } }
        )
      );

      if (response.data?.error || response.data?.errorMessage) {
        return err(
          systemError(
            "TeamsAppPublishUpdateError",
            `Update publish API error: ${JSON.stringify(response.data.error ?? response.data.errorMessage)}`,
            { source: "TeamsDevPortalClient" }
          )
        );
      }

      if (response.data?.teamsAppId) {
        return ok(response.data.teamsAppId as string);
      }
      return err(
        systemError("TeamsAppPublishUpdateError", "Empty response from publish-update API", {
          source: "TeamsDevPortalClient",
        })
      );
    } catch (e: unknown) {
      return err(this.wrapAxiosError("publishTeamsAppUpdate", e));
    }
  }

  /**
   * Check if an app is already published in the tenant catalog.
   * Returns the latest published definition, or undefined if not found.
   */
  async getStagedApp(
    teamsAppId: string
  ): Promise<Result<PublishedAppDefinition | undefined, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/api/publishing/${encodeURIComponent(teamsAppId)}`)
      );

      if (response.data?.value?.[0]?.appDefinitions?.length) {
        const defs: PublishedAppDefinition[] = response.data.value[0].appDefinitions.map(
          (item: Record<string, unknown>) => ({
            teamsAppId: item.teamsAppId as string,
            displayName: item.displayName as string,
            publishingState: item.publishingState as string,
            lastModifiedDateTime: item.lastModifiedDateTime
              ? new Date(item.lastModifiedDateTime as string)
              : null,
          })
        );
        return ok(defs[defs.length - 1]);
      }
      return ok(undefined);
    } catch {
      // 404 or other failures — treat as "not published"
      return ok(undefined);
    }
  }

  /* ─── Bot Framework registration via TDP ─────────────────────── */

  /**
   * Get a bot registration by ID.
   * Returns undefined if the bot is not found (404 is swallowed).
   */
  async getBotRegistration(botId: string): Promise<Result<BotRegistration | undefined, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/api/botframework/${encodeURIComponent(botId)}`)
      );
      return ok(response.data as BotRegistration);
    } catch (e: unknown) {
      const axiosErr = e as { response?: { status?: number } };
      if (axiosErr.response?.status === 404) {
        return ok(undefined);
      }
      return err(this.wrapAxiosError("getBotRegistration", e));
    }
  }

  /**
   * Create a new bot registration.
   */
  async createBotRegistration(registration: BotRegistration): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(() => this.axios.post("/api/botframework", registration));
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapAxiosError("createBotRegistration", e));
    }
  }

  /**
   * Update an existing bot registration.
   */
  async updateBotRegistration(registration: BotRegistration): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(() =>
        this.axios.post(
          `/api/botframework/${encodeURIComponent(registration.botId!)}`,
          registration
        )
      );
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapAxiosError("updateBotRegistration", e));
    }
  }

  /* ─── OAuth configuration registration via TDP ───────────────── */

  /**
   * Get an OAuth configuration by ID.
   * Returns undefined if not found (404 is swallowed).
   */
  async getOauthRegistration(
    oAuthConfigId: string
  ): Promise<Result<OauthRegistration | undefined, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/api/v1.0/oAuthConfigurations/${encodeURIComponent(oAuthConfigId)}`)
      );
      return ok(response.data as OauthRegistration);
    } catch (e: unknown) {
      const axiosErr = e as { response?: { status?: number } };
      if (axiosErr.response?.status === 404) {
        return ok(undefined);
      }
      return err(this.wrapAxiosError("getOauthRegistration", e));
    }
  }

  /**
   * Create an OAuth configuration registration.
   * Returns the configuration ID and resource identifier URI.
   */
  async createOauthRegistration(
    registration: OauthRegistration
  ): Promise<Result<OauthConfigurationId, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.post("/api/v1.0/oAuthConfigurations", registration)
      );
      return ok(response.data as OauthConfigurationId);
    } catch (e: unknown) {
      return err(this.wrapAxiosError("createOauthRegistration", e));
    }
  }

  /* ─── API Key (secret) registration via TDP ──────────────────── */

  /**
   * Get an API key registration by ID.
   * Returns undefined if not found (404 is swallowed).
   */
  async getApiKeyRegistration(
    registrationId: string
  ): Promise<Result<ApiSecretRegistration | undefined, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/api/v1.0/apiSecretRegistrations/${encodeURIComponent(registrationId)}`)
      );
      return ok(response.data as ApiSecretRegistration);
    } catch (e: unknown) {
      const axiosErr = e as { response?: { status?: number } };
      if (axiosErr.response?.status === 404) {
        return ok(undefined);
      }
      return err(this.wrapAxiosError("getApiKeyRegistration", e));
    }
  }

  /**
   * Create an API key (secret) registration.
   * Returns the created registration with its ID.
   */
  async createApiKeyRegistration(
    registration: ApiSecretRegistration
  ): Promise<Result<ApiSecretRegistration, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.post("/api/v1.0/apiSecretRegistrations", registration)
      );
      return ok(response.data as ApiSecretRegistration);
    } catch (e: unknown) {
      return err(this.wrapAxiosError("createApiKeyRegistration", e));
    }
  }

  /**
   * Validate that a buffer has the ZIP local-file-header magic bytes.
   * Returns an AtkError if invalid, or undefined if OK.
   */
  private validateZipBuffer(file: Buffer, caller: string): AtkError | undefined {
    if (file.length < 4 || file.subarray(0, 4).compare(TeamsDevPortalClient.ZIP_MAGIC) !== 0) {
      return userError(
        "InvalidAppPackage",
        `${caller}: Buffer is not a valid ZIP file (missing PK magic header)`,
        { source: "TeamsDevPortalClient" }
      );
    }
    return undefined;
  }

  /**
   * Map an Axios error (or any thrown value) to an AtkError.
   */
  private wrapAxiosError(apiName: string, e: unknown): AtkError {
    const axiosErr = e as { response?: { status?: number; data?: unknown }; message?: string };
    const status = axiosErr.response?.status;
    const data = axiosErr.response?.data;
    const message = axiosErr.message ?? String(e);

    const detail = data ? ` — data: ${JSON.stringify(data)}` : "";
    const statusPart = status ? `Status ${status}. ` : "";

    if (status === 409) {
      return userError(
        "TeamsAppConflictError",
        `${apiName}: ${statusPart}App already exists in another tenant${detail}`,
        { source: "TeamsDevPortalClient", help: "https://aka.ms/teamsfx-switch-tenant" }
      );
    }

    if (status === 404) {
      return userError("TeamsAppNotFound", `${apiName}: ${statusPart}App not found${detail}`, {
        source: "TeamsDevPortalClient",
      });
    }

    if (status && status >= 400 && status < 500) {
      return userError("TeamsDevPortalApiError", `${apiName}: ${statusPart}${message}${detail}`, {
        source: "TeamsDevPortalClient",
      });
    }

    return systemError("TeamsDevPortalApiError", `${apiName}: ${statusPart}${message}${detail}`, {
      source: "TeamsDevPortalClient",
      inner: e instanceof Error ? e : undefined,
    });
  }
}
