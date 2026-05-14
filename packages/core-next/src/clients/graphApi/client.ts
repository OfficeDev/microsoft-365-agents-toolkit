// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance } from "axios";
import { Result, ok, err } from "neverthrow";
import { AtkContext } from "../../core/context";
import { AtkError, userError, systemError } from "../../core/error";
import { createHttpClient } from "../../http/httpClient";
import { sendWithRetry } from "../../http/retry";
import {
  AADApplication,
  AadOwner,
  PasswordCredential,
  PublishedAppDefinition,
  PublishingState,
  GRAPH_BASE_URL,
  GRAPH_BETA_URL,
} from "./types";

/**
 * Client for Microsoft Graph API — Entra ID (Azure AD) application operations.
 *
 * Instantiated per-operation with an already-acquired Graph bearer token.
 */
export class GraphApiClient {
  private readonly axios: AxiosInstance;
  private readonly betaAxios: AxiosInstance;

  constructor(ctx: AtkContext, token: string) {
    this.axios = createHttpClient(ctx, { baseURL: GRAPH_BASE_URL });
    this.axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    this.axios.defaults.headers.common["Content-Type"] = "application/json";

    this.betaAxios = createHttpClient(ctx, { baseURL: GRAPH_BETA_URL });
    this.betaAxios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    this.betaAxios.defaults.headers.common["Content-Type"] = "application/json";
  }

  /**
   * Create a new Entra ID application.
   */
  async createAadApp(
    displayName: string,
    signInAudience?: string,
    serviceManagementReference?: string
  ): Promise<Result<AADApplication, AtkError>> {
    const body: Record<string, unknown> = { displayName };
    if (signInAudience) body.signInAudience = signInAudience;
    if (serviceManagementReference) body.serviceManagementReference = serviceManagementReference;

    try {
      const response = await sendWithRetry(() => this.axios.post("/applications", body));
      return ok(response.data as AADApplication);
    } catch (e: unknown) {
      return err(this.wrapError("createAadApp", e));
    }
  }

  /**
   * Generate a client secret (password credential) for an application.
   *
   * Uses extra retries for 404 to handle Entra replication delay.
   */
  async generateClientSecret(
    objectId: string,
    expiresInDays = 180,
    description = "default"
  ): Promise<Result<string, AtkError>> {
    const now = new Date();
    const end = new Date(now);
    end.setDate(end.getDate() + expiresInDays);

    const body = {
      passwordCredential: {
        displayName: description,
        startDateTime: now.toISOString(),
        endDateTime: end.toISOString(),
      },
    };

    try {
      const response = await sendWithRetry(
        () => this.axios.post(`/applications/${objectId}/addPassword`, body),
        5 // extra retries for Entra sync delay
      );
      const cred = response.data as PasswordCredential;
      if (!cred.secretText) {
        return err(
          systemError("EmptyClientSecret", "Graph API returned an empty secretText", {
            source: "GraphApiClient",
          })
        );
      }
      return ok(cred.secretText);
    } catch (e: unknown) {
      return err(this.wrapError("generateClientSecret", e));
    }
  }

  /**
   * Update an Entra ID application (PATCH semantics).
   */
  async updateAadApp(
    objectId: string,
    updates: Partial<AADApplication>
  ): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(
        () => this.axios.patch(`/applications/${objectId}`, updates),
        5 // extra retries for 404/400 race conditions during permission sync
      );
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapError("updateAadApp", e));
    }
  }

  /**
   * Get the owners of an application.
   */
  async getOwners(objectId: string): Promise<Result<AadOwner[], AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/applications/${objectId}/owners`)
      );
      return ok((response.data?.value ?? []) as AadOwner[]);
    } catch (e: unknown) {
      return err(this.wrapError("getOwners", e));
    }
  }

  /**
   * Add an owner to an application.
   */
  async addOwner(objectId: string, userObjectId: string): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(() =>
        this.axios.post(`/applications/${objectId}/owners/$ref`, {
          "@odata.id": `${GRAPH_BASE_URL}/directoryObjects/${userObjectId}`,
        })
      );
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapError("addOwner", e));
    }
  }

  /* ------------------------------------------------------------------ */
  /*  App Catalog — publish Teams apps via Graph                        */
  /* ------------------------------------------------------------------ */

  private static readonly teamsAppsPath = "/appCatalogs/teamsApps";

  /**
   * Check if a Teams app is already published in the tenant catalog.
   * Returns the latest published definition, or `undefined` if not found.
   * Swallows all errors and returns `undefined` (mirrors fx-core behavior).
   */
  async getStagedApp(
    teamsAppExternalId: string
  ): Promise<Result<PublishedAppDefinition | undefined, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.betaAxios.get(
          `${GraphApiClient.teamsAppsPath}?$filter=externalId eq '${teamsAppExternalId}'&$expand=appDefinitions`
        )
      );

      if (!response?.data?.value || response.data.value.length === 0) {
        return ok(undefined);
      }

      const appDefinitions = response.data.value[0].appDefinitions;
      if (!Array.isArray(appDefinitions) || appDefinitions.length === 0) {
        return ok(undefined);
      }

      const latest = appDefinitions[appDefinitions.length - 1];
      return ok({
        lastModifiedDateTime: latest.lastModifiedDateTime
          ? new Date(latest.lastModifiedDateTime)
          : null,
        publishingState: latest.publishingState as PublishingState,
        teamsAppId: response.data.value[0].id,
        displayName: response.data.value[0].displayName,
      });
    } catch {
      // 404 or other failures — treat as "not published"
      return ok(undefined);
    }
  }

  /**
   * Publish a Teams app to the organization catalog (first publish).
   *
   * Handles resilient edge cases from the Graph API:
   * - BadGateway → falls back to getStagedApp to retrieve the ID
   * - 409 Conflict / AppDefinitionAlreadyExists → falls through to publishTeamsAppUpdate
   */
  async publishTeamsApp(
    teamsAppExternalId: string,
    file: Buffer
  ): Promise<Result<string, AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.betaAxios.post(`${GraphApiClient.teamsAppsPath}?requiresReview=true`, file, {
          headers: { "Content-Type": "application/zip" },
        })
      );

      // Graph may return 200 with an error body
      if (response?.data?.error) {
        if (response.data.error.code === "BadGateway") {
          const stagedRes = await this.getStagedApp(teamsAppExternalId);
          if (stagedRes.isOk() && stagedRes.value) {
            return ok(stagedRes.value.teamsAppId);
          }
        }

        if (
          response.data.error.code === "Conflict" &&
          response.data.error.innerError?.code === "AppDefinitionAlreadyExists"
        ) {
          return this.publishTeamsAppUpdate(teamsAppExternalId, file);
        }

        return err(
          systemError(
            "GraphPublishError",
            `[publishTeamsApp] ${response.data.error.message ?? JSON.stringify(response.data.error)}`,
            { source: "GraphApiClient" }
          )
        );
      }

      if (response?.data?.id) {
        return ok(response.data.id as string);
      }

      // Fallback: query the staged app to get the ID
      const stagedRes = await this.getStagedApp(teamsAppExternalId);
      if (stagedRes.isOk() && stagedRes.value?.teamsAppId) {
        return ok(stagedRes.value.teamsAppId);
      }

      return err(
        systemError("GraphPublishError", "[publishTeamsApp] Empty response from Graph API", {
          source: "GraphApiClient",
        })
      );
    } catch (e: unknown) {
      // HTTP 409 — app already exists, fall through to update
      if (e && typeof e === "object" && "response" in e) {
        const status = (e as any).response?.status;
        if (status === 409) {
          return this.publishTeamsAppUpdate(teamsAppExternalId, file);
        }
      }
      return err(this.wrapError("publishTeamsApp", e));
    }
  }

  /**
   * Update a previously published Teams app in the organization catalog.
   * Looks up the internal catalog ID via `getStagedApp`, then POSTs the new ZIP.
   *
   * If the update with `requiresReview=true` fails with 400, retries without
   * the parameter. This handles apps that were sideloaded via Shared scope
   * (`extendToM365`) — they appear in the catalog but don't support the
   * admin-review update path.
   */
  async publishTeamsAppUpdate(
    teamsAppExternalId: string,
    file: Buffer
  ): Promise<Result<string, AtkError>> {
    const stagedRes = await this.getStagedApp(teamsAppExternalId);
    if (stagedRes.isErr()) return err(stagedRes.error);

    const staged = stagedRes.value;
    if (!staged) {
      return err(
        userError(
          "TeamsAppNotPublished",
          `[publishTeamsAppUpdate] Published app not found for externalId: ${teamsAppExternalId}`,
          { source: "GraphApiClient" }
        )
      );
    }

    const url = `${GraphApiClient.teamsAppsPath}/${staged.teamsAppId}/appDefinitions`;
    const headers = { "Content-Type": "application/zip" };

    let response;
    try {
      response = await sendWithRetry(() =>
        this.betaAxios.post(`${url}?requiresReview=true`, file, { headers })
      );
    } catch (e: unknown) {
      // 400 can occur when the app was sideloaded (Shared scope) rather than
      // published through the admin-review flow. Retry without requiresReview.
      if (e && typeof e === "object" && "response" in e && (e as any).response?.status === 400) {
        try {
          response = await sendWithRetry(() => this.betaAxios.post(url, file, { headers }));
        } catch (e2: unknown) {
          // 404 after 400 means the app was sideloaded (Shared scope) and
          // appears in the catalog query but doesn't have a real REST resource.
          // The app IS published through sideloading — return the existing ID.
          if (
            e2 &&
            typeof e2 === "object" &&
            "response" in e2 &&
            (e2 as any).response?.status === 404
          ) {
            return ok(staged.teamsAppId);
          }
          return err(this.wrapError("publishTeamsAppUpdate", e2));
        }
      } else {
        return err(this.wrapError("publishTeamsAppUpdate", e));
      }
    }

    if (response?.data?.error || response?.data?.errorMessage) {
      return err(
        systemError(
          "GraphPublishUpdateError",
          `[publishTeamsAppUpdate] ${response.data.error?.message ?? response.data.errorMessage}`,
          { source: "GraphApiClient" }
        )
      );
    }

    if (response?.data?.teamsAppId) {
      return ok(response.data.teamsAppId as string);
    }
    if (response?.data?.id) {
      return ok(response.data.id as string);
    }
    // Fall back to the known catalog ID
    return ok(staged.teamsAppId);
  }

  /**
   * Remove a published Teams app from the organization catalog.
   *
   * @param catalogAppId The app catalog internal ID (returned by `publishTeamsApp` /
   *   `getStagedApp`), **not** the manifest externalId.
   *
   * Graph endpoint: `DELETE /appCatalogs/teamsApps/{id}`
   * Returns 204 No Content on success.
   */
  async unpublishTeamsApp(catalogAppId: string): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(() =>
        this.betaAxios.delete(`${GraphApiClient.teamsAppsPath}/${catalogAppId}`)
      );
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapError("unpublishTeamsApp", e));
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Error mapping                                                      */
  /* ------------------------------------------------------------------ */

  private wrapError(apiName: string, e: unknown): AtkError {
    if (e && typeof e === "object" && "response" in e) {
      const resp = (e as any).response;
      const status: number = resp?.status ?? 0;
      const graphCode: string = resp?.data?.error?.code ?? "";
      const graphMsg: string = resp?.data?.error?.message ?? String(e);

      if (status >= 400 && status < 500) {
        return userError(`GraphApiError_${graphCode || status}`, `[${apiName}] ${graphMsg}`, {
          source: "GraphApiClient",
        });
      }
      return systemError(`GraphApiError_${status}`, `[${apiName}] ${graphMsg}`, {
        source: "GraphApiClient",
      });
    }
    return systemError(
      "GraphApiUnexpectedError",
      `[${apiName}] ${e instanceof Error ? e.message : String(e)}`,
      { source: "GraphApiClient", inner: e instanceof Error ? e : undefined }
    );
  }
}
