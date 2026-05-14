// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance } from "axios";
import { Result, ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import AdmZip from "adm-zip";
import FormData from "form-data";
import { AtkContext } from "../../core/context";
import { AtkError, userError, systemError } from "../../core/error";
import { createHttpClient } from "../../http/httpClient";
import { sendWithRetry } from "../../http/retry";

/** MOS3 endpoint for public cloud. */
const MOS_ENDPOINT = "https://titles.prod.mos.microsoft.com";

/** M365 sideloading token scope. */
export const mosServiceScopes = (): string[] => [`${MOS_ENDPOINT}/.default`];

/** App scope for sideloading. */
export enum AppScope {
  Personal = "Personal",
  Shared = "Shared",
  Tenant = "Tenant",
}

/** Result type for sideLoad: [titleId, appId, shareLink]. */
export interface SideLoadResult {
  titleId: string;
  appId: string;
  shareLink: string;
}

/**
 * Check if a manifest describes a Declarative Agent app.
 */
function isDeclarativeAgentManifest(manifest: Record<string, unknown>): boolean {
  const agents = (manifest as { copilotAgents?: { declarativeAgents?: unknown[] } }).copilotAgents
    ?.declarativeAgents;
  return Array.isArray(agents) && agents.length > 0;
}

/**
 * M365 PackageService client for sideloading apps.
 *
 * Uses `createHttpClient(ctx)` for telemetry-wired requests and
 * `sendWithRetry()` for transient failure resilience. All public methods
 * return `Result<T, AtkError>` — callers don't need try/catch.
 */
export class M365PackageService {
  private readonly http: AxiosInstance;
  private readonly initEndpoint: string;

  constructor(ctx: AtkContext, token: string, endpoint = MOS_ENDPOINT) {
    this.initEndpoint = endpoint;
    this.http = createHttpClient(ctx, { timeout: 60_000 });
    this.http.defaults.headers.common["Authorization"] = `Bearer ${token}`;
  }

  /**
   * Resolve the actual titles service URL from the MOS config endpoint.
   */
  private async getTitleServiceUrl(): Promise<string> {
    const res = await sendWithRetry(() =>
      this.http.get("/config/v1/environment", {
        baseURL: this.initEndpoint,
      })
    );
    return res.data.titlesServiceUrl as string;
  }

  /**
   * Sideload an app package. Returns `{ titleId, appId, shareLink }`.
   */
  async sideLoad(
    packagePath: string,
    appScope = AppScope.Personal
  ): Promise<Result<SideLoadResult, AtkError>> {
    try {
      const zip = new AdmZip(packagePath);
      const manifestEntry = zip.getEntry("manifest.json");
      if (!manifestEntry) {
        return err(
          userError("InvalidAppPackage", "Invalid app package zip — manifest.json is missing", {
            source: "M365PackageService",
          })
        );
      }
      const manifest = JSON.parse(manifestEntry.getData().toString("utf8")) as Record<
        string,
        unknown
      >;

      if (isDeclarativeAgentManifest(manifest)) {
        const v2Res = await this.sideLoadV2(packagePath, appScope);
        if (v2Res.isErr()) return err(v2Res.error);
        const [titleId, appId] = v2Res.value;
        let shareLink = "";
        if (appScope === AppScope.Shared) {
          shareLink = await this.getShareLink(titleId);
        }
        return ok({ titleId, appId, shareLink });
      } else {
        const v1Res = await this.sideLoadV1(packagePath);
        if (v1Res.isErr()) return err(v1Res.error);
        const [titleId, appId] = v1Res.value;
        return ok({ titleId, appId, shareLink: "" });
      }
    } catch (e: unknown) {
      return err(this.wrapError("sideLoad", e));
    }
  }

  /**
   * V2 sideloading (Builder API) — used for declarative agents.
   */
  private async sideLoadV2(
    packagePath: string,
    appScope: AppScope
  ): Promise<Result<[string, string], AtkError>> {
    try {
      const data = await fs.readFile(packagePath);
      const form = new FormData();
      form.append("package", data);
      form.append("info", JSON.stringify({ builderName: "TeamsToolKit" }));
      const serviceUrl = await this.getTitleServiceUrl();

      const uploadRes = await sendWithRetry(() =>
        this.http.post("/builder/v1/users/packages", form, {
          baseURL: serviceUrl,
          headers: form.getHeaders(),
          params: { scope: appScope },
        })
      );

      const statusId: string = uploadRes.data.statusId;

      // Poll for completion
      const maxAttempts = 60;
      for (let i = 0; i < maxAttempts; i++) {
        const statusRes = await sendWithRetry(() =>
          this.http.get(`/builder/v1/users/packages/status/${statusId}`, {
            baseURL: serviceUrl,
          })
        );
        if (statusRes.status === 200 && statusRes.data.titleId) {
          return ok([statusRes.data.titleId as string, statusRes.data.appId as string]);
        }
        await delay(5_000);
      }
      return err(
        systemError("SideloadingTimeout", "Sideloading timed out waiting for package status.", {
          source: "M365PackageService",
        })
      );
    } catch (e: unknown) {
      return err(this.wrapError("sideLoadV2", e));
    }
  }

  /**
   * V1 sideloading — used for classic Teams apps.
   */
  private async sideLoadV1(packagePath: string): Promise<Result<[string, string], AtkError>> {
    try {
      const data = await fs.readFile(packagePath);
      const form = new FormData();
      form.append("package", data);
      const serviceUrl = await this.getTitleServiceUrl();

      const uploadRes = await sendWithRetry(() =>
        this.http.post("/dev/v1/users/packages", form.getBuffer(), {
          baseURL: serviceUrl,
          headers: form.getHeaders(),
        })
      );

      const operationId: string = uploadRes.data.operationId;

      // Acquire
      const acquireRes = await sendWithRetry(() =>
        this.http.post(
          "/dev/v1/users/packages/acquisitions",
          { operationId },
          { baseURL: serviceUrl }
        )
      );
      const statusId: string = acquireRes.data.statusId;

      // Poll for completion
      const maxAttempts = 60;
      for (let i = 0; i < maxAttempts; i++) {
        const statusRes = await sendWithRetry(() =>
          this.http.get(`/dev/v1/users/packages/status/${statusId}`, {
            baseURL: serviceUrl,
          })
        );
        if (statusRes.status === 200 && statusRes.data.titleId) {
          return ok([statusRes.data.titleId as string, statusRes.data.appId as string]);
        }
        await delay(5_000);
      }
      return err(
        systemError("SideloadingTimeout", "Sideloading timed out waiting for package status.", {
          source: "M365PackageService",
        })
      );
    } catch (e: unknown) {
      return err(this.wrapError("sideLoadV1", e));
    }
  }

  /**
   * Get a share link for a sideloaded title.
   */
  private async getShareLink(titleId: string): Promise<string> {
    try {
      const serviceUrl = await this.getTitleServiceUrl();
      const res = await sendWithRetry(() =>
        this.http.get(`/marketplace/v1/users/titles/${titleId}/sharingInfo`, {
          baseURL: serviceUrl,
        })
      );
      return (res.data?.shareLink as string) ?? "";
    } catch {
      return "";
    }
  }

  /** Wrap a caught error into an AtkError based on HTTP status. */
  private wrapError(method: string, e: unknown): AtkError {
    if (e && typeof e === "object" && "response" in e) {
      const status = (e as { response?: { status?: number } }).response?.status;
      const data = (e as { response?: { data?: unknown } }).response?.data;
      const body = typeof data === "string" ? data : JSON.stringify(data ?? "");
      if (status && status >= 400 && status < 500) {
        return userError("M365ServiceUserError", `[${method}] HTTP ${status}: ${body}`, {
          source: "M365PackageService",
        });
      }
      return systemError("M365ServiceError", `[${method}] HTTP ${status}: ${body}`, {
        source: "M365PackageService",
      });
    }
    return systemError(
      "M365ServiceError",
      `[${method}] ${e instanceof Error ? e.message : String(e)}`,
      { source: "M365PackageService", inner: e instanceof Error ? e : undefined }
    );
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
