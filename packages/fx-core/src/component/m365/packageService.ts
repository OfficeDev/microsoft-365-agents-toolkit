// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { LogProvider, SystemError, TeamsAppManifest, UserError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import FormData from "form-data";
import fs from "fs-extra";
import { ErrorContextMW, TOOLS } from "../../common/globalVars";
import { assembleError } from "../../error/common";
import { ErrorCategory } from "../../error/types";
import {
  Component,
  TelemetryEvent,
  TelemetryProperty,
  sendTelemetryErrorEvent,
  sendTelemetryEvent,
} from "../../common/telemetry";
import { waitSeconds } from "../../common/utils";
import { WrappedAxiosClient } from "../../common/wrappedAxiosClient";
import { NotExtendedToM365Error } from "./errors";
import { MosServiceEndpoint } from "./serviceConstant";
import { IsDeclarativeAgentManifest } from "../../common/projectTypeChecker";
import stripBom from "strip-bom";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";

const M365ErrorSource = "M365";
const M365ErrorComponent = "PackageService";

export enum AppScope {
  Personal = "Personal",
  Shared = "Shared",
}

// Call m365 service for package CRUD
export class PackageService {
  private static sharedInstance: PackageService;

  private readonly axiosInstance;
  private readonly initEndpoint;
  private readonly logger: LogProvider | undefined;

  public static GetSharedInstance(): PackageService {
    if (!PackageService.sharedInstance) {
      PackageService.sharedInstance = new PackageService(
        process.env.SIDELOADING_SERVICE_ENDPOINT ?? MosServiceEndpoint,
        TOOLS.logProvider
      );
    }
    return PackageService.sharedInstance;
  }

  public constructor(endpoint: string, logger?: LogProvider) {
    this.axiosInstance = WrappedAxiosClient.create({
      timeout: 30000,
    });
    this.initEndpoint = endpoint;
    this.logger = logger;
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async getTitleServiceUrl(token: string): Promise<string> {
    try {
      try {
        new URL(this.initEndpoint);
      } catch (_) {
        throw new Error("Invalid URL. Mis-configuration SIDELOADING_SERVICE_ENDPOINT.");
      }

      const envInfo = await this.axiosInstance.get("/config/v1/environment", {
        baseURL: this.initEndpoint,
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
      this.logger?.debug(JSON.stringify(envInfo.data));
      new URL(envInfo.data.titlesServiceUrl);
      return envInfo.data.titlesServiceUrl;
    } catch (error: any) {
      this.logger?.error(`Get ServiceUrl failed. ${error.message as string}`);
      throw error;
    }
  }

  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async sideLoadXmlManifest(token: string, manifestPath: string): Promise<[string, string]> {
    try {
      const data = await fs.readFile(manifestPath);
      const content = new FormData();
      content.append("package", data);
      content.append("culture", "en-us");
      const serviceUrl = await this.getTitleServiceUrl(token);
      const uploadResponse = await this.axiosInstance.post(
        "/dev/v1/users/packages/addins",
        content.getBuffer(),
        {
          baseURL: serviceUrl,
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/xml",
          },
        }
      );
      if (uploadResponse.status === 200) {
        const titleId: string = uploadResponse.data.titleId;
        const appId: string = uploadResponse.data.appId;
        this.logger?.info(`TitleId: ${titleId}`);
        this.logger?.info(`AppId: ${appId}`);
        this.logger?.verbose("Sideloading done.");
        return [titleId, appId];
      } else if (uploadResponse.status === 202) {
        const statusId = uploadResponse.data.statusId;
        this.logger?.debug(`Acquiring package with statusId: ${statusId as string} ...`);
        do {
          const statusResponse = await this.axiosInstance.get(
            `/dev/v1/users/packages/status/${statusId as string}`,
            {
              baseURL: serviceUrl,
              headers: { Authorization: `Bearer ${token}` },
            }
          );
          const resCode = statusResponse.status;
          this.logger?.debug(`Package status: ${resCode} ...`);
          if (resCode === 200) {
            const titleId: string = statusResponse.data.titleId;
            const appId: string = statusResponse.data.appId;
            this.logger?.info(`TitleId: ${titleId}`);
            this.logger?.info(`AppId: ${appId}`);
            this.logger?.verbose("Sideloading done.");
            return [titleId, appId];
          } else {
            await waitSeconds(2);
          }
        } while (true);
      } else {
        throw new Error(`Unknown response code: ${uploadResponse.status}}`);
      }
    } catch (error: any) {
      // this.logger?.error("Sideloading failed.");
      if (error.response) {
        // this.logger?.error(JSON.stringify(error.response.data));
        error = this.traceError(error);
      } else {
        // this.logger?.error(error.message);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }

  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async sideLoading(
    token: string,
    packagePath: string,
    appScope = AppScope.Personal
  ): Promise<[string, string, string]> {
    const manifest = this.getManifestFromZip(packagePath);
    if (!manifest) {
      throw new Error("Invalid app package zip. manifest.json is missing");
    }
    const isDelcarativeAgentApp = IsDeclarativeAgentManifest(manifest);
    if (
      isDelcarativeAgentApp &&
      featureFlagManager.getBooleanValue(FeatureFlags.BuilderAPIEnabled)
    ) {
      const res = await this.sideLoadingV2(token, packagePath, appScope);
      let shareLink = "";
      if (appScope == AppScope.Shared) {
        shareLink = await this.getShareLink(token, res[0]);
      }
      return [res[0], res[1], shareLink];
    } else {
      const res = await this.sideLoadingV1(token, packagePath);
      return [res[0], res[1], ""];
    }
  }
  // Side loading using Builder API
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async sideLoadingV2(
    token: string,
    manifestPath: string,
    appScope: AppScope
  ): Promise<[string, string]> {
    try {
      this.checkZip(manifestPath);
      const data = await fs.readFile(manifestPath);
      const content = new FormData();
      content.append("package", data);
      const serviceUrl = await this.getTitleServiceUrl(token);
      this.logger?.debug("Uploading package with sideLoading V2 ...");
      const uploadHeaders = content.getHeaders();
      uploadHeaders["Authorization"] = `Bearer ${token}`;
      const uploadResponse = await this.axiosInstance.post(
        "/builder/v1/users/packages",
        content.getBuffer(),
        {
          baseURL: serviceUrl,
          headers: uploadHeaders,
          params: {
            scope: appScope,
          },
        }
      );

      const statusId = uploadResponse.data.statusId;
      this.logger?.debug(`Acquiring package with statusId: ${statusId as string} ...`);

      do {
        const statusResponse = await this.axiosInstance.get(
          `/builder/v1/users/packages/status/${statusId as string}`,
          {
            baseURL: serviceUrl,
            headers: { Authorization: `Bearer ${token}` },
          }
        );
        const resCode = statusResponse.status;
        this.logger?.debug(`Package status: ${resCode} ...`);
        if (resCode === 200) {
          const titleId: string = statusResponse.data.titleId;
          const appId: string = statusResponse.data.appId;
          this.logger?.info(`TitleId: ${titleId}`);
          this.logger?.info(`AppId: ${appId}`);
          this.logger?.verbose("Sideloading done.");
          return [titleId, appId];
        } else {
          await waitSeconds(2);
        }
      } while (true);
    } catch (error: any) {
      if (error.response) {
        error = this.traceError(error);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async sideLoadingV1(token: string, manifestPath: string): Promise<[string, string]> {
    try {
      this.checkZip(manifestPath);
      const data = await fs.readFile(manifestPath);
      const content = new FormData();
      content.append("package", data);
      const serviceUrl = await this.getTitleServiceUrl(token);
      this.logger?.debug("Uploading package with sideLoading V1 ...");
      const uploadHeaders = content.getHeaders();
      uploadHeaders["Authorization"] = `Bearer ${token}`;
      const uploadResponse = await this.axiosInstance.post(
        "/dev/v1/users/packages",
        content.getBuffer(),
        {
          baseURL: serviceUrl,
          headers: uploadHeaders,
        }
      );

      const operationId = uploadResponse.data.operationId;
      this.logger?.debug(`Package uploaded. OperationId: ${operationId as string}`);

      this.logger?.verbose("Acquiring package ...");
      const acquireResponse = await this.axiosInstance.post(
        "/dev/v1/users/packages/acquisitions",
        {
          operationId: operationId,
        },
        {
          baseURL: serviceUrl,
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );

      const statusId = acquireResponse.data.statusId;
      this.logger?.debug(`Acquiring package with statusId: ${statusId as string} ...`);

      do {
        const statusResponse = await this.axiosInstance.get(
          `/dev/v1/users/packages/status/${statusId as string}`,
          {
            baseURL: serviceUrl,
            headers: { Authorization: `Bearer ${token}` },
          }
        );
        const resCode = statusResponse.status;
        this.logger?.debug(`Package status: ${resCode} ...`);
        if (resCode === 200) {
          const titleId: string = statusResponse.data.titleId;
          const appId: string = statusResponse.data.appId;
          this.logger?.info(`TitleId: ${titleId}`);
          this.logger?.info(`AppId: ${appId}`);
          this.logger?.verbose("Sideloading done.");
          return [titleId, appId];
        } else {
          await waitSeconds(2);
        }
      } while (true);
    } catch (error: any) {
      // this.logger?.error("Sideloading failed.");
      if (error.response) {
        // this.logger?.error(JSON.stringify(error.response.data));
        error = this.traceError(error);
      } else {
        // this.logger?.error(error.message);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async getShareLink(token: string, titleId: string): Promise<string> {
    const serviceUrl = await this.getTitleServiceUrl(token);
    try {
      const resp = await this.axiosInstance.get(
        `/marketplace/v1/users/titles/${titleId}/sharingInfo`,
        {
          baseURL: serviceUrl,
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );
      return resp.data.unifiedStoreLink;
    } catch (error: any) {
      if (error.response) {
        error = this.traceError(error);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async getLaunchInfoByManifestId(token: string, manifestId: string): Promise<any> {
    try {
      const serviceUrl = await this.getTitleServiceUrl(token);
      this.logger?.verbose(`Getting LaunchInfo with ManifestId ${manifestId} ...`);
      const launchInfo = await this.axiosInstance.post(
        "/catalog/v1/users/titles/launchInfo",
        {
          Id: manifestId,
          IdType: "ManifestId",
          Filter: {
            SupportedElementTypes: [
              // "Extensions", // Extensions require ClientDetails to be determined later
              "OfficeAddIns",
              "ExchangeAddIns",
              "FirstPartyPages",
              "Dynamics",
              "AAD",
              "LineOfBusiness",
              "StaticTabs",
              "ComposeExtensions",
              "Bots",
              "GraphConnector",
              "ConfigurableTabs",
              "Activities",
              "MeetingExtensionDefinition",
              "OpenAIPlugins",
              "Gpts",
              "DeclarativeCopilots",
              "Plugins",
            ],
          },
        },
        {
          baseURL: serviceUrl,
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );

      this.logger?.debug(JSON.stringify(launchInfo.data));
      return launchInfo.data;
    } catch (error: any) {
      this.logger?.error("Get LaunchInfo failed.");
      if (error.response) {
        if (error.response.status === 404) {
          throw new NotExtendedToM365Error(M365ErrorSource);
        }
        error = this.traceError(error);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async retrieveTitleId(token: string, manifestId: string): Promise<string> {
    const launchInfo = await this.getLaunchInfoByManifestId(token, manifestId);
    const titleId =
      (launchInfo.acquisition?.titleId?.id as string) ??
      (launchInfo.acquisition?.titleId as string);
    this.logger?.debug(`TitleId: ${titleId}`);
    return titleId;
  }

  public async retrieveAppId(token: string, manifestId: string): Promise<string | undefined> {
    const launchInfo = await this.getLaunchInfoByManifestId(token, manifestId);
    const appId = launchInfo.acquisition?.appId;
    this.logger?.debug(`AppId: ${appId as string}`);
    return appId;
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async unacquire(token: string, titleId: string): Promise<void> {
    try {
      const serviceUrl = await this.getTitleServiceUrl(token);
      this.logger?.verbose(`Unacquiring package with TitleId ${titleId} ...`);
      await this.axiosInstance.delete(`/catalog/v1/users/acquisitions/${titleId}`, {
        baseURL: serviceUrl,
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
      this.logger?.verbose("Unacquiring done.");
    } catch (error: any) {
      if (error.response) {
        error = this.traceError(error);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }
  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async getLaunchInfoByTitleId(token: string, titleId: string): Promise<unknown> {
    try {
      const serviceUrl = await this.getTitleServiceUrl(token);
      this.logger?.verbose(`Getting LaunchInfo with TitleId ${titleId} ...`);
      const launchInfo = await this.axiosInstance.get(
        `/catalog/v1/users/titles/${titleId}/launchInfo`,
        {
          baseURL: serviceUrl,
          params: {
            SupportedElementTypes:
              // eslint-disable-next-line no-secrets/no-secrets
              "Extensions,OfficeAddIns,ExchangeAddIns,FirstPartyPages,Dynamics,AAD,LineOfBusiness,StaticTabs,ComposeExtensions,Bots,GraphConnector,ConfigurableTabs,Activities,MeetingExtensionDefinition,OpenAIPlugins,Gpts,DeclarativeCopilots,Plugins",
          },
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );
      this.logger?.info(JSON.stringify(launchInfo.data));
      return launchInfo.data;
    } catch (error: any) {
      if (error.response) {
        error = this.traceError(error);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }

  @hooks([ErrorContextMW({ source: M365ErrorSource, component: M365ErrorComponent })])
  public async getActiveExperiences(
    token: string,
    ensureUpToDate = false
  ): Promise<string[] | undefined> {
    try {
      const serviceUrl = await this.getTitleServiceUrl(token);
      this.logger?.debug(`Get active experiences from service URL ${serviceUrl} ...`);
      // users/experiences is deprecating, using users/uitypes instead
      let response = await this.axiosInstance.get("/catalog/v1/users/uitypes", {
        baseURL: serviceUrl,
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
      let status = response.status;
      let activeExperiences = response.data?.activeExperiences as string[];
      const nextInterval = (response.data?.nextInterval as number) ?? -1;
      this.logger?.debug(`(${status}) Active experiences: ${JSON.stringify(activeExperiences)}`);

      // Short nextInterval means cache is refreshing
      if (ensureUpToDate && nextInterval > 0 && nextInterval < 10) {
        this.logger?.debug(`Active experiences is refreshing, wait for ${nextInterval} seconds.`);
        await waitSeconds(nextInterval);
        response = await this.axiosInstance.get("/catalog/v1/users/uitypes", {
          baseURL: serviceUrl,
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });
        status = response.status;
        activeExperiences = response.data?.activeExperiences as string[];
        this.logger?.debug(`(${status}) Active experiences: ${JSON.stringify(activeExperiences)}`);
      }

      return activeExperiences;
    } catch (error: any) {
      if (error.response) {
        error = this.traceError(error);
      }
      throw assembleError(error, M365ErrorSource);
    }
  }

  public async getCopilotStatus(
    token: string,
    ensureUpToDate = false
  ): Promise<boolean | undefined> {
    try {
      const activeExperiences = await this.getActiveExperiences(token, ensureUpToDate);
      const copilotAllowed =
        activeExperiences == undefined ? undefined : activeExperiences.includes("CopilotTeams");
      sendTelemetryEvent(Component.core, TelemetryEvent.CheckCopilot, {
        [TelemetryProperty.IsCopilotAllowed]: copilotAllowed?.toString() ?? "undefined",
      });
      return copilotAllowed;
    } catch (error: any) {
      sendTelemetryErrorEvent(
        Component.core,
        TelemetryEvent.CheckCopilot,
        new SystemError({
          error,
          source: M365ErrorSource,
          message: error.message ?? "Failed to get copilot status.",
          categories: [ErrorCategory.External],
        }),
        {
          [TelemetryProperty.CheckCopilotTracingId]: `${
            // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
            error.response?.headers?.traceresponse ??
            error.innerError?.response?.headers?.traceresponse ??
            ""
          }`,
        }
      );
      return undefined;
    }
  }

  private traceError(error: any): any {
    // add error details and trace to message
    const tracingId = (error.response.headers?.traceresponse ?? "") as string;
    const originalMessage = error.message as string;
    const innerError = error.response.data?.error || { code: "", message: "" };
    const finalMessage = `${originalMessage} (tracingId: ${tracingId}) ${
      innerError.code as string
    }: ${innerError.message as string} `;

    error.message = finalMessage;

    // HTTP 400 as user error due to invalid input
    if (error.response?.status === 400) {
      error = new UserError({
        name: "PackageServiceError",
        error,
        source: M365ErrorSource,
        message: finalMessage,
      });
    }

    return error;
  }

  private checkZip(path: string) {
    try {
      const zip = new AdmZip(path, {});
      zip.getEntries();
    } catch (error: any) {
      this.logger?.debug(`Invalid input zip ${path}. ${error.message as string}`);
      this.logger?.warning(`Please make sure input path is a valid app package zip. ${path}`);
    }
  }

  private getManifestFromZip(path: string): TeamsAppManifest | undefined {
    const zip = new AdmZip(path);
    const manifestEntry = zip.getEntry("manifest.json");
    if (!manifestEntry) {
      return undefined;
    }
    let manifestContent = manifestEntry.getData().toString("utf8");
    manifestContent = stripBom(manifestContent);
    return JSON.parse(manifestContent) as TeamsAppManifest;
  }
}
