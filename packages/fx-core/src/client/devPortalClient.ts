// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { AxiosInstance, AxiosResponse } from "axios";
import FormData from "form-data";
import { getResourceServiceEndpoint, ResourceServiceType } from "../common/constants";
import { ErrorContextMW, TOOLS } from "../common/globalVars";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { RetryHandler } from "../common/retryHandler";
import { WrappedAxiosClient } from "../common/wrappedAxiosClient";
import {
  APP_STUDIO_API_NAMES,
  Constants,
  ErrorMessages,
} from "../component/driver/teamsApp/constants";
import { AsyncAppValidationDetailsResponse } from "../component/driver/teamsApp/interfaces/AsyncAppValidationDetailsResponse";
import { AsyncAppValidationResponse } from "../component/driver/teamsApp/interfaces/AsyncAppValidationResponse";
import { AsyncAppValidationResultsResponse } from "../component/driver/teamsApp/interfaces/AsyncAppValidationResultsResponse";
import { IValidationResult } from "../component/driver/teamsApp/interfaces/appdefinitions/IValidationResult";
import { AppUser } from "../component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { IBotRegistration } from "../component/resource/botService/appStudio/interfaces/IBotRegistration";
import { isHappyResponse } from "../component/resource/botService/common";
import { ConfigUpdatingError, ProvisionError } from "../component/resource/botService/errors";
import { Messages } from "../component/resource/botService/messages";
import { CommonStrings, ConfigNames } from "../component/resource/botService/strings";
import { TeamsDevPortalClient } from "./teamsDevPortalClient";
import {
  AppResponse,
  AppSummary,
  AppUserPayload,
  isAppSummary,
  hasValidAppResponse,
} from "./devPortalModels";

interface PagedResponse<T> {
  items: T[];
  continuationToken?: string | null;
}

const newDeveloperPortalEndpoints: Record<string, string> = {
  apac: "https://dev.teams.microsoft.com/cosmicprodapac",
  amer: "https://dev.teams.microsoft.com/cosmicprodamer",
  emea: "https://dev.teams.microsoft.com/cosmicprodemea",
};

export class DevPortalClient {
  regionEndpoint?: string;
  // Reuse the unchanged transport and non-app operations, with this client's endpoint and
  // dispatch. Native app methods are deliberately not overrides of the legacy contract.
  setRegionEndpoint = TeamsDevPortalClient.prototype.setRegionEndpoint;
  getEndpoint = TeamsDevPortalClient.prototype.getEndpoint;
  createRequesterWithToken = TeamsDevPortalClient.prototype.createRequesterWithToken;
  wrapException = TeamsDevPortalClient.prototype.wrapException;
  wrapResponse = TeamsDevPortalClient.prototype.wrapResponse;
  handleBotFrameworkError = TeamsDevPortalClient.prototype.handleBotFrameworkError;
  importApp = TeamsDevPortalClient.prototype.importApp;
  publishTeamsApp = TeamsDevPortalClient.prototype.publishTeamsApp;
  publishTeamsAppUpdate = TeamsDevPortalClient.prototype.publishTeamsAppUpdate;
  getStaggedApp = TeamsDevPortalClient.prototype.getStaggedApp;
  getSideloadingStatus = TeamsDevPortalClient.prototype.getSideloadingStatus;
  createApiKeyRegistration = TeamsDevPortalClient.prototype.createApiKeyRegistration;
  getApiKeyRegistrationById = TeamsDevPortalClient.prototype.getApiKeyRegistrationById;
  updateApiKeyRegistration = TeamsDevPortalClient.prototype.updateApiKeyRegistration;
  getOauthRegistrationById = TeamsDevPortalClient.prototype.getOauthRegistrationById;
  createOauthRegistration = TeamsDevPortalClient.prototype.createOauthRegistration;
  updateOauthRegistration = TeamsDevPortalClient.prototype.updateOauthRegistration;
  updateMessageEndpoint = TeamsDevPortalClient.prototype.updateMessageEndpoint;
  createAADApp = TeamsDevPortalClient.prototype.createAADApp;

  private readonly appEtags = new Map<string, string>();

  async setRegionEndpointByToken(authSvcToken: string): Promise<void> {
    if (
      getResourceServiceEndpoint(ResourceServiceType.TDP) === "https://dev-int.teams.microsoft.com"
    ) {
      return;
    }
    const requester = WrappedAxiosClient.create({
      baseURL: getResourceServiceEndpoint(ResourceServiceType.AuthSvc),
    });
    requester.defaults.headers.common["Authorization"] = `Bearer ${authSvcToken}`;
    requester.defaults.headers.common["Client-Source"] = "teamstoolkit";
    const response = await RetryHandler.Retry(() => requester.post("/v1.0/users/region"));
    const regionGtms = response?.data?.regionGtms;
    this.regionEndpoint = this.resolveRegionEndpoint(
      regionGtms?.teamsDeveloperPortal,
      regionGtms?.teamsDevPortal
    );
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async createApp(token: string, name: string): Promise<AppResponse> {
    const requester = this.createRequesterWithToken(token);
    try {
      const response = await RetryHandler.Retry(() =>
        requester.post<AppResponse>("/v1.0/apps", {
          appDetails: {
            applicationManifest: {
              manifestVersion: "1.15",
              name: { short: name },
              icons: {
                color: "default-app-icons/images/color.png",
                outline: "default-app-icons/images/outline.png",
              },
            },
          },
        })
      );
      if (!response || !hasValidAppResponse(response.data)) {
        throw new Error("Cannot create Teams app");
      }
      return response.data;
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.CREATE_APP);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async updateApp(token: string, appId: string, file: Buffer): Promise<AppResponse> {
    try {
      const requester = this.createRequesterWithToken(token);
      const response = await RetryHandler.Retry(() => {
        const content = this.createAppPackageForm(file);
        return requester.put<AppResponse>(`/v1.0/apps/${appId}/apppackage`, content, {
          headers: content.getHeaders(),
        });
      });
      if (!response || !hasValidAppResponse(response.data) || response.data.appId !== appId) {
        throw new Error(`Cannot update the app with app ID ${appId}`);
      }
      return response.data;
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.UPDATE_APP);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async listApps(token: string): Promise<AppSummary[]> {
    try {
      const requester = this.createRequesterWithToken(token);
      TOOLS.logProvider.debug(`Sent API Request: GET ${this.getEndpoint()}/v1.0/apps`);
      const apps: AppSummary[] = [];
      let continuationToken: string | null | undefined;
      do {
        const response = await RetryHandler.Retry(() =>
          requester.get<PagedResponse<unknown>>("/v1.0/apps", {
            params: { pageSize: 100 },
            headers: continuationToken ? { "x-ms-continuation": continuationToken } : undefined,
          })
        );
        const page = response?.data;
        if (
          !Array.isArray(page?.items) ||
          !page.items.every(isAppSummary) ||
          (page.continuationToken != null && typeof page.continuationToken !== "string")
        ) {
          throw new Error("Cannot get the app definitions");
        }
        apps.push(...page.items);
        continuationToken = page.continuationToken;
      } while (continuationToken);
      return apps;
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.LIST_APPS);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async deleteApp(token: string, appId: string): Promise<boolean> {
    try {
      const requester = this.createRequesterWithToken(token);
      TOOLS.logProvider.debug(`Sent API Request: DELETE ${this.getEndpoint()}/v1.0/apps/${appId}`);
      const response = await RetryHandler.Retry(() => requester.delete(`/v1.0/apps/${appId}`));
      if (response?.status === 204) {
        return true;
      }
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.DELETE_APP);
    }
    throw this.wrapException(
      new Error("cannot delete the app: " + appId),
      APP_STUDIO_API_NAMES.DELETE_APP
    );
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async getApp(token: string, appId: string): Promise<AppResponse> {
    try {
      const response = await this.getAppResponse(token, appId);
      if (hasValidAppResponse(response.data) && response.data.appId === appId) {
        const etag = response.headers?.etag;
        if (typeof etag === "string" && etag) {
          this.appEtags.set(appId, etag);
        }
        return response.data;
      }
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.GET_APP);
    }
    throw this.wrapException(
      new Error(`cannot get the app definition with app ID ${appId}`),
      APP_STUDIO_API_NAMES.GET_APP
    );
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async getAppPackage(token: string, appId: string): Promise<Buffer> {
    TOOLS.logProvider?.info("Downloading app package for app " + appId);
    const requester = this.createRequesterWithToken(token);
    try {
      const response = await RetryHandler.Retry(() =>
        requester.get<Buffer>(`/v1.0/apps/${appId}/appPackage`, { responseType: "arraybuffer" })
      );
      if (response?.data) {
        TOOLS.logProvider?.info("Download app package successfully");
        return response.data;
      }
      throw this.wrapException(
        new Error(getLocalizedString("plugins.appstudio.emptyAppPackage", appId)),
        APP_STUDIO_API_NAMES.GET_APP_PACKAGE
      );
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.GET_APP_PACKAGE);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async checkExistsInTenant(token: string, appId: string): Promise<boolean> {
    try {
      await this.getApp(token, appId);
      return true;
    } catch {
      return false;
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async removePermission(token: string, appId: string, userToRemove: AppUser): Promise<void> {
    const app = await this.getApp(token, appId);
    if (!this.checkUser(app, userToRemove)) return;
    const updatedUsers = (app.appProfile?.appAccessControl?.appUsers ?? []).filter(
      (user) => user.aadId !== userToRemove.aadId
    );
    await this.updateOwners(token, app.appId, updatedUsers);
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async grantPermission(token: string, appId: string, newUser: AppUser): Promise<void> {
    const app = await this.getApp(token, appId);
    if (this.checkUser(app, newUser)) return;
    const updatedUsers = [
      ...(app.appProfile?.appAccessControl?.appUsers ?? []),
      this.toAppUserPayload(newUser),
    ];
    await this.updateOwners(token, app.appId, updatedUsers);
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async partnerCenterAppPackageValidation(
    token: string,
    file: Buffer,
    signal?: AbortSignal
  ): Promise<IValidationResult> {
    const requester = this.createRequesterWithToken(token);
    try {
      const response = await RetryHandler.Retry(() => {
        const content = new FormData();
        content.append("appPackage", file, {
          filename: "appPackage.zip",
          contentType: "application/zip",
        });
        return requester.post("/v1.0/appvalidation/apppackage/validate", content, {
          headers: content.getHeaders(),
          signal,
        });
      });
      return response?.data;
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.VALIDATE_APP_PACKAGE);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async submitAppValidationRequest(
    token: string,
    appId: string
  ): Promise<AsyncAppValidationResponse> {
    try {
      const requester = this.createRequesterWithToken(token);
      const response = await this.submitAppValidation(requester, appId);
      return response.data;
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.SUBMIT_APP_VALIDATION);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async getAppValidationRequestList(
    token: string,
    appId: string
  ): Promise<AsyncAppValidationDetailsResponse> {
    try {
      const requester = this.createRequesterWithToken(token);
      return await this.getAppValidationRequests(requester, appId);
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.GET_APP_VALIDATION_REQUESTS);
    }
  }

  private async getAppValidationRequests(
    requester: AxiosInstance,
    appId: string
  ): Promise<AsyncAppValidationDetailsResponse> {
    const items: Record<string, any>[] = [];
    let continuationToken: string | null | undefined;
    do {
      const response = await RetryHandler.Retry(() =>
        requester.get(`/v1.0/appValidations/apps/${appId}`, {
          params: { pageSize: 100 },
          headers: continuationToken ? { "x-ms-continuation": continuationToken } : undefined,
        })
      );
      const page = response?.data as PagedResponse<Record<string, any>> | undefined;
      if (!page?.items) throw new Error("Cannot get app validation requests");
      items.push(...page.items);
      continuationToken =
        page.continuationToken ??
        (response?.headers?.["x-continuation-token"] as string | undefined);
    } while (continuationToken);
    return {
      appValidations: items.map((item) => ({
        id: item.appValidationId,
        appId: item.appId,
        appVersion: item.appVersion,
        manifestVersion: item.manifestVersion,
        status: item.status,
        createdAt: item.submittedDate,
        updatedAt: item.completedDate,
      })),
    };
  }

  private async submitAppValidation(
    requester: AxiosInstance,
    appId: string
  ): Promise<AxiosResponse<AsyncAppValidationResponse>> {
    const response = await RetryHandler.Retry(() =>
      requester.post("/v1.0/appvalidation/validate", {
        appId,
        appEnvironmentId: null,
        testSuites: null,
      })
    );
    if (!response) {
      throw new Error("App validation request returned no response.");
    }
    return response;
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async getAppValidationById(
    token: string,
    appValidationId: string
  ): Promise<AsyncAppValidationResultsResponse> {
    try {
      const requester = this.createRequesterWithToken(token);
      const response = await RetryHandler.Retry(() =>
        requester.get(`/v1.0/appValidations/${appValidationId}`)
      );
      return response?.data as AsyncAppValidationResultsResponse;
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.GET_APP_VALIDATION_RESULT);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async getBotRegistration(token: string, botId: string): Promise<IBotRegistration | undefined> {
    return this.getBotRegistrationAtPath(token, `/v1.0/botregistrations/${botId}`);
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async listBots(token: string): Promise<IBotRegistration[] | undefined> {
    const requester = this.createRequesterWithToken(token);
    try {
      const response = await RetryHandler.Retry(() => requester.get("/v1.0/botregistrations"));
      if (isHappyResponse(response)) return response!.data as IBotRegistration[];
      throw this.wrapException(
        this.wrapResponse(undefined, response),
        APP_STUDIO_API_NAMES.LIST_BOT,
        getDefaultString("error.appstudio.apiFailed.name.common"),
        "Failed to get data"
      );
    } catch (error) {
      this.handleBotFrameworkError(error, APP_STUDIO_API_NAMES.LIST_BOT);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async deleteBot(token: string, botId: string): Promise<void> {
    try {
      const requester = this.createRequesterWithToken(token);
      await RetryHandler.Retry(() => requester.delete(`/v1.0/botregistrations/${botId}`));
    } catch (error) {
      this.handleBotFrameworkError(error, APP_STUDIO_API_NAMES.DELETE_BOT);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async createBotRegistration(
    token: string,
    registration: IBotRegistration,
    checkExistence = true
  ): Promise<void> {
    if (registration.botId && checkExistence) {
      const botRegistration = await this.getBotRegistration(token, registration.botId);
      if (botRegistration) {
        TOOLS.logProvider.info(Messages.BotResourceExist("Appstudio"));
        return;
      }
    }
    try {
      const requester = this.createRequesterWithToken(token);
      const response = await RetryHandler.Retry(() =>
        requester.post("/v1.0/botregistrations", registration)
      );
      if (!isHappyResponse(response)) {
        throw new ProvisionError(CommonStrings.APP_STUDIO_BOT_REGISTRATION);
      }
    } catch (error) {
      this.handleBotFrameworkError(error, APP_STUDIO_API_NAMES.CREATE_BOT);
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "TeamsDevPortalClient" })])
  async updateBotRegistration(token: string, botReg: IBotRegistration): Promise<void> {
    try {
      const requester = this.createRequesterWithToken(token);
      const response = await RetryHandler.Retry(() =>
        requester.put(`/v1.0/botregistrations/${botReg.botId!}`, botReg)
      );
      if (!isHappyResponse(response)) {
        throw new ConfigUpdatingError(ConfigNames.MESSAGE_ENDPOINT);
      }
    } catch (error) {
      this.handleBotFrameworkError(error, APP_STUDIO_API_NAMES.UPDATE_BOT);
    }
  }

  private resolveRegionEndpoint(teamsDeveloperPortal?: string, teamsDevPortal?: string): string {
    if (teamsDeveloperPortal?.startsWith("https://")) return teamsDeveloperPortal;
    const oldEndpointRegion = Object.keys(newDeveloperPortalEndpoints).find((region) =>
      teamsDevPortal?.match(new RegExp(`/${region}(?:/api)?/?$`, "i"))
    );
    return newDeveloperPortalEndpoints[oldEndpointRegion ?? "amer"];
  }

  private createAppPackageForm(file: Buffer): FormData {
    const content = new FormData();
    content.append("AppPackageZip", file, {
      filename: "appPackage.zip",
      contentType: "application/zip",
    });
    return content;
  }

  private async getAppResponse(token: string, appId: string): Promise<AxiosResponse<AppResponse>> {
    const requester = this.createRequesterWithToken(token);
    TOOLS.logProvider.debug(`Sent API Request: GET ${this.getEndpoint()}/v1.0/apps/${appId}`);
    const response = await RetryHandler.Retry(() =>
      requester.get<AppResponse>(`/v1.0/apps/${appId}`)
    );
    if (!response) throw new Error(`Cannot get the app with app ID ${appId}`);
    return response;
  }

  private async updateOwners(
    token: string,
    appId: string,
    appUsers: AppUserPayload[]
  ): Promise<void> {
    try {
      TOOLS.logProvider.debug(
        getLocalizedString(
          "core.common.SendingApiRequest",
          `${this.getEndpoint()}/v1.0/apps/{appId}/owners`,
          JSON.stringify({ appUsers })
        )
      );
      const requester = this.createRequesterWithToken(token);
      const response = await RetryHandler.Retry(() =>
        requester.post(
          `/v1.0/apps/${appId}/owners`,
          { appUsers },
          { headers: { "If-Match": this.appEtags.get(appId) ?? "*" } }
        )
      );
      TOOLS.logProvider.debug(
        getLocalizedString("core.common.ReceiveApiResponse", JSON.stringify(response?.data))
      );
      if (!response) throw new Error(ErrorMessages.GrantPermissionFailed);
    } catch (error) {
      throw this.wrapException(error, APP_STUDIO_API_NAMES.UPDATE_OWNER);
    }
  }

  private async getBotRegistrationAtPath(
    token: string,
    path: string
  ): Promise<IBotRegistration | undefined> {
    const requester = this.createRequesterWithToken(token);
    try {
      const response = await RetryHandler.Retry(() => requester.get(path));
      if (isHappyResponse(response)) return response!.data as IBotRegistration;
      throw this.wrapException(
        this.wrapResponse(undefined, response),
        APP_STUDIO_API_NAMES.GET_BOT,
        getDefaultString("error.appstudio.apiFailed.name.common"),
        "Failed to get data"
      );
    } catch (error) {
      this.handleBotFrameworkError(error, APP_STUDIO_API_NAMES.GET_BOT);
    }
  }

  async getBotId(token: string, appId: string): Promise<string | undefined> {
    const app = await this.getApp(token, appId);
    return app.appProfile?.appDetails?.applicationManifest?.bots?.[0]?.botId ?? undefined;
  }

  async getUserList(token: string, appId: string): Promise<AppUserPayload[] | undefined> {
    const app = await this.getApp(token, appId);
    return app.appProfile?.appAccessControl?.appUsers ?? undefined;
  }

  async checkPermission(token: string, appId: string, userObjectId: string): Promise<string> {
    try {
      const users = await this.getUserList(token, appId);
      const user = users?.find((user) => user.aadId === userObjectId);
      if (user) {
        return user.role === "Owner"
          ? Constants.PERMISSIONS.admin
          : Constants.PERMISSIONS.operative;
      }
    } catch {
      return Constants.PERMISSIONS.noPermission;
    }
    return Constants.PERMISSIONS.noPermission;
  }

  checkUser(app: AppResponse, user: AppUser): boolean {
    return (
      app.appProfile?.appAccessControl?.appUsers?.some((item) => item.aadId === user.aadId) ?? false
    );
  }

  private toAppUserPayload(user: AppUser): AppUserPayload {
    return {
      aadId: user.aadId,
      tenantId: user.tenantId,
      role: user.isAdministrator ? "Owner" : "Member",
    };
  }
}

export const devPortalClient = new DevPortalClient();
