// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  err,
  FxError,
  LogProvider,
  M365TokenProvider,
  ManifestProperties,
  ok,
  Result,
} from "@microsoft/teamsfx-api";

import { hooks } from "@feathersjs/hooks";
import { AppStudioScopes } from "../../common/constants";
import { ErrorContextMW } from "../../common/globalVars";
import { CoreSource } from "../../error";
import { assembleError } from "../../error/common";
import { HubTypes } from "../../question/constants";
import { NotExtendedToM365Error } from "./errors";
import { PackageService } from "./packageService";
import { MosServiceEndpoint, MosServiceScope } from "./serviceConstant";
import { officeBaseUrl, outlookBaseUrl, outlookCopilotAppId } from "./constants";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";

export class LaunchHelper {
  private readonly m365TokenProvider: M365TokenProvider;
  private readonly logger?: LogProvider;

  public constructor(m365TokenProvider: M365TokenProvider, logger?: LogProvider) {
    this.m365TokenProvider = m365TokenProvider;
    this.logger = logger;
  }
  @hooks([ErrorContextMW({ component: "LaunchHelper" })])
  public async getLaunchUrl(
    hub: HubTypes,
    teamsAppId: string,
    properties: ManifestProperties,
    withLoginHint = true
  ): Promise<Result<string, FxError>> {
    const capabilities = properties.capabilities;
    const loginHint = withLoginHint
      ? (await this.getUpnFromToken()) ?? "login_your_m365_account" // a workaround that user has the chance to login
      : undefined;
    let url: URL;
    const copilotCapabilities = ["plugin", "copilotGpt"];
    const hasCopilotExtensionOnly =
      capabilities.length > 0 &&
      capabilities.filter((capability: string) => !copilotCapabilities.includes(capability))
        .length === 0;
    switch (hub) {
      case HubTypes.teams: {
        let installAppPackage = true;
        if (
          capabilities.length > 0 &&
          (hasCopilotExtensionOnly ||
            (!capabilities.includes("staticTab") &&
              !capabilities.includes("Bot") &&
              !capabilities.includes("configurableTab") &&
              properties.isApiMeAAD))
        ) {
          installAppPackage = false;
        }
        const baseUrl = installAppPackage
          ? `https://teams.microsoft.com/l/app/${teamsAppId}?installAppPackage=true&webjoin=true`
          : "https://teams.microsoft.com";
        url = new URL(baseUrl);
        const tid = await this.getTidFromToken();
        if (tid) {
          url.searchParams.append("tenantId", tid);
          url.searchParams.append("appTenantId", tid);
        }
        break;
      }
      case HubTypes.outlook: {
        const result = await this.getM365AppId(teamsAppId);
        if (result.isErr()) {
          return err(result.error);
        }
        const baseUrl = hasCopilotExtensionOnly
          ? `${outlookBaseUrl}/host/${outlookCopilotAppId}`
          : capabilities.includes("staticTab")
          ? `${outlookBaseUrl}/host/${result.value}`
          : `${outlookBaseUrl}/mail`;
        url = new URL(baseUrl);
        break;
      }
      case HubTypes.office:
        {
          const result = await this.getM365AppId(teamsAppId);
          if (result.isErr()) {
            return err(result.error);
          }
          const baseUrl = hasCopilotExtensionOnly
            ? `${officeBaseUrl}/chat?auth=2`
            : `${officeBaseUrl}/m365apps/${result.value}?auth=2`;
          url = new URL(baseUrl);
        }
        break;
    }
    if (loginHint) {
      url.searchParams.append("login_hint", loginHint);
    }
    return ok(url.toString());
  }

  public async getM365AppId(teamsAppId: string): Promise<Result<string, FxError>> {
    const sideloadingServiceEndpoint =
      process.env.SIDELOADING_SERVICE_ENDPOINT ?? MosServiceEndpoint;
    const sideloadingServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const packageService = new PackageService(sideloadingServiceEndpoint, this.logger);

    const sideloadingTokenRes = await this.m365TokenProvider.getAccessToken({
      scopes: [sideloadingServiceScope],
    });
    if (sideloadingTokenRes.isErr()) {
      return err(sideloadingTokenRes.error);
    }
    const sideloadingToken = sideloadingTokenRes.value;

    try {
      const m365AppId = await packageService.retrieveAppId(sideloadingToken, teamsAppId);
      if (!m365AppId) {
        return err(new NotExtendedToM365Error(CoreSource));
      }
      return ok(m365AppId);
    } catch (error) {
      return err(assembleError(error));
    }
  }

  private async getTidFromToken(): Promise<string | undefined> {
    try {
      const statusRes = await this.m365TokenProvider.getStatus({ scopes: AppStudioScopes });
      const tokenObject = statusRes.isOk() ? statusRes.value.accountInfo : undefined;
      return tokenObject?.tid as string;
    } catch {
      return undefined;
    }
  }

  private async getUpnFromToken(): Promise<string | undefined> {
    try {
      const statusRes = await this.m365TokenProvider.getStatus({ scopes: AppStudioScopes });
      const tokenObject = statusRes.isOk() ? statusRes.value.accountInfo : undefined;
      return tokenObject?.upn as string;
    } catch {
      return undefined;
    }
  }
}
