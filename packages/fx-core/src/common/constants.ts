// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { getLocalizedString } from "./localizeUtils";

export class ConstantString {
  static readonly UTF8Encoding = "utf-8";
  static readonly DeploymentResourceType = "Microsoft.Resources/deployments";
  static readonly DeploymentNotFound = "DeploymentNotFound";
  static readonly RootFolder = "TeamsApps";
}

export class HelpLinks {
  static readonly WhyNeedProvision = "https://aka.ms/teamsfx/whyneedprovision";
  static readonly SwitchTenant = "https://aka.ms/teamsfx-switch-tenant";
}

export class VSCodeExtensionCommand {
  static readonly showOutputChannel = "command:fx-extension.showOutputChannel";
  static readonly openFolder = "command:fx-extension.openFolder";
  static readonly openReadme = "command:fx-extension.openReadMe?%5B%22Notification%22%5D";
}

export class TeamsClientId {
  static readonly MobileDesktop = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
  static readonly Web = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
}

export class OfficeClientId {
  static readonly Desktop = "0ec893e0-5785-4de6-99da-4ed124e5296c";
  static readonly Web1 = "4345a7b9-9a63-4910-a426-35363201d503";
  static readonly Web2 = "4765445b-32c6-49b0-83e6-1d93765276ca";
}

export class OutlookClientId {
  static readonly Desktop = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
  static readonly Web1 = "00000002-0000-0ff1-ce00-000000000000";
  static readonly Web2 = "bc59ab01-8403-45c6-8796-ac3ef710b3e3";
  static readonly Mobile = "27922004-5251-4030-b22d-91ecd9a37ea4";
}

export function getAllowedAppMaps(): Record<string, string> {
  return {
    [TeamsClientId.MobileDesktop]: getLocalizedString("core.common.TeamsMobileDesktopClientName"),
    [TeamsClientId.Web]: getLocalizedString("core.common.TeamsWebClientName"),
    [OfficeClientId.Desktop]: getLocalizedString("core.common.OfficeDesktopClientName"),
    [OfficeClientId.Web1]: getLocalizedString("core.common.OfficeWebClientName1"),
    [OfficeClientId.Web2]: getLocalizedString("core.common.OfficeWebClientName2"),
    [OutlookClientId.Desktop]: getLocalizedString("core.common.OutlookDesktopClientName"),
    [OutlookClientId.Web1]: getLocalizedString("core.common.OutlookWebClientName1"),
    [OutlookClientId.Web2]: getLocalizedString("core.common.OutlookWebClientName2"),
  };
}

const AzurePortalUrl = "https://portal.azure.com";
export function getResourceGroupInPortal(
  subscriptionId?: string,
  tenantId?: string,
  resourceGroupName?: string
): string | undefined {
  if (subscriptionId && tenantId && resourceGroupName) {
    return `${AzurePortalUrl}/#@${tenantId}/resource/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}`;
  } else {
    return undefined;
  }
}
export function getAppStudioEndpoint(): string {
  if (process.env.APP_STUDIO_ENV && process.env.APP_STUDIO_ENV === "int") {
    return "https://dev-int.teams.microsoft.com";
  } else {
    return "https://dev.teams.microsoft.com";
  }
}

export const AuthSvcScopes = ["https://authsvc.teams.microsoft.com/Region.ReadWrite"];
export const GraphScopes = ["Application.ReadWrite.All", "TeamsAppInstallation.ReadForUser"];
export const GCScopes = ["ExternalConnection.Read.All"];
export const GraphReadUserScopes = ["https://graph.microsoft.com/User.ReadBasic.All"];
export const SPFxScopes = (tenant: string) => [`${tenant}/Sites.FullControl.All`];
export const AzureScopes = ["https://management.core.windows.net/user_impersonation"];
export const AppStudioScopes = [`${getAppStudioEndpoint()}/AppDefinitions.ReadWrite`];
export const SpecParserSource = "SpecParser";
