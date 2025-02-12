// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Huajie Zhang <huajiezhang@microsoft.com>
 */
import { OptionItem } from "@microsoft/teamsfx-api";
import path from "path";
import { getLocalizedString } from "../common/localizeUtils";

export enum SolutionError {
  NoAppStudioToken = "NoAppStudioToken",
  FailedToRetrieveUserInfo = "FailedToRetrieveUserInfo",
  CannotFindUserInCurrentTenant = "CannotFindUserInCurrentTenant",
  EmailCannotBeEmptyOrSame = "EmailCannotBeEmptyOrSame",
  InvalidProjectPath = "InvalidProjectPath",
  FailedToCreateAuthFiles = "FailedToCreateAuthFiles",
  FailedToLoadDotEnvFile = "FailedToLoadDotEnvFile",
  InvalidManifestError = "InvalidManifestError",
  FailedToLoadManifestFile = "FailedToLoadManifestFile",
}

export const ViewAadAppHelpLinkV5 = "https://aka.ms/teamsfx-view-aad-app-v5";

// This is the max length specified in
// https://developer.microsoft.com/en-us/json-schemas/teams/v1.7/MicrosoftTeams.schema.json

export enum SolutionTelemetryEvent {
  AddSsoStart = "add-sso-start",
  AddSso = "add-sso",
}

export enum SolutionTelemetryProperty {
  Component = "component",
  Success = "success",
  CollaboratorCount = "collaborator-count",
  AadOwnerCount = "aad-owner-count",
  AadPermission = "aad-permission",
  ArmDeploymentError = "arm-deployment-error",
  TeamsAppPermission = "teams-app-permission",
  Env = "env",
  SubscriptionId = "subscription-id",
  M365TenantId = "m365-tenant-id",
  PreviousM365TenantId = "previous-m365-tenant-id",
  ConfirmRes = "confirm-res",
}

export enum SolutionTelemetrySuccess {
  Yes = "yes",
  No = "no",
}

export const SolutionTelemetryComponentName = "core";
export const SolutionSource = "core";
export const CoordinatorSource = "core";

export class AddSsoParameters {
  static readonly filePath = path.join("plugins", "resource", "aad", "auth");
  static readonly Bot = "bot";
  static readonly Tab = "tab";
  static readonly V3 = "V3";
  static readonly V3AuthFolder = "TeamsFx-Auth";
  static readonly Readme = "README.md";
  static readonly ReadmeCSharp = "README.txt";
  static readonly LearnMore = () => getLocalizedString("core.provision.learnMore");
  static readonly LearnMoreUrl = "https://aka.ms/teamsfx-add-sso-readme";
  static readonly AddSso = "addSso";
  static readonly AppSettings = "appsettings.json";
  static readonly AppSettingsDev = "appsettings.Development.json";
  static readonly AppSettingsToAdd = {
    Authentication: {
      ClientId: "$clientId$",
      ClientSecret: "$client-secret$",
      OAuthAuthority: "$oauthAuthority$",
    },
  };
  static readonly AppSettingsToAddForBot = {
    Authentication: {
      ClientId: "$clientId$",
      ClientSecret: "$client-secret$",
      OAuthAuthority: "$oauthAuthority$",
      ApplicationIdUri: "$applicationIdUri$",
      Bot: {
        InitiateLoginEndpoint: "$initiateLoginEndpoint$",
      },
    },
  };
}

export const SingleSignOnOptionItem: OptionItem = {
  id: "sso",
  label: `$(unlock) ${getLocalizedString("core.SingleSignOnOption.label")}`,
  detail: getLocalizedString("core.SingleSignOnOption.detail"),
  groupName: getLocalizedString("core.options.separator.additional"),
  data: "https://aka.ms/teamsfx-add-sso",
  buttons: [
    {
      iconPath: "tasklist",
      tooltip: getLocalizedString("core.option.tutorial"),
      command: "fx-extension.openTutorial",
    },
  ],
};

export enum BotScenario {
  NotificationBot = "notificationBot",
  CommandAndResponseBot = "commandAndResponseBot",
  WorkflowBot = "workflowBot",
}

export const AadConstants = {
  DefaultTemplateFileName: "aad.manifest.json",
};

export const KiotaLastCommands = {
  createPluginWithManifest: "createPluginWithManifest",
  createDeclarativeCopilotWithManifest: "createDeclarativeCopilotWithManifest",
  addPlugin: "addPlugin",
};

export const OpenAIEnvironmentVariables = {
  OPENAI_API_KEY: "OPENAI_API_KEY",
  AZURE_OPENAI_API_KEY: "AZURE_OPENAI_API_KEY",
  SECRET_OPENAI_API_KEY: "SECRET_OPENAI_API_KEY",
  SECRET_AZURE_OPENAI_API_KEY: "SECRET_AZURE_OPENAI_API_KEY",
  AZURE_OPENAI_ENDPOINT: "AZURE_OPENAI_ENDPOINT",
  AZURE_OPENAI_DEPLOYMENT_NAME: "AZURE_OPENAI_DEPLOYMENT_NAME",
};
