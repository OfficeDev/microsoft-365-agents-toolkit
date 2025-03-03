// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Constants } from "./constants";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

export class AppStudioError {
  public static readonly FileNotFoundError = {
    name: "FileNotFoundError",
    message: (filePath: string): [string, string] => [
      getDefaultString("error.common.FileNotFoundError", filePath),
      getLocalizedString("error.common.FileNotFoundError", filePath),
    ],
  };

  public static readonly AuthServiceAPIFailedError = {
    name: "AuthServiceAPIFailed",
    message: (e: any, requestPath: string, apiName: string): [string, string] => [
      getDefaultString("error.appstudio.authServiceApiFailed", e.name, e.message, apiName),
      getLocalizedString("error.appstudio.authServiceApiFailed", e.name, e.message, requestPath),
    ],
  };

  public static readonly TeamsAppCreateFailedError = {
    name: "TeamsAppCreateFailed",
    message: (error?: any): [string, string] => [
      getDefaultString("error.appstudio.teamsAppCreateFailed", error?.message),
      getLocalizedString("error.appstudio.teamsAppCreateFailed", error?.displayMessage),
    ],
  };

  public static readonly TeamsAppUpdateFailedError = {
    name: "TeamsAppUpdateFailed",
    message: (teamsAppId: string, error: any): [string, string] => [
      getDefaultString("error.appstudio.teamsAppUpdateFailed", teamsAppId, error.message),
      getLocalizedString("error.appstudio.teamsAppUpdateFailed", teamsAppId, error.displayMessage),
    ],
  };

  public static readonly InvalidTeamsAppIdError = {
    name: "InvalidTeamsAppId",
    message: (teamsAppId: string): [string, string] => [
      getDefaultString("error.teamsApp.InvalidAppIdError", teamsAppId),
      getLocalizedString("error.teamsApp.InvalidAppIdError", teamsAppId),
    ],
  };

  public static readonly TeamsAppNotExistsError = {
    name: "TeasmAppNotExists",
    message: (teamsAppId: string): [string, string] => [
      getDefaultString("error.teamsApp.AppIdNotExistError", teamsAppId),
      getLocalizedString("error.teamsApp.AppIdNotExistError", teamsAppId),
    ],
  };

  public static readonly ValidationFailedError = {
    name: "ManifestValidationFailed",
    message: (errors: string[]): [string, string] => [
      getDefaultString("plugins.appstudio.validationFailedNotice") + " " + errors.join("\n"),
      getLocalizedString("plugins.appstudio.validationFailedNotice") + " " + errors.join("\n"),
    ],
  };

  public static readonly SyncManifestFailedError = {
    name: "SyncManifestFailed",
    message: (errors: string[]): [string, string] => [
      getDefaultString("plugins.appstudio.syncManifestFailedNotice") + " " + errors.join("\n"),
      getLocalizedString("plugins.appstudio.syncManifestFailedNotice") + " " + errors.join("\n"),
    ],
  };

  public static readonly TeamsAppPublishFailedError = {
    name: "TeamsAppPublishFailed",
    message: (teamsAppId: string, requestPath: string): [string, string] => [
      getDefaultString("error.appstudio.publishFailed", teamsAppId) +
        ` Request path: ${requestPath}`,
      getLocalizedString("error.appstudio.publishFailed", teamsAppId) +
        ` Request path: ${requestPath}`,
    ],
  };

  public static readonly TeamsPackageBuildError = {
    name: "TeamsPackageBuildError",
    message: (error: any): [string, string] => [
      error.message ?? getDefaultString("error.appstudio.buildError"),
      error.displayMessage ?? getLocalizedString("error.appstudio.buildError"),
    ],
  };

  public static readonly CheckPermissionFailedError = {
    name: "CheckPermissionFailed",
    message: (error: any): [string, string] => [
      getDefaultString("error.appstudio.checkPermissionFailed", error.message),
      getLocalizedString("error.appstudio.checkPermissionFailed", error.message),
    ],
  };

  public static readonly GrantPermissionFailedError = {
    name: "GrantPermissionFailed",
    message: (errorMessage: string, id?: string): [string, string] => [
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      `${Constants.PERMISSIONS.name}: ${id}. ` +
        getDefaultString("error.appstudio.grantPermissionFailed", errorMessage),
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      `${Constants.PERMISSIONS.name}: ${id}. ` +
        getLocalizedString("error.appstudio.grantPermissionFailed", errorMessage),
    ],
  };

  public static readonly ListCollaboratorFailedError = {
    name: "ListCollaboratorFailedError",
    message: (error: any): [string, string] => [
      getDefaultString("error.appstudio.listCollaboratorFailed", error.message),
      getLocalizedString("error.appstudio.listCollaboratorFailed", error.message),
    ],
  };

  public static readonly UpdateManifestWithInvalidAppError = {
    name: "UpdateManifestWithInvalidAppError",
    message: (appId: string): [string, string] => [
      getDefaultString("error.appstudio.updateManifestInvalidApp", appId),
      getLocalizedString("error.appstudio.updateManifestInvalidApp", appId),
    ],
  };

  public static readonly InvalidCapabilityError = {
    name: "InvalidCapabilityError",
    message: (capability: string): [string, string] => [
      getDefaultString("error.appstudio.invalidCapability", capability),
      getLocalizedString("error.appstudio.invalidCapability", capability),
    ],
  };

  public static readonly CapabilityExceedLimitError = {
    name: "CapabilityExceedLimitError",
    message: (capability: string): [string, string] => [
      getDefaultString("error.appstudio.capabilityExceedLimit", capability),
      getLocalizedString("error.appstudio.capabilityExceedLimit", capability),
    ],
  };

  public static readonly StaticTabNotExistError = {
    name: "StaticTabNotExist",
    message: (index: string): [string, string] => [
      getDefaultString("error.appstudio.staticTabNotExist", index),
      getLocalizedString("error.appstudio.staticTabNotExist", index),
    ],
  };

  public static readonly CapabilityNotExistError = {
    name: "CapabilityNotExist",
    message: (capability: string): [string, string] => [
      getDefaultString("error.appstudio.capabilityNotExist", capability),
      getLocalizedString("error.appstudio.capabilityNotExist", capability),
    ],
  };

  public static readonly TeamsAppCreateConflictError = {
    name: "TeamsAppCreateConflict",
    message: (): [string, string] => [
      getDefaultString("error.appstudio.teamsAppCreateConflict"),
      getLocalizedString("error.appstudio.teamsAppCreateConflict"),
    ],
  };

  public static readonly TeamsAppCreateConflictWithPublishedAppError = {
    name: "TeamsAppCreateConflictWithPublishedApp",
    message: (): [string, string] => [
      getDefaultString("error.appstudio.teamsAppCreateConflictWithPublishedApp"),
      getLocalizedString("error.appstudio.teamsAppCreateConflictWithPublishedApp"),
    ],
  };

  public static readonly TeamsAppPublishConflictError = {
    name: "TeamsAppPublishConflict",
    message: (teamsAppId: string): [string, string] => [
      getDefaultString("error.appstudio.teamsAppPublishConflict", teamsAppId),
      getLocalizedString("error.appstudio.teamsAppPublishConflict", teamsAppId),
    ],
  };

  public static readonly TeamsAppRequiredPropertyMissingError = {
    name: "TeamsAppMissingRequiredCapability",
    message: (property: string, path: string): [string, string] => [
      getDefaultString("error.appstudio.teamsAppRequiredPropertyMissing", property, path),
      getLocalizedString("error.appstudio.teamsAppRequiredPropertyMissing", property, path),
    ],
  };
}
