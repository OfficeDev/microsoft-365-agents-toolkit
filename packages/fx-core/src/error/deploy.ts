// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { BlobDeleteResponse, BlobUploadCommonResponse } from "@azure/storage-blob";
import { ErrorCategory } from "./types";

export class DeployEmptyFolderError extends UserError {
  constructor(folderPath: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString("error.deploy.DeployEmptyFolderError", folderPath),
      displayMessage: getLocalizedString("error.deploy.DeployEmptyFolderError", folderPath),
      categories: [ErrorCategory.Internal],
    });
  }
}

export class CheckDeploymentStatusTimeoutError extends UserError {
  constructor(helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString("error.deploy.CheckDeploymentStatusTimeoutError"),
      displayMessage: getLocalizedString("error.deploy.CheckDeploymentStatusTimeoutError"),
      helpLink: helpLink,
      categories: [ErrorCategory.External],
    });
  }
}

export class ZipFileError extends UserError {
  constructor(error: Error) {
    super({
      source: "azureDeploy",
      message: getDefaultString("error.deploy.ZipFileError"),
      displayMessage: getLocalizedString("error.deploy.ZipFileError"),
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class CacheFileInUse extends UserError {
  constructor(path: string, error: Error) {
    super({
      source: "azureDeploy",
      message: getDefaultString("error.deploy.ZipFileTargetInUse", path),
      displayMessage: getLocalizedString("error.deploy.ZipFileTargetInUse", path),
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class GetPublishingCredentialsError extends UserError {
  constructor(appName: string, resourceGroup: string, error: Error, helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.GetPublishingCredentialsError",
        appName,
        resourceGroup,
        stringifyError(error),
        "https://learn.microsoft.com/en-us/rest/api/appservice/web-apps/list-publishing-credentials#code-try-0"
      ),
      displayMessage: getLocalizedString(
        "error.deploy.GetPublishingCredentialsError.Notification",
        appName,
        resourceGroup
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class GetZipDeployEndpointError extends UserError {
  constructor(appName: string, resourceGroup: string, error: Error, helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.GetZipDeployEndpointError",
        appName,
        resourceGroup,
        stringifyError(error)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.GetZipDeployEndpointError.Notification",
        appName,
        resourceGroup
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class DeployZipPackageError extends UserError {
  constructor(endpoint: string, error: Error, helpLink?: string) {
    super({
      name: `DeployZipPackageError${error.name ?? ""}`,
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.DeployZipPackageError",
        endpoint,
        stringifyError(error),
        "https://learn.microsoft.com/azure/app-service/deploy-zip?tabs=cli"
      ),
      displayMessage: getLocalizedString(
        "error.deploy.DeployZipPackageError.Notification",
        endpoint
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class CheckDeploymentStatusError extends UserError {
  constructor(location: string, error: Error, helpLink?: string, displayMessage?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.CheckDeploymentStatusError",
        location,
        stringifyError(error)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.CheckDeploymentStatusError",
        location,
        displayMessage || error.message
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class AzureStorageClearBlobsError extends UserError {
  constructor(storageName: string, errorResponse: BlobDeleteResponse, helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.AzureStorageClearBlobsError",
        storageName,
        JSON.stringify(errorResponse, Object.getOwnPropertyNames(errorResponse), 4)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.AzureStorageClearBlobsError.Notification",
        storageName
      ),
      error: errorResponse as any,
      helpLink: helpLink,
      categories: [ErrorCategory.External],
    });
  }
}

export class AzureStorageUploadFilesError extends UserError {
  constructor(
    storageName: string,
    localFolder: string,
    errorResponse: BlobUploadCommonResponse,
    helpLink?: string
  ) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.AzureStorageUploadFilesError",
        localFolder,
        storageName,
        JSON.stringify(errorResponse, Object.getOwnPropertyNames(errorResponse), 4)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.AzureStorageUploadFilesError.Notification",
        localFolder,
        storageName
      ),
      helpLink: helpLink,
      error: errorResponse as any,
      categories: [ErrorCategory.External],
    });
  }
}

export class AzureStorageGetContainerError extends UserError {
  constructor(storageName: string, containerName: string, error: any, helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.AzureStorageGetContainerError",
        containerName,
        storageName,
        stringifyError(error)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.AzureStorageGetContainerError.Notification",
        containerName,
        storageName,
        error.message || ""
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class AzureStorageGetContainerPropertiesError extends UserError {
  constructor(storageName: string, containerName: string, error: any, helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.AzureStorageGetContainerPropertiesError",
        containerName,
        storageName,
        stringifyError(error)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.AzureStorageGetContainerPropertiesError.Notification",
        containerName,
        storageName,
        error.message || ""
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

export class AzureStorageSetContainerPropertiesError extends UserError {
  constructor(storageName: string, containerName: string, error: any, helpLink?: string) {
    super({
      source: "azureDeploy",
      message: getDefaultString(
        "error.deploy.AzureStorageSetContainerPropertiesError",
        containerName,
        storageName,
        stringifyError(error)
      ),
      displayMessage: getLocalizedString(
        "error.deploy.AzureStorageSetContainerPropertiesError.Notification",
        containerName,
        storageName,
        error.message || ""
      ),
      helpLink: helpLink,
      error: error,
      categories: [ErrorCategory.External],
    });
  }
}

function stringifyError(error: Error): string {
  try {
    return JSON.stringify(error, Object.getOwnPropertyNames(error));
  } catch (e) {
    return `${error.toString()} stack: ${error.stack ?? ""}`;
  }
}
