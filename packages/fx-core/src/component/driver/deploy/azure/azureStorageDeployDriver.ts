// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author FanH <Siglud@gmail.com>
 */
import { AzureDeployImpl } from "./impl/azureDeployImpl";
import { DeployStepArgs } from "../../interface/buildAndDeployArgs";
import {
  BlobDeleteResponse,
  BlobItem,
  BlobUploadCommonResponse,
  BlockBlobParallelUploadOptions,
  ContainerClient,
} from "@azure/storage-blob";
import { DeployConstant } from "../../../constant/deployConstant";
import { forEachFileAndDir } from "../../../utils/fileOperation";
import * as fs from "fs-extra";
import path from "path";
import * as mime from "mime";
import { FxError, LogProvider, Result } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { DriverContext, AzureResourceInfo } from "../../interface/commonArgs";
import { createBlobServiceClient } from "../../../utils/azureResourceOperation";
import { TokenCredential } from "@azure/identity";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../../constant/commonConstant";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { wrapAzureOperation } from "../../../utils/azureSdkErrorHandler";
import {
  AzureStorageClearBlobsError,
  AzureStorageGetContainerError,
  AzureStorageUploadFilesError,
} from "../../../../error";
import { ProgressMessages } from "../../../messages";
import { ErrorContextMW } from "../../../../common/globalVars";

const ACTION_NAME = "azureStorage/deploy";

@Service(ACTION_NAME)
export class AzureStorageDeployDriver implements StepDriver {
  readonly description: string = getLocalizedString(
    "driver.deploy.deployToAzureStorageDescription"
  );

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    const impl = new AzureStorageDeployDriverImpl(args, ctx);
    return impl.run();
  }
}

/**
 * deploy to Azure Storage
 */
export class AzureStorageDeployDriverImpl extends AzureDeployImpl {
  protected summaries: () => string[] = () => [
    getLocalizedString("driver.deploy.azureStorageDeployDetailSummary", this.distDirectory),
  ];
  protected summaryPrepare: () => string[] = () => [];

  pattern =
    /\/subscriptions\/([^\/]*)\/resourceGroups\/([^\/]*)\/providers\/Microsoft.Storage\/storageAccounts\/([^\/]*)/i;

  protected helpLink = "https://aka.ms/teamsfx-actions/azure-storage-deploy";

  @hooks([ErrorContextMW({ source: "Azure", component: "AzureStorageDeployDriverImpl" })])
  async azureDeploy(
    args: DeployStepArgs,
    azureResource: AzureResourceInfo,
    azureCredential: TokenCredential
  ): Promise<void> {
    this.context.logProvider.debug("Start deploying to Azure Storage Service");
    this.context.logProvider.debug("Get Azure Storage Service deploy credential");
    const containerClient = await AzureStorageDeployDriverImpl.createContainerClient(
      azureResource,
      azureCredential
    );
    // delete all existing blobs
    await this.deleteAllBlobs(containerClient, azureResource.instanceId, this.context.logProvider);
    this.context.logProvider.debug("Uploading files to Azure Storage Service");
    // upload all to storage
    const ig = await this.handleIgnore(args, this.context);
    const sourceFolder = this.distDirectory;
    const tasks: Promise<BlobUploadCommonResponse>[] = [];
    await forEachFileAndDir(
      sourceFolder,
      (filePath: string, stats: fs.Stats) => {
        const destFilePath: string = path.relative(sourceFolder, filePath);
        if (!destFilePath || stats.isDirectory()) {
          return;
        }
        const options: BlockBlobParallelUploadOptions = {
          blobHTTPHeaders: {
            blobContentType: mime.getType(filePath) || undefined,
          },
        };
        const client = containerClient.getBlockBlobClient(destFilePath);
        tasks.push(client.uploadFile(filePath, options));
      },
      (itemPath: string) => {
        return !ig.test(path.relative(sourceFolder, itemPath)).ignored;
      }
    );
    const responses = await Promise.all(tasks);
    const errorResponse = responses.find((res) => res.errorCode);
    if (errorResponse) {
      throw new AzureStorageUploadFilesError(
        azureResource.instanceId,
        sourceFolder,
        errorResponse,
        this.helpLink
      );
    }
    this.context.logProvider.debug("Upload files to Azure Storage Service successfully");
    return;
  }

  private static async createContainerClient(
    azureResource: AzureResourceInfo,
    azureCredential: TokenCredential
  ): Promise<ContainerClient> {
    const blobServiceClient = await createBlobServiceClient(azureResource, azureCredential);
    return await wrapAzureOperation(
      async () => {
        const container = blobServiceClient.getContainerClient(
          DeployConstant.AZURE_STORAGE_CONTAINER_NAME
        );
        if (!(await container.exists())) {
          await container.create();
        }
        return container;
      },
      (e) =>
        new AzureStorageGetContainerError(
          azureResource.instanceId,
          DeployConstant.AZURE_STORAGE_CONTAINER_NAME,
          e
        ),
      (e) =>
        new AzureStorageGetContainerError(
          azureResource.instanceId,
          DeployConstant.AZURE_STORAGE_CONTAINER_NAME,
          e
        )
    );
  }

  private async deleteAllBlobs(
    client: ContainerClient,
    storageName: string,
    logProvider: LogProvider
  ): Promise<void> {
    logProvider.debug(
      `Deleting all existing blobs in container '${DeployConstant.AZURE_STORAGE_CONTAINER_NAME}' for Azure Storage account '${storageName}'.`
    );

    const deleteJobs: Promise<BlobDeleteResponse>[] = [];
    for await (const blob of client.listBlobsFlat()) {
      if (AzureStorageDeployDriverImpl.isBlobFile(blob)) {
        deleteJobs.push(client.deleteBlob(blob.name));
      }
    }

    const responses = await Promise.all(deleteJobs);
    logProvider.verbose(`Delete all blobs responses: ${JSON.stringify(responses)}`);
    const errorResponse = responses.find((res) => res.errorCode);
    if (errorResponse) {
      throw new AzureStorageClearBlobsError(storageName, errorResponse);
    }
  }

  private static isBlobFile(blob: BlobItem): boolean {
    return (blob.properties.contentLength ?? -1) > 0;
  }

  updateProgressbar() {
    void this.progressBar?.next(
      ProgressMessages.deployToAzure(this.workingDirectory, "Azure Storage Service")
    );
  }
}
