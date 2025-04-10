// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes } from "@microsoft/teamsfx-core";
import axios from "axios";
import * as chai from "chai";
import * as fs from "fs";
import * as path from "path";
import { EnvConstants } from "./constants";
import {
  getResourceGroupNameFromResourceId,
  getSubscriptionIdFromResourceId,
  parseFromResourceId,
} from "./utilities";

const baseUrlContainer = (
  subscriptionId: string,
  resourceGroupName: string,
  storageName: string,
  containerName: string
) =>
  `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Storage/storageAccounts/${storageName}/blobServices/default/containers/${containerName}?api-version=2021-01-01`;

const baseUrlBlob = (
  storageName: string,
  containerName: string,
  sasToken: string
) =>
  `https://${storageName}.blob.core.windows.net/${containerName}?restype=container&comp=list&${sasToken}`;

const baseUrlSasToken = (
  subscriptionId: string,
  resourceGroupName: string,
  storageName: string
) =>
  `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Storage/storageAccounts/${storageName}/ListAccountSas?api-version=2021-01-01`;

class DependentPluginInfo {
  public static readonly functionPluginName = "fx-resource-function";
  public static readonly apiEndpoint = "functionEndpoint";

  public static readonly solutionPluginName = "solution";
  public static readonly resourceGroupName: string = "resourceGroupName";
  public static readonly subscriptionId: string = "subscriptionId";
  public static readonly resourceNameSuffix: string = "resourceNameSuffix";
  public static readonly location: string = "location";
  public static readonly programmingLanguage: string = "programmingLanguage";

  public static readonly aadPluginName: string =
    "fx-resource-aad-app-for-teams";
  public static readonly aadClientId: string = "clientId";
  public static readonly aadClientSecret: string = "clientSecret";
  public static readonly oauthHost: string = "oauthHost";
  public static readonly tenantId: string = "tenantId";
  public static readonly applicationIdUris: string = "applicationIdUris";

  public static readonly frontendPluginName: string =
    "fx-resource-frontend-hosting";
  public static readonly frontendEndpoint: string = "endpoint";
  public static readonly frontendDomain: string = "domain";
}

interface IFrontendObject {
  storageName: string;
  containerName: string;
}

export class FrontendValidator {
  private static subscriptionId: string;
  private static resourceGroupName: string;

  private static storageResourceIdKeyName = "storageResourceId";

  public static init(ctx: any): IFrontendObject {
    console.log("Start to init validator for Frontend.");

    const resourceId = this.getResourceIdFromCtx(ctx);
    chai.assert.exists(resourceId);

    this.subscriptionId = getSubscriptionIdFromResourceId(resourceId);
    this.resourceGroupName = getResourceGroupNameFromResourceId(resourceId);

    const frontendObject = {
      storageName: this.getStorageAccountName(resourceId),
      containerName: "$web",
    };

    console.log("Successfully init validator for Frontend.");
    return frontendObject;
  }

  public static async validateScaffold(
    projectPath: string,
    programmingLanguage: string
  ): Promise<void> {
    const indexFile: { [key: string]: string } = {
      typescript: "index.tsx",
      javascript: "index.jsx",
    };
    const indexPath = path.resolve(
      projectPath,
      "tabs",
      "src",
      indexFile[programmingLanguage]
    );

    fs.access(indexPath, fs.constants.F_OK, (err) => {
      // err is null means file exists
      chai.assert.isNull(err);
    });
  }

  public static async validateScaffoldV3(
    projectPath: string,
    programmingLanguage: string
  ): Promise<void> {
    const indexFile: { [key: string]: string } = {
      typescript: "index.tsx",
      javascript: "index.jsx",
    };
    const indexPath = path.resolve(
      projectPath,
      "src",
      indexFile[programmingLanguage]
    );
    fs.access(indexPath, fs.constants.F_OK, (err) => {
      chai.assert.isNull(err);
    });
  }

  public static async validateProvision(
    frontendObject: IFrontendObject
  ): Promise<void> {
    console.log("Start to validate Frontend Provision.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;
    chai.assert.exists(token);

    console.log("Validating Storage Container.");
    const response = await this.getContainer(
      this.subscriptionId,
      this.resourceGroupName,
      frontendObject,
      token as string
    );
    chai.assert.exists(response);

    console.log("Successfully validate Frontend Provision.");
  }

  public static async validateDeploy(
    frontendObject: IFrontendObject
  ): Promise<void> {
    console.log("Start to validate Frontend Deploy.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;
    chai.assert.exists(token);

    const sasToken = await this.getSasToken(
      this.subscriptionId,
      this.resourceGroupName,
      frontendObject.storageName,
      token as string
    );
    chai.assert.exists(sasToken);

    console.log("Validating Storage blobs.");
    const response = await this.getBlobs(
      frontendObject.storageName,
      frontendObject.containerName,
      sasToken as string
    );
    chai.assert.exists(response);

    console.log("Successfully validate Frontend Deploy.");
  }

  private static getStorageAccountName(storageResourceId: string): string {
    const result = parseFromResourceId(
      /providers\/Microsoft.Storage\/storageAccounts\/([^\/]*)/i,
      storageResourceId
    );
    if (!result) {
      throw new Error(
        `Cannot parse storage accounts name from resource id ${storageResourceId}`
      );
    }
    return result;
  }

  private static async getContainer(
    subscriptionId: string,
    resourceGroupName: string,
    frontendObject: IFrontendObject,
    token: string
  ) {
    try {
      const frontendContainerResponse = await axios.get(
        baseUrlContainer(
          subscriptionId,
          resourceGroupName,
          frontendObject.storageName,
          frontendObject.containerName
        ),
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );

      return frontendContainerResponse?.data?.name;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }

  private static async getBlobs(
    storageName: string,
    containerName: string,
    sasToken: string
  ) {
    try {
      const frontendBlobResponse = await axios.get(
        baseUrlBlob(storageName, containerName, sasToken),
        {
          transformRequest: (data, headers) => {
            delete headers?.["Authorization"];
          },
        }
      );
      return frontendBlobResponse?.data;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }

  private static getResourceIdFromCtx(ctx: any): string {
    return (
      ctx[EnvConstants.TAB_AZURE_STORAGE_RESOURCE_ID] ??
      ctx[EnvConstants.TAB_AZURE_APP_SERVICE_RESOURCE_ID]
    );
  }

  private static async getSasToken(
    subscriptionId: string,
    resourceGroupName: string,
    storageName: string,
    token: string
  ) {
    try {
      const expiredDate = new Date();
      expiredDate.setDate(new Date().getDate() + 3);

      const sasTokenResponse = await axios.post(
        baseUrlSasToken(subscriptionId, resourceGroupName, storageName),
        {
          signedExpiry: expiredDate.toISOString(),
          signedPermission: "rl",
          signedResourceTypes: "sco",
          signedServices: "bf",
        },
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );
      return sasTokenResponse?.data?.accountSasToken;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }
}

export class FrontendWebAppConfig {
  public static readonly clientId = "TeamsFx__Authentication__ClientId";
  public static readonly clientSecret = "TeamsFx__Authentication__ClientSecret";
  public static readonly authEndpoint =
    "TeamsFx__Authentication__InitiateLoginEndpoint";
  public static readonly authority = "TeamsFx__Authentication__OAuthAuthority";
}
