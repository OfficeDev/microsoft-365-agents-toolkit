// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { UsernamePasswordCredential } from "@azure/identity";
import { ResourceManagementClient } from "@azure/arm-resources";

import * as azureConfig from "@microsoft/m365agentstoolkit-cli/src/commonlib/common/userPasswordConfig";

const tenantId = azureConfig.AZURE_TENANT_ID || "";
const clientId = azureConfig.client_id;
const username = azureConfig.AZURE_ACCOUNT_NAME || "";
const password = azureConfig.AZURE_ACCOUNT_PASSWORD || "";
const subscriptionId = azureConfig.AZURE_SUBSCRIPTION_ID || "";

function delay(ms: number) {
  // tslint:disable-next-line no-string-based-set-timeout
  return new Promise((resolve) => setTimeout(resolve, ms));
}
export class ResourceGroupManager {
  private static client?: ResourceManagementClient;

  private constructor() {
    ResourceGroupManager.client = undefined;
  }

  private static async init() {
    if (!ResourceGroupManager.client) {
      const credential = new UsernamePasswordCredential(
        tenantId,
        clientId,
        username,
        password
      );
      ResourceGroupManager.client = new ResourceManagementClient(
        credential,
        subscriptionId
      );
    }
  }

  public static async getResourceGroup(name: string) {
    await ResourceGroupManager.init();
    return ResourceGroupManager.client!.resourceGroups.get(name);
  }

  public static async hasResourceGroup(name: string): Promise<boolean> {
    await ResourceGroupManager.init();
    try {
      await this.getResourceGroup(name);
      return Promise.resolve(true);
    } catch {
      return Promise.resolve(false);
    }
  }

  public static async searchResourceGroups(contain: string) {
    await ResourceGroupManager.init();

    const groups: any[] = [];
    for await (const page of ResourceGroupManager.client!.resourceGroups.list().byPage(
      {
        maxPageSize: 100,
      }
    )) {
      for (const group of page) {
        if (group.name?.includes(contain)) {
          groups.push(group);
        }
      }
    }
    return groups;
  }

  public static async deleteResourceGroup(
    name: string,
    retryTimes = 5
  ): Promise<boolean> {
    await ResourceGroupManager.init();
    return new Promise<boolean>(async (resolve) => {
      for (let i = 0; i < retryTimes; ++i) {
        try {
          await ResourceGroupManager.client!.resourceGroups.beginDeleteAndWait(
            name
          );
          return resolve(true);
        } catch (e) {
          await delay(2000);
          if (i < retryTimes - 1) {
            console.warn(
              `[Retry] clean up the Azure resoure group failed with name: ${name}`
            );
          }
        }
      }
      return resolve(false);
    });
  }

  public static async createOrUpdateResourceGroup(
    name: string,
    location: string
  ): Promise<boolean> {
    await ResourceGroupManager.init();
    return new Promise<boolean>(async (resolve) => {
      try {
        const resourceGroup = {
          location: location,
          name: name,
        };
        await ResourceGroupManager.client!.resourceGroups.createOrUpdate(
          name,
          resourceGroup
        );
        return resolve(true);
      } catch (e) {
        console.error(
          `Failed to create or update resource group ${name}. Error message: ${e.message}`
        );
        return resolve(false);
      }
    });
  }
}
