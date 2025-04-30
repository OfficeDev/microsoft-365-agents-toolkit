// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ResourceGroup, ResourceManagementClient } from "@azure/arm-resources";
import { UsernamePasswordCredential } from "@azure/identity";
import * as azureConfig from "@microsoft/m365agentstoolkit-cli/src/commonlib/common/userPasswordConfig";
import { strings } from "../utils/constants";

const tenantId = azureConfig.AZURE_TENANT_ID || "";
const clientId = azureConfig.client_id;
const username = azureConfig.AZURE_ACCOUNT_NAME || "";
const password = azureConfig.AZURE_ACCOUNT_PASSWORD || "";
const subscriptionId = azureConfig.AZURE_SUBSCRIPTION_ID || "";

function delay(ms: number) {
  // tslint:disable-next-line no-string-based-set-timeout
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class AzureHelper {
  private static instance: AzureHelper;

  private client: ResourceManagementClient;

  private constructor() {
    const credential = new UsernamePasswordCredential(
      tenantId,
      clientId,
      username,
      password
    );
    this.client = new ResourceManagementClient(credential, subscriptionId);
  }

  public static init() {
    if (!this.instance) {
      this.instance = new AzureHelper();
    }
    return this.instance;
  }

  public async getResourceGroup(name: string) {
    return this.client.resourceGroups.get(name);
  }

  public async hasResourceGroup(name: string) {
    try {
      await this.getResourceGroup(name);
      return true;
    } catch {
      return false;
    }
  }

  public async searchResourceGroups(contain: string) {
    const groups: ResourceGroup[] = [];
    for await (const page of this.client.resourceGroups.list().byPage({
      maxPageSize: 100,
    })) {
      groups.concat(page.filter((group) => group.name?.includes(contain)));
    }
    return groups;
  }

  public async deleteResourceGroup(
    name: string,
    retryTimes = 5
  ): Promise<string> {
    if (!name || !(await this.hasResourceGroup(name))) {
      return strings.deleteResourceGroup.skipped.replace("%s", name);
    }
    return new Promise(async (resolve) => {
      for (let i = 0; i < retryTimes; ++i) {
        try {
          await this.client.resourceGroups.beginDeleteAndWait(name);
          return resolve(
            strings.deleteResourceGroup.success.replace("%s", name)
          );
        } catch (e) {
          await delay(2000);
        }
      }
      return resolve(strings.deleteResourceGroup.failed.replace("%s", name));
    });
  }
}
