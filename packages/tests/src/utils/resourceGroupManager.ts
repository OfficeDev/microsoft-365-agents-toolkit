// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { ResourceManagementClient } from "@azure/arm-resources";
import {
  UsernamePasswordCredential,
  ClientSecretCredential,
} from "@azure/identity";
import { Env } from "./env";
function delay(ms: number) {
  // tslint:disable-next-line no-string-based-set-timeout
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class ResourceGroupManager {
  private static instance: ResourceGroupManager;

  private static client?: ResourceManagementClient;

  private constructor() {
    ResourceGroupManager.client = undefined;
  }

  public static async init(): Promise<ResourceGroupManager> {
    if (!ResourceGroupManager.instance) {
      ResourceGroupManager.instance = new ResourceGroupManager();
      let credential;
      if (Env.servicePrincipalId && Env.servicePrincipalSecret) {
        credential = new ClientSecretCredential(
          Env.azureTenantId,
          Env.servicePrincipalId,
          Env.servicePrincipalSecret
        );
      } else {
        credential = new UsernamePasswordCredential(
          Env.azureTenantId,
          "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
          Env.azureAccountName,
          Env.azureAccountPassword
        );
      }
      ResourceGroupManager.client = new ResourceManagementClient(
        credential,
        Env.azureSubscriptionId
      );
    }
    return Promise.resolve(ResourceGroupManager.instance);
  }

  public async getResourceGroup(name: string) {
    return ResourceGroupManager.client!.resourceGroups.get(name);
  }

  public async hasResourceGroup(name: string) {
    try {
      await this.getResourceGroup(name);
      return Promise.resolve(true);
    } catch {
      return Promise.resolve(false);
    }
  }

  public async searchResourceGroups(contain: string) {
    // const groups = await ResourceGroupManager.client!.resourceGroups.list();
    // return groups.filter((group) => group.name?.includes(contain));
    const results: string[] = [];
    const res = ResourceGroupManager.client!.resourceGroups.list();
    let result;
    do {
      result = await res.next();
      if (result.value?.name?.includes(contain))
        results.push(result.value.name);
    } while (!result.done);
    return results;
  }

  public async createResourceGroup(
    name: string,
    location: string,
    retryTimes = 5
  ): Promise<boolean> {
    return new Promise<boolean>(async (resolve) => {
      for (let i = 0; i < retryTimes; ++i) {
        try {
          await ResourceGroupManager.client!.resourceGroups.createOrUpdate(
            name,
            { location: location }
          );
          return resolve(true);
        } catch (e) {
          await delay(2000);
          if (i < retryTimes - 1) {
            console.warn(
              `[Retry] create the Azure resoure group failed with name: ${name}`
            );
          }
        }
      }
      return resolve(false);
    });
  }

  public async deleteResourceGroup(
    name: string,
    retryTimes = 5
  ): Promise<boolean> {
    return new Promise<boolean>(async (resolve) => {
      for (let i = 0; i < retryTimes; ++i) {
        try {
          await ResourceGroupManager.client!.resourceGroups.beginDelete(name);
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
}
