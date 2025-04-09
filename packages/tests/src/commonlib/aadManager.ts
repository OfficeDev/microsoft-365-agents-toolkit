// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import axios, { AxiosInstance } from "axios";

import { M365TokenProvider } from "@microsoft/teamsfx-api";

import MockM365TokenProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365LoginUserPassword";
import { GraphScopes } from "@microsoft/teamsfx-core";

interface IAadAppInfo {
  displayName: string;
  id: string;
  appId: string;
  createdDateTime: string;
  deletedDateTime: string;
}

function delay(ms: number) {
  // tslint:disable-next-line no-string-based-set-timeout
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class AadManager {
  private static instance: AadManager;

  private axios: AxiosInstance;

  private constructor(access?: string) {
    this.axios = axios.create({
      baseURL: "https://graph.microsoft.com/v1.0/",
      headers: {
        authorization: `Bearer ${access}`,
        ConsistencyLevel: "eventual",
        "content-type": "application/json",
      },
    });
  }

  public static async init(
    provider: M365TokenProvider = MockM365TokenProvider
  ): Promise<AadManager> {
    if (!AadManager.instance) {
      const res = await provider.getAccessToken({
        scopes: GraphScopes,
      });
      if (res.isErr()) {
        throw res.error;
      }
      this.instance = new AadManager(res.value);
    }
    return this.instance;
  }

  public async searchAliveAadApps(
    prefix: string,
    offsetHour = 0
  ): Promise<IAadAppInfo[]> {
    const result = await this.axios!.get(
      `applications?$filter=startswith(displayName, '${prefix}')&$count=true&$top=100&$orderby=displayName`
    );
    const apps = result.data.value as IAadAppInfo[];
    return Promise.resolve(
      apps.filter(
        (app) =>
          Date.now() - new Date(app.createdDateTime).getTime() >
          offsetHour * 3600 * 1000
      )
    );
  }

  public async searchDeletedAadApps(
    prefix: string,
    offsetHour = 0
  ): Promise<IAadAppInfo[]> {
    const result = await this.axios!.get(
      `directory/deleteditems/microsoft.graph.application?$filter=startswith(displayName, '${prefix}')&$count=true&$top=100&$orderby=displayName`
    );
    const apps = result.data.value as IAadAppInfo[];
    return Promise.resolve(
      apps.filter(
        (app) =>
          Date.now() - new Date(app.createdDateTime).getTime() >
          offsetHour * 3600 * 1000
      )
    );
  }

  public async searchAadApps(
    contain: string,
    offsetHour = 0
  ): Promise<IAadAppInfo[]> {
    return new Promise<IAadAppInfo[]>(async (resolve) => {
      const [aliveAadApps, deletedAadApps] = await Promise.all([
        this.searchAliveAadApps(contain, offsetHour),
        this.searchDeletedAadApps(contain, offsetHour),
      ]);
      return resolve(aliveAadApps.concat(deletedAadApps));
    });
  }

  public async searchAadAppsByClientId(
    clientId: string
  ): Promise<IAadAppInfo[]> {
    return new Promise<IAadAppInfo[]>(async (resolve) => {
      const result = await this.axios!.get(
        `applications?$filter=appId eq '${clientId}'`
      );
      const apps = result.data.value as IAadAppInfo[];
      return resolve(apps);
    });
  }

  public async deleteAadAppById(id: string, retryTimes = 5) {
    return new Promise<boolean>(async (resolve) => {
      try {
        await this.axios!.delete(`applications/${id}`);
      } finally {
        for (let i = 0; i < retryTimes; ++i) {
          try {
            await this.axios!.delete(`directory/deletedItems/${id}`);
            return resolve(true);
          } catch {
            await delay(2000);
            if (i < retryTimes - 1) {
              console.warn(
                `[Retry] failed to delete the Aad app with id: ${id}`
              );
            }
          }
        }
        return resolve(false);
      }
    });
  }

  public async deleteAadAppsByClientId(clientId: string, retryTimes = 5) {
    if (!clientId) {
      return [true];
    }
    const aadApps = await this.searchAadAppsByClientId(clientId);
    const promises = aadApps.map((app) =>
      this.deleteAadAppById(app.id, retryTimes)
    );
    return Promise.all(promises).then((results) =>
      results.map((result, index) => {
        if (result) {
          console.log(
            `[Success] delete the Aad app with id: ${aadApps[index].id}, appId: ${aadApps[index].appId}`
          );
          return true;
        } else {
          console.log(
            `[Failed] delete the Aad app with id: ${aadApps[index].id}, appId: ${aadApps[index].appId}`
          );
          return false;
        }
      })
    );
  }

  public async deleteAadApps(contain: string, offsetHour = 0, retryTimes = 5) {
    const aadApps = await this.searchAadApps(contain, offsetHour);
    console.log(
      `There are ${aadApps.length} Aad apps created ${offsetHour.toFixed(
        2
      )} hours ago. Deleting...`
    );

    const promises = aadApps.map((app) =>
      this.deleteAadAppById(app.id, retryTimes)
    );
    const results = await Promise.all(promises);
    results.forEach((result, index) => {
      if (result) {
        console.log(
          `[Success] delete the Aad app with id: ${aadApps[index].id}, appId: ${aadApps[index].appId}`
        );
      } else {
        console.log(
          `[Failed] delete the Aad app with id: ${aadApps[index].id}, appId: ${aadApps[index].appId}`
        );
      }
    });
  }
}
