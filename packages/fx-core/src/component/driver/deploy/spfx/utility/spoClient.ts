// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";

// eslint-disable-next-line @typescript-eslint/no-namespace
export namespace SPOClient {
  let baseUrl = "";

  export function setBaseUrl(tenant: string) {
    baseUrl = tenant;
  }

  /**
   * Creates a new axios instance to call SharePoint Service to prevent setting the accessToken on global instance.
   * @param {string}  spoToken
   * @returns {AxiosInstance}
   */
  function createRequesterWithToken(spoToken: string): AxiosInstance {
    const instance = axios.create({
      baseURL: baseUrl,
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${spoToken}`;
    instance.defaults.headers.common["Client-Source"] = "teamstoolkit";
    instance.interceptors.request.use(function (config) {
      config.params = { teamstoolkit: true, ...config.params };
      return config;
    });
    return instance;
  }

  /**
   *
   * @param spoToken
   * @returns tenant app catalog URL
   */
  export async function getAppCatalogSite(spoToken: string): Promise<string | undefined> {
    const requester = createRequesterWithToken(spoToken);
    const res = await requester.get("/_api/SP_TenantSettings_Current");
    if (res && res.data && res.data.CorporateCatalogUrl) {
      return res.data.CorporateCatalogUrl;
    } else {
      return undefined;
    }
  }

  /**
   *
   * @param spoToken
   * @param fileName - *.sppkg file name
   * @param file - binary of *.sppkg
   */
  export async function uploadAppPackage(
    spoToken: string,
    fileName: string,
    file: Buffer
  ): Promise<any> {
    const requester = createRequesterWithToken(spoToken);
    await requester.post(
      `/_api/web/tenantappcatalog/Add(overwrite=true, url='${fileName}')`,
      file,
      {
        maxBodyLength: Infinity,
      }
    );
  }

  /**
   *
   * @param spoToken
   * @param appId - appId in package-solution.json
   */
  export async function deployAppPackage(spoToken: string, appId: string): Promise<any> {
    const deploySetting = {
      skipFeatureDeployment: true,
    };
    const requester = createRequesterWithToken(spoToken);
    await requester.post(
      `/_api/web/tenantappcatalog/AvailableApps/GetById('${appId}')/Deploy`,
      deploySetting
    );
  }

  export async function createAppCatalog(spoToken: string): Promise<any> {
    const requester = createRequesterWithToken(spoToken);
    await requester.post(`/_api/web/EnsureTenantAppCatalog(callerId='teamsdev')`);
  }
}
