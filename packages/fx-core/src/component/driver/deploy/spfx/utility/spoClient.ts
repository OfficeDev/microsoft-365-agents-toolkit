// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";

export interface PreparedSharePointAppState {
  id: string;
  version: string;
  deployed: boolean;
  currentVersionDeployed: boolean;
  skipFeatureDeployment: boolean;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

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

  export async function getPreparedAvailableApp(
    spoToken: string,
    appId: string
  ): Promise<PreparedSharePointAppState | undefined> {
    const requester = createRequesterWithToken(spoToken);
    try {
      const response = await requester.get(
        `/_api/web/tenantappcatalog/AvailableApps/GetById('${appId}')`
      );
      const data: unknown = response.data;
      if (
        !isRecord(data) ||
        typeof data["ID"] !== "string" ||
        typeof data["AppCatalogVersion"] !== "string" ||
        !/^\d+\.\d+\.\d+\.\d+$/.test(data["AppCatalogVersion"]) ||
        typeof data["Deployed"] !== "boolean" ||
        typeof data["CurrentVersionDeployed"] !== "boolean" ||
        typeof data["IsPackageDefaultSkipFeatureDeployment"] !== "boolean"
      ) {
        throw new Error("SharePoint returned incomplete app catalog metadata.");
      }
      return {
        id: data["ID"],
        version: data["AppCatalogVersion"],
        deployed: data["Deployed"],
        currentVersionDeployed: data["CurrentVersionDeployed"],
        skipFeatureDeployment: data["IsPackageDefaultSkipFeatureDeployment"],
      };
    } catch (error) {
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        return undefined;
      }
      throw error;
    }
  }

  export async function uploadPreparedAppPackage(
    spoToken: string,
    fileName: string,
    file: Buffer,
    overwrite: boolean
  ): Promise<void> {
    if (!/^[a-zA-Z0-9._-]+\.sppkg$/i.test(fileName)) {
      throw new Error("The prepared SharePoint package file name is invalid.");
    }
    const requester = createRequesterWithToken(spoToken);
    await requester.post(
      `/_api/web/tenantappcatalog/Add(overwrite=${overwrite}, url='${fileName}')`,
      file,
      {
        headers: { "Content-Type": "application/octet-stream" },
        maxBodyLength: file.length,
      }
    );
  }

  export async function readPreparedAppPackage(
    spoToken: string,
    fileName: string
  ): Promise<Buffer | undefined> {
    if (!/^[a-zA-Z0-9._-]+\.sppkg$/i.test(fileName)) {
      throw new Error("The prepared SharePoint package file name is invalid.");
    }
    const requester = createRequesterWithToken(spoToken);
    const listResponse = await requester.get(
      `/_api/web/lists/getbytitle('Apps for SharePoint')/items?$select=File/ServerRelativeUrl&$expand=File&$filter=FileLeafRef eq '${fileName}'&$top=2`
    );
    const listData: unknown = listResponse.data;
    if (!isRecord(listData) || !Array.isArray(listData["value"])) {
      throw new Error("SharePoint returned an invalid app catalog file listing.");
    }
    const files = listData["value"].filter(isRecord);
    if (files.length === 0) {
      return undefined;
    }
    if (files.length !== 1 || !isRecord(files[0]["File"])) {
      throw new Error("SharePoint returned an ambiguous app catalog package file.");
    }
    const serverRelativeUrl = files[0]["File"]["ServerRelativeUrl"];
    if (
      typeof serverRelativeUrl !== "string" ||
      !serverRelativeUrl.startsWith("/") ||
      serverRelativeUrl.includes("\\") ||
      /[\r\n]/.test(serverRelativeUrl)
    ) {
      throw new Error("SharePoint returned an invalid app catalog package URL.");
    }
    const escapedUrl = serverRelativeUrl.replace(/'/g, "''");
    const fileResponse = await requester.get(
      `/_api/web/GetFileByServerRelativeUrl('${escapedUrl}')/$value`,
      { responseType: "arraybuffer", maxContentLength: 100 * 1024 * 1024 }
    );
    return Buffer.from(fileResponse.data);
  }

  export async function deployPreparedAppPackage(
    spoToken: string,
    appId: string,
    skipFeatureDeployment: boolean
  ): Promise<void> {
    const requester = createRequesterWithToken(spoToken);
    await requester.post(`/_api/web/tenantappcatalog/AvailableApps/GetById('${appId}')/Deploy`, {
      skipFeatureDeployment,
    });
  }

  export async function createAppCatalog(spoToken: string): Promise<any> {
    const requester = createRequesterWithToken(spoToken);
    await requester.post(`/_api/web/EnsureTenantAppCatalog(callerId='teamsdev')`);
  }
}
