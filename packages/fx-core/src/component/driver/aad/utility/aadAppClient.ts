// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { LogProvider, M365TokenProvider } from "@microsoft/teamsfx-api";
import axios, { AxiosError, AxiosInstance, AxiosRequestHeaders } from "axios";
import axiosRetry, { IAxiosRetryConfig } from "axios-retry";
import { GraphScopes } from "../../../../common/constants";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { AadOwner } from "../../../../common/permissionInterface";
import { ErrorContextMW } from "../../../../common/globalVars";
import {
  DeleteOrUpdatePermissionFailedError,
  HostNameNotOnVerifiedDomainError,
} from "../error/aadManifestError";
import { ClientSecretNotAllowedError } from "../error/clientSecretNotAllowedError";
import { CredentialInvalidLifetimeError } from "../error/credentialInvalidLifetimeError";
import { AADApplication } from "../interface/AADApplication";
import { AADManifest } from "../interface/AADManifest";
import { IAADDefinition } from "../interface/IAADDefinition";
import { SignInAudience } from "../interface/signInAudience";
import { AadManifestHelper } from "./aadManifestHelper";
import { aadErrorCode } from "./constants";
// Another implementation of src\component\resource\aadApp\graph.ts to reduce call stacks
// It's our internal utility so make sure pass valid parameters to it instead of relying on it to handle parameter errors

// Missing this part will cause build failure when adding 'axios-retry' in AxiosRequestConfig
declare module "axios" {
  export interface AxiosRequestConfig {
    "axios-retry"?: IAxiosRetryConfig;
  }
}

export class AadAppClient {
  private readonly retryNumber: number = 5;
  private readonly tokenProvider: M365TokenProvider;
  private readonly logProvider: LogProvider | undefined;
  private readonly axios: AxiosInstance;
  private readonly baseUrl: string = "https://graph.microsoft.com/v1.0";

  constructor(m365TokenProvider: M365TokenProvider, logProvider?: LogProvider) {
    this.tokenProvider = m365TokenProvider;
    this.logProvider = logProvider;
    // Create axios instance which sets authorization header automatically before each MS Graph request
    this.axios = axios.create({
      baseURL: this.baseUrl,
    });
    this.axios.interceptors.request.use(async (config) => {
      this.logProvider?.debug(
        getLocalizedString("core.common.SendingApiRequest", config.url, JSON.stringify(config.data))
      );

      const tokenResponse = await this.tokenProvider.getAccessToken({ scopes: GraphScopes });
      if (tokenResponse.isErr()) {
        throw tokenResponse.error;
      }
      const token = tokenResponse.value;

      if (!config.headers) {
        config.headers = {} as AxiosRequestHeaders;
      }
      config.headers["Authorization"] = `Bearer ${token}`;

      return config;
    });
    this.axios.interceptors.response.use((response) => {
      this.logProvider?.debug(
        getLocalizedString("core.common.ReceiveApiResponse", JSON.stringify(response.data))
      );
      return response;
    });
    // Add retry logic. Retry post request may result in creating additional resources but should be fine in Microsoft Entra driver.
    axiosRetry(this.axios, {
      retries: this.retryNumber,
      retryDelay: axiosRetry.exponentialDelay, // exponetial delay time: Math.pow(2, retryNumber) * 100
      retryCondition: (error) =>
        axiosRetry.isNetworkError(error) || axiosRetry.isRetryableError(error), // retry when there's network error or 5xx error
    });
  }
  @hooks([ErrorContextMW({ source: "Graph", component: "AadAppClient" })])
  public async createAadApp(
    displayName: string,
    signInAudience: SignInAudience = SignInAudience.AzureADMyOrg,
    serviceManagementReference?: string
  ): Promise<AADApplication> {
    const requestBody: IAADDefinition = {
      displayName: displayName,
      signInAudience: signInAudience,
      serviceManagementReference: serviceManagementReference,
    }; // Create a Microsoft Entra app and optionally set service tree id

    const response = await this.axios.post("applications", requestBody);

    return <AADApplication>response.data;
  }

  @hooks([ErrorContextMW({ source: "Graph", component: "AadAppClient" })])
  public async deleteAadApp(id: string): Promise<void> {
    await this.axios.delete(`applications/${id}`);
  }

  @hooks([ErrorContextMW({ source: "Graph", component: "AadAppClient" })])
  public async generateClientSecret(
    objectId: string,
    clientSecretExpireDays = 180, // Recommended lifetime from Azure Portal
    clientSecretDescription = "default"
  ): Promise<string> {
    const startDate = new Date();
    const endDate = new Date(startDate.getTime());
    endDate.setDate(endDate.getDate() + clientSecretExpireDays);
    const requestBody = {
      passwordCredential: {
        displayName: clientSecretDescription,
        endDateTime: endDate.toISOString(),
        startDateTime: startDate.toISOString(),
      },
    };

    try {
      const response = await this.axios.post(`applications/${objectId}/addPassword`, requestBody, {
        "axios-retry": {
          retries: this.retryNumber,
          retryDelay: axiosRetry.exponentialDelay,
          retryCondition: (error) =>
            axiosRetry.isNetworkError(error) ||
            axiosRetry.isRetryableError(error) ||
            this.is404Error(error), // also retry 404 error since Microsoft Entra need sometime to sync created Microsoft Entra app data
        },
      });

      return response.data.secretText;
    } catch (err) {
      if (axios.isAxiosError(err) && err.response) {
        if (
          err.response.data?.error?.code === aadErrorCode.credentialInvalidLifetimeAsPerAppPolicy
        ) {
          throw new CredentialInvalidLifetimeError(AadAppClient.name);
        }
        if (
          err.response.data?.error?.code === aadErrorCode.credentialTypeNotAllowedAsPerAppPolicy
        ) {
          throw new ClientSecretNotAllowedError(AadAppClient.name);
        }
      }
      throw err;
    }
  }

  @hooks([ErrorContextMW({ source: "Graph", component: "AadAppClient" })])
  public async updateAadApp(manifest: AADManifest | AADApplication): Promise<void> {
    const objectId = manifest.id!; // You need to ensure the object id exists in manifest
    const requestBody = AadManifestHelper.manifestToApplication(manifest);
    try {
      await this.axios.patch(`applications/${objectId}`, requestBody, {
        "axios-retry": {
          retries: this.retryNumber,
          retryDelay: axiosRetry.exponentialDelay,
          retryCondition: (error) =>
            axiosRetry.isNetworkError(error) ||
            axiosRetry.isRetryableError(error) ||
            this.is404Error(error) || // also retry 404 error since Microsoft Entra need sometime to sync created Microsoft Entra app data
            this.is400Error(error), // sometimes Microsoft Entra will complain OAuth permission not found if we pre-authorize a newly created permission
        },
      });
    } catch (err) {
      if (axios.isAxiosError(err) && err.response && err.response.status === 400) {
        if (err.response.data.error?.code === aadErrorCode.permissionErrorCode) {
          throw new DeleteOrUpdatePermissionFailedError(AadAppClient.name);
        }
        if (err.response.data.error?.code === aadErrorCode.hostNameNotOnVerifiedDomain) {
          throw new HostNameNotOnVerifiedDomainError(
            AadAppClient.name,
            err.response.data.error.message
          );
        }
      }
      throw err;
    }
  }
  @hooks([ErrorContextMW({ source: "Graph", component: "AadAppClient" })])
  public async getOwners(objectId: string): Promise<AadOwner[] | undefined> {
    const response = await this.axios.get(`applications/${objectId}/owners`, {
      "axios-retry": {
        retries: this.retryNumber,
        retryDelay: axiosRetry.exponentialDelay,
        retryCondition: (error) =>
          axiosRetry.isNetworkError(error) ||
          axiosRetry.isRetryableError(error) ||
          this.is404Error(error), // also retry 404 error since Microsoft Entra need sometime to sync created Microsoft Entra app data
      },
    });

    const aadOwners: AadOwner[] = [];
    for (const aadOwner of response.data.value) {
      aadOwners.push({
        userObjectId: aadOwner.id,
        resourceId: objectId,
        displayName: aadOwner.displayName,
        // For guest account, aadOwner.userPrincipalName will contains "EXT", thus use mail instead.
        userPrincipalName: aadOwner.mail ?? aadOwner.userPrincipalName,
      });
    }

    return aadOwners;
  }
  @hooks([ErrorContextMW({ source: "Graph", component: "AadAppClient" })])
  public async addOwner(objectId: string, userObjectId: string): Promise<void> {
    const requestBody = {
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      "@odata.id": `${this.axios.defaults.baseURL}/directoryObjects/${userObjectId}`,
    };

    await this.axios.post(`applications/${objectId}/owners/$ref`, requestBody, {
      "axios-retry": {
        retries: this.retryNumber,
        retryDelay: axiosRetry.exponentialDelay,
        retryCondition: (error) =>
          axiosRetry.isNetworkError(error) ||
          axiosRetry.isRetryableError(error) ||
          this.is404Error(error), // also retry 404 error since Microsoft Entra need sometime to sync created Microsoft Entra app data
      },
    });
  }

  // only use it to retry 404 errors for create client secret / update Microsoft Entra app requests right after Microsoft Entra app creation
  private is404Error(error: AxiosError<any>): boolean {
    return error.code !== "ECONNABORTED" && (!error.response || error.response.status === 404);
  }

  private is400Error(error: AxiosError<any>): boolean {
    return error.code !== "ECONNABORTED" && (!error.response || error.response.status === 400);
  }
}
