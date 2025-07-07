// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { LogProvider, M365TokenProvider } from "@microsoft/teamsfx-api";
import axios, { AxiosError, AxiosInstance, AxiosRequestHeaders } from "axios";
import axiosRetry, { IAxiosRetryConfig } from "axios-retry";
import { GraphScopes } from "../../../../common/constants";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { ErrorContextMW } from "../../../../common/globalVars";
import {
  ISharePointEmbeddedContainerType,
  SharePointEmbeddedContainerTypeSettings,
  SPContainerTypeBillingClassification,
} from "../interface/sharePointEmbeddedContainerType";

// Missing this part will cause build failure when adding 'axios-retry' in AxiosRequestConfig
declare module "axios" {
  export interface AxiosRequestConfig {
    "axios-retry"?: IAxiosRetryConfig;
  }
}

export class SPEContainerTypeAppClient {
  private readonly baseUrl: string = "https://graph.microsoft.com/beta/storage/fileStorage";
  private readonly tokenProvider: M365TokenProvider;
  private readonly logProvider: LogProvider | undefined;
  private readonly retryNumber: number = 5;
  private readonly axios: AxiosInstance;

  constructor(tokenProvider: M365TokenProvider, logProvider?: LogProvider) {
    this.tokenProvider = tokenProvider;
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
      // harcode token for now

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

  @hooks([ErrorContextMW({ source: "Graph", component: "SharePointEmbeddedCTAppClient" })])
  public async createSPEContainerType(
    owningApplicationId: string,
    billingClassification: string,
    name?: string,
    discoverable?: boolean
  ): Promise<ISharePointEmbeddedContainerType> {
    const speContainerTypeSettings: SharePointEmbeddedContainerTypeSettings = {
      isDiscoverabilityEnabled: discoverable,
    };

    const requestBody: ISharePointEmbeddedContainerType = {
      name: name,
      owningAppId: owningApplicationId,
      billingClassification: billingClassification as SPContainerTypeBillingClassification,
      settings: speContainerTypeSettings,
    };

    const response = await this.axios.post("containerTypes", requestBody);
    return <ISharePointEmbeddedContainerType>response.data;
  }
}
