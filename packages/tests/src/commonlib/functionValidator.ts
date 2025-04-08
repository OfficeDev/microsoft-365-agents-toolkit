// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes } from "@microsoft/teamsfx-core";
import axios from "axios";
import * as chai from "chai";
import { EnvConstants, PluginId, StateConfigKey } from "./constants";
import {
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getSubscriptionIdFromResourceId,
  getWebappSettings,
  runWithRetry,
} from "./utilities";

const baseUrlListDeployments = (
  subscriptionId: string,
  rg: string,
  name: string
) =>
  `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.Web/sites/${name}/deployments?api-version=2019-08-01`;
const baseUrlListDeploymentLogs = (
  subscriptionId: string,
  rg: string,
  name: string,
  id: string
) =>
  `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.Web/sites/${name}/deployments/${id}/log?api-version=2019-08-01`;

enum BaseConfig {
  M365_CLIENT_ID = "M365_CLIENT_ID",
  M365_CLIENT_SECRET = "M365_CLIENT_SECRET",
  M365_AUTHORITY_HOST = "M365_AUTHORITY_HOST",
  M365_TENANT_ID = "M365_TENANT_ID",
  ALLOWED_APP_IDS = "ALLOWED_APP_IDS",
  API_ENDPOINT = "API_ENDPOINT",
  M365_APPLICATION_ID_URI = "M365_APPLICATION_ID_URI",
  IDENTITY_ID = "IDENTITY_ID",
  WEBSITE_CONTENTSHARE = "WEBSITE_CONTENTSHARE",
}

export enum ValidatorType {
  API_ENDPOINT = "API_ENDPOINT",
  FUNCTION_NAME = "FUNCTION_NAME",
}

type ConfigValidator = (webappSettings: any) => Promise<void>;

export class FunctionValidator {
  private ctx: any;
  private projectPath: string;
  private env: string;
  private customConfigsToValidate: ValidatorType[];

  private subscriptionId: string;
  private rg: string;
  private functionAppName: string;

  private readonly configValidatorMap: Record<ValidatorType, ConfigValidator> =
    {
      [ValidatorType.API_ENDPOINT]: async (
        webappSettings: any
      ): Promise<void> => {
        const endpoint =
          this.ctx?.[EnvConstants.FUNCTION_ENDPOINT] ??
          this.ctx?.[EnvConstants.FUNCTION_ENDPOINT_2];
        chai.assert.exists(endpoint);
        chai.assert.equal(webappSettings[BaseConfig.API_ENDPOINT], endpoint);
      },

      [ValidatorType.FUNCTION_NAME]: async (
        webappSettings: any
      ): Promise<void> => {
        const contentShare = webappSettings[BaseConfig.WEBSITE_CONTENTSHARE];
        if (contentShare) {
          const functionNameInAzure = contentShare.split("-")[0];
          chai.assert.exists(functionNameInAzure);
          chai.assert.equal(functionNameInAzure, this.functionAppName);
        }
      },
    };

  constructor(
    ctx: any,
    projectPath: string,
    env: string,
    customConfigsToValidate?: ValidatorType[]
  ) {
    console.log("Start to init validator for function.");

    this.ctx = ctx;
    this.projectPath = projectPath;
    this.env = env;

    const resourceId =
      ctx[EnvConstants.FUNCTION_ID] ??
      ctx[EnvConstants.FUNCTION_ID_2] ??
      ctx[PluginId.Function][StateConfigKey.functionAppResourceId];
    chai.assert.exists(resourceId);
    this.subscriptionId = getSubscriptionIdFromResourceId(resourceId);
    chai.assert.exists(this.subscriptionId);
    this.rg = getResourceGroupNameFromResourceId(resourceId);
    chai.assert.exists(this.rg);
    this.functionAppName = getSiteNameFromResourceId(resourceId);
    chai.assert.exists(this.functionAppName);

    this.customConfigsToValidate = customConfigsToValidate ?? [
      ValidatorType.API_ENDPOINT,
    ];

    console.log("Successfully init validator for function.");
  }

  public async validateProvision(): Promise<void> {
    console.log("Start to validate Function Provision.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    // Validating app settings
    console.log("validating app settings.");
    const webappSettingsResponse = await getWebappSettings(
      this.subscriptionId,
      this.rg,
      this.functionAppName,
      token as string
    );
    chai.assert.exists(webappSettingsResponse);

    for (const validatorType of this.customConfigsToValidate) {
      const validator = this.configValidatorMap[validatorType];
      await validator(webappSettingsResponse);
    }

    console.log("Successfully validate Function Provision.");
  }

  public async validateDeploy(): Promise<void> {
    console.log("Start to validate Function Deployment.");

    // Disable validate deployment since we have too many requests and the test is not stable.
    const tokenCredential =
      await MockAzureAccountProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    const deployments = await this.getDeployments(
      this.subscriptionId,
      this.rg,
      this.functionAppName,
      token as string
    );
    const deploymentId = deployments?.[0]?.properties?.id;
    const deploymentLog = await this.getDeploymentLog(
      this.subscriptionId,
      this.rg,
      this.functionAppName,
      token as string,
      deploymentId!
    );

    chai.assert.exists(
      deploymentLog?.find(
        (item: any) => item.properties.message === "Deployment successful."
      )
    );

    console.log("Successfully validate Function Deployment.");
  }

  private async getDeployments(
    subscriptionId: string,
    rg: string,
    name: string,
    token: string
  ) {
    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const functionGetResponse = await runWithRetry(() =>
        axios.get(baseUrlListDeployments(subscriptionId, rg, name))
      );

      return functionGetResponse?.data?.value;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }

  private async getDeploymentLog(
    subscriptionId: string,
    rg: string,
    name: string,
    token: string,
    id: string
  ) {
    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const functionGetResponse = await runWithRetry(() =>
        axios.get(baseUrlListDeploymentLogs(subscriptionId, rg, name, id))
      );

      return functionGetResponse?.data?.value;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }
}
