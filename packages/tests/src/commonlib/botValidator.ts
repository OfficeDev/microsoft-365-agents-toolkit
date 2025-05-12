// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AzureScopes } from "@microsoft/teamsfx-core";
import axios from "axios";
import * as chai from "chai";
import * as fs from "fs";
import * as path from "path";

import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { EnvConstants, PluginId, StateConfigKey } from "./constants";

import {
  getExpectedM365ApplicationIdUri,
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getSubscriptionIdFromResourceId,
  getWebappSettings,
  getActivePluginsFromProjectSetting,
} from "./utilities";
import {
  getExpectedM365ClientSecret,
  getExpectedBotClientSecret,
} from "./cliHelper";

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
  BOT_ID = "BOT_ID",
  BOT_PASSWORD = "BOT_PASSWORD",
  INITIATE_LOGIN_ENDPOINT = "INITIATE_LOGIN_ENDPOINT",
  M365_APPLICATION_ID_URI = "M365_APPLICATION_ID_URI",
  M365_AUTHORITY_HOST = "M365_AUTHORITY_HOST",
  M365_CLIENT_ID = "M365_CLIENT_ID",
  M365_CLIENT_SECRET = "M365_CLIENT_SECRET",
  IDENTITY_ID = "IDENTITY_ID",
  M365_TENANT_ID = "M365_TENANT_ID",
  clientId = "clientId",
  tenantId = "tenantId",
}

enum FunctionConfig {
  API_ENDPOINT = "API_ENDPOINT",
}

enum SQLConfig {
  SQL_DATABASE_NAME = "SQL_DATABASE_NAME",
  SQL_ENDPOINT = "SQL_ENDPOINT",
}

export class BotValidator {
  private ctx: any;
  private projectPath: string;
  private env: string;
  private subscriptionId: string;
  private rg: string;
  private botAppSiteName: string;

  constructor(ctx: any, projectPath: string, env: string) {
    console.log("Start to init validator for Bot.");

    this.ctx = ctx;
    this.projectPath = projectPath;
    this.env = env;

    const resourceId = this.getResourceIdV3(ctx);
    chai.assert.exists(resourceId);
    this.subscriptionId = getSubscriptionIdFromResourceId(resourceId);
    chai.assert.exists(this.subscriptionId);
    this.rg = getResourceGroupNameFromResourceId(resourceId);
    chai.assert.exists(this.rg);
    this.botAppSiteName = getSiteNameFromResourceId(resourceId);
    chai.assert.exists(this.botAppSiteName);

    console.log("Successfully init validator for Bot.");
  }

  private getResourceIdV3(ctx: any): string {
    const botWebAppResourceId =
      ctx[EnvConstants.BOT_AZURE_APP_SERVICE_RESOURCE_ID];
    const botFunctionAppResourceId =
      ctx[EnvConstants.BOT_AZURE_FUNCTION_RESOURCE_ID];
    const botResourceId = ctx[EnvConstants.BOT_ID];
    const resourceId =
      botWebAppResourceId || botFunctionAppResourceId || botResourceId;
    return resourceId;
  }

  public static async validateScaffold(
    projectPath: string,
    programmingLanguage: string,
    srcPath = ""
  ): Promise<void> {
    const indexFile: { [key: string]: string } = {
      typescript: "index.ts",
      javascript: "index.js",
    };
    const indexPath = path.resolve(
      projectPath,
      "bot",
      srcPath,
      indexFile[programmingLanguage]
    );

    fs.access(indexPath, fs.constants.F_OK, (err) => {
      // err is null means file exists
      chai.assert.isNull(err);
    });
  }

  public async validateProvisionV3(includeAAD = true): Promise<void> {
    console.log("Start to validate Bot Provision.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    console.log("Validating env variables");
    const response = await getWebappSettings(
      this.subscriptionId,
      this.rg,
      this.botAppSiteName,
      token as string
    );
    chai.assert.exists(response);
    if (response[BaseConfig.clientId]) {
      chai.assert.equal(
        response[BaseConfig.clientId] ||
          response["Connections__BotServiceConnection__Settings__ClientId"],
        this.ctx[EnvConstants.clientId] as string
      );
    } else {
      chai.assert.equal(
        response[BaseConfig.BOT_ID] ||
          response["Connections__BotServiceConnection__Settings__ClientId"],
        this.ctx[EnvConstants.BOT_ID] as string
      );
    }

    if (includeAAD) {
      // TODO
    }
    // if (activeResourcePlugins.includes(PluginId.Function)) {
    //   chai.assert.equal(
    //     response[FunctionConfig.API_ENDPOINT],
    //     this.ctx[PluginId.Function][StateConfigKey.functionEndpoint] as string
    //   );
    // }
    // if (activeResourcePlugins.includes(PluginId.AzureSQL)) {
    //   chai.assert.equal(
    //     response[SQLConfig.SQL_ENDPOINT],
    //     this.ctx[PluginId.AzureSQL][StateConfigKey.sqlEndpoint] as string
    //   );
    //   chai.assert.equal(
    //     response[SQLConfig.SQL_DATABASE_NAME],
    //     this.ctx[PluginId.AzureSQL][StateConfigKey.databaseName] as string
    //   );
    // }

    console.log("Successfully validate Bot Provision.");
  }

  public async validateProvision(includeAAD = true): Promise<void> {
    console.log("Start to validate Bot Provision.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    const activeResourcePlugins = await getActivePluginsFromProjectSetting(
      this.projectPath
    );
    chai.assert.isArray(activeResourcePlugins);

    console.log("Validating app settings.");
    const response = await getWebappSettings(
      this.subscriptionId,
      this.rg,
      this.botAppSiteName,
      token as string
    );
    chai.assert.exists(response);
    chai.assert.equal(
      response[BaseConfig.BOT_ID],
      this.ctx[PluginId.Bot][StateConfigKey.botId] as string
    );
    chai.assert.equal(
      response[BaseConfig.BOT_PASSWORD],
      await getExpectedBotClientSecret(
        this.ctx,
        this.projectPath,
        this.env,
        activeResourcePlugins
      )
    );
    if (includeAAD) {
      chai.assert.equal(
        response[BaseConfig.M365_AUTHORITY_HOST],
        this.ctx[PluginId.Aad][StateConfigKey.oauthHost] as string
      );
      chai.assert.equal(
        response[BaseConfig.M365_CLIENT_ID],
        this.ctx[PluginId.Aad][StateConfigKey.clientId] as string
      );
      chai.assert.equal(
        response[BaseConfig.M365_CLIENT_SECRET],
        await getExpectedM365ClientSecret(
          this.ctx,
          this.projectPath,
          this.env,
          activeResourcePlugins
        )
      );
      chai.assert.equal(
        response[BaseConfig.M365_TENANT_ID],
        this.ctx[PluginId.Aad][StateConfigKey.tenantId] as string
      );
      chai.assert.equal(
        response[BaseConfig.M365_APPLICATION_ID_URI],
        getExpectedM365ApplicationIdUri(this.ctx, activeResourcePlugins)
      );
    }
    chai.assert.equal(
      response[BaseConfig.IDENTITY_ID],
      this.ctx[PluginId.Identity][StateConfigKey.identityClientId] as string
    );

    if (activeResourcePlugins.includes(PluginId.Function)) {
      chai.assert.equal(
        response[FunctionConfig.API_ENDPOINT],
        this.ctx[PluginId.Function][StateConfigKey.functionEndpoint] as string
      );
    }
    if (activeResourcePlugins.includes(PluginId.AzureSQL)) {
      chai.assert.equal(
        response[SQLConfig.SQL_ENDPOINT],
        this.ctx[PluginId.AzureSQL][StateConfigKey.sqlEndpoint] as string
      );
      chai.assert.equal(
        response[SQLConfig.SQL_DATABASE_NAME],
        this.ctx[PluginId.AzureSQL][StateConfigKey.databaseName] as string
      );
    }

    console.log("Successfully validate Bot Provision.");
  }

  public async validateDeploy(): Promise<void> {
    // ToDo: uncomment this function in the future.
    /*
        console.log("Start to validate Bot Deployment.");

        const tokenProvider: MockAzureAccountProvider = MockAzureAccountProvider.getInstance();
        const tokenCredential = await tokenProvider.getAccountCredentialAsync();
        const token = (await tokenCredential?.getToken())?.accessToken;

        const deployments = await this.getDeployments(this.subscriptionId, this.rg, botObject.siteName, token as string);
        const deploymentId = deployments?.[0]?.properties?.id;
        const deploymentLog = await this.getDeploymentLog(this.subscriptionId, this.rg, botObject.siteName, token as string, deploymentId!);

        chai.assert.exists(deploymentLog?.find((item: any) => item.properties.message === "Deployment successful."));
        console.log("Successfully validate Bot Deployment.");
        */
  }

  private static async getDeployments(
    subscriptionId: string,
    rg: string,
    name: string,
    token: string
  ) {
    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const getResponse = await axios.get(
        baseUrlListDeployments(subscriptionId, rg, name)
      );

      return getResponse?.data?.value;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }

  private static async getDeploymentLog(
    subscriptionId: string,
    rg: string,
    name: string,
    token: string,
    id: string
  ) {
    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const getResponse = await axios.get(
        baseUrlListDeploymentLogs(subscriptionId, rg, name, id)
      );

      return getResponse?.data?.value;
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }
}
