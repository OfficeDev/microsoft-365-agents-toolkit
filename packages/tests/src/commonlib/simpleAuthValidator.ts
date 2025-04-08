// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import { PluginId, StateConfigKey, provisionParametersKey } from "./constants";
import {
  getExpectedM365ApplicationIdUri,
  getResourceGroupNameFromResourceId,
  getSubscriptionIdFromResourceId,
  getWebappServicePlan,
  getWebappSettings,
  getActivePluginsFromProjectSetting,
  getProvisionParameterValueByKey,
} from "./utilities";
import { getExpectedM365ClientSecret } from "./cliHelper";

export class PropertiesKeys {
  static clientId = "CLIENT_ID";
  static clientSecret = "CLIENT_SECRET";
  static oauthAuthority = "OAUTH_AUTHORITY";
  static identifierUri = "IDENTIFIER_URI";
  static aadMetadataAddreass = "AAD_METADATA_ADDRESS";
  static tabAppEndpoint = "TAB_APP_ENDPOINT";
}

export interface ISimpleAuthObject {
  endpoint: string;
  webAppResourceId?: string;
}

export class SimpleAuthValidator {
  private ctx: any;
  private projectPath: string;
  private env: string;

  private subscriptionId: string;
  private rg: string;
  private simpleAuthObject: ISimpleAuthObject;

  constructor(ctx: any, projectPath: string, env: string) {
    console.log("Start to init validator for Simple Auth.");

    this.ctx = ctx;
    this.projectPath = projectPath;
    this.env = env;

    const resourceId =
      ctx[PluginId.SimpleAuth][StateConfigKey.webAppResourceId];
    chai.assert.exists(resourceId);
    this.subscriptionId = getSubscriptionIdFromResourceId(resourceId);
    chai.assert.exists(this.subscriptionId);
    this.rg = getResourceGroupNameFromResourceId(resourceId);
    chai.assert.exists(this.rg);
    this.simpleAuthObject = {
      endpoint: ctx[PluginId.SimpleAuth][StateConfigKey.endpoint],
      webAppResourceId:
        ctx[PluginId.SimpleAuth][StateConfigKey.webAppResourceId],
    };

    console.log("Successfully init validator for Simple Auth.");
  }

  public async validate() {
    console.log("Start to validate Simple Auth.");

    const resourceName: string = this.simpleAuthObject.endpoint.slice(8, -18);
    chai.assert.exists(resourceName);

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    console.log("Validating app settings.");
    const activeResourcePlugins = await getActivePluginsFromProjectSetting(
      this.projectPath
    );
    chai.assert.isArray(activeResourcePlugins);
    const response = await getWebappSettings(
      this.subscriptionId,
      this.rg,
      resourceName,
      token as string
    );
    chai.assert.exists(response);
    chai.assert.equal(
      response[PropertiesKeys.clientId],
      this.ctx[PluginId.Aad][StateConfigKey.clientId]
    );
    chai.assert.equal(
      response[PropertiesKeys.clientSecret],
      await getExpectedM365ClientSecret(
        this.ctx,
        this.projectPath,
        this.env,
        activeResourcePlugins
      )
    );
    chai.assert.equal(
      response[PropertiesKeys.identifierUri],
      getExpectedM365ApplicationIdUri(this.ctx, activeResourcePlugins)
    );
    chai.assert.equal(
      response[PropertiesKeys.oauthAuthority],
      this.ctx[PluginId.Aad][StateConfigKey.oauthAuthority]
    );
    chai.assert.equal(
      response[PropertiesKeys.aadMetadataAddreass],
      `${
        this.ctx[PluginId.Aad][StateConfigKey.oauthAuthority]
      }/v2.0/.well-known/openid-configuration`
    );
    if (activeResourcePlugins.includes(PluginId.FrontendHosting)) {
      chai.assert.equal(
        response[PropertiesKeys.tabAppEndpoint],
        this.ctx[PluginId.FrontendHosting][StateConfigKey.endpoint]
      );
    }

    console.log("Validating app service plan.");
    const servicePlanName = resourceName.replace("-webapp", "-serverfarms");
    const serivcePlanResponse = await getWebappServicePlan(
      this.subscriptionId,
      this.rg,
      servicePlanName,
      token as string
    );
    const expectedServicePlan =
      (await getProvisionParameterValueByKey(
        this.projectPath,
        this.env,
        provisionParametersKey.simpleAuthSku
      )) ?? "F1";
    chai.assert(serivcePlanResponse, expectedServicePlan);

    console.log("Successfully validate Simple Auth.");
  }
}
