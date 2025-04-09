// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import MockAzureAccountProvider, {
  AzureAccountProviderUserPassword,
} from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes } from "@microsoft/teamsfx-core";
import axios from "axios";
import * as chai from "chai";
import {
  getActivePluginsFromProjectSetting,
  getAzureAccountObjectId,
  getAzureTenantId,
  getProvisionParameterValueByKey,
} from "../e2e/commonUtils";
import { CliHelper } from "./cliHelper";
import { PluginId, StateConfigKey, provisionParametersKey } from "./constants";
import {
  getKeyVaultNameFromResourceId,
  getResourceGroupNameFromResourceId,
  getSubscriptionIdFromResourceId,
} from "./utilities";

const keyvaultScope = "https://vault.azure.net/.default";
const baseUrlVaultSecrets = (
  vaultBaseUrl: string,
  secretName: string,
  secretVersion: string
) => `${vaultBaseUrl}/secrets/${secretName}/${secretVersion}?api-version=7.2`;
const baseUrlVaults = (subscriptionId: string, rg: string, vaultName: string) =>
  `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.KeyVault/vaults/${vaultName}?api-version=2019-09-01`;
const baseUrlVaultUpdateAccessPolicy = (
  subscriptionId: string,
  rg: string,
  vaultName: string,
  updateKind: string
) =>
  `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.KeyVault/vaults/${vaultName}/accessPolicies/${updateKind}?api-version=2019-09-01`;

export interface IKeyVaultObject {
  name: string;
  vaultUri: string;
}

enum AccessPolicyUpdateKind {
  Add = "add",
  Remove = "remove",
  Replace = "replace",
}

export class KeyVaultValidator {
  private projectPath: string;
  private env: string;

  private subscriptionId: string;
  private rg: string;
  private keyVault: IKeyVaultObject;

  constructor(ctx: any, projectPath: string, env: string) {
    console.log("Start to init validator for Key Vault.");

    this.projectPath = projectPath;
    this.env = env;

    const resourceId =
      ctx[PluginId.KeyVault][StateConfigKey.keyVaultResourceId];
    chai.assert.exists(resourceId);
    this.subscriptionId = getSubscriptionIdFromResourceId(resourceId);
    chai.assert.exists(this.subscriptionId);
    this.rg = getResourceGroupNameFromResourceId(resourceId);
    chai.assert.exists(this.rg);

    const keyVaultName = getKeyVaultNameFromResourceId(resourceId);
    chai.assert.exists(keyVaultName);
    this.keyVault = {
      name: keyVaultName,
      vaultUri: `https://${keyVaultName}.vault.azure.net`,
    };
    console.log("Successfully init validator for Key Vault.");
  }

  public async validate() {
    console.log("Start to validate Key Vault.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    console.log("Validating key vault instance.");
    const keyVaultResponse = await this.getKeyVault(
      this.subscriptionId,
      this.rg,
      this.keyVault.name,
      token as string
    );
    chai.assert.exists(keyVaultResponse);

    await this.validateKeyVaultSecrets(tokenProvider, token as string);

    console.log("Successfully validate Key Vault.");
  }

  private async validateKeyVaultSecrets(
    tokenProvider: AzureAccountProviderUserPassword,
    token: string
  ) {
    console.log("Validating key vault secrets.");

    const activeResourcePlugins = await getActivePluginsFromProjectSetting(
      this.projectPath
    );
    chai.assert.isArray(activeResourcePlugins);
    if (
      !activeResourcePlugins.includes(PluginId.Aad) &&
      !activeResourcePlugins.includes(PluginId.Bot)
    ) {
      return;
    }

    // Add "get secret" permission for test account
    await this.updateKeyVaultGetSecretPermission(
      this.subscriptionId,
      this.rg,
      this.keyVault.name,
      token,
      getAzureTenantId(),
      getAzureAccountObjectId(),
      AccessPolicyUpdateKind.Add
    );

    const identityTokenCredential =
      await tokenProvider.getIdentityCredentialAsync();
    const tokenToGetSecret = (
      await identityTokenCredential?.getToken(keyvaultScope)
    )?.token;

    if (activeResourcePlugins.includes(PluginId.Aad)) {
      const m365ClientSecretName =
        (await getProvisionParameterValueByKey(
          this.projectPath,
          this.env,
          provisionParametersKey.m365ClientSecretName
        )) ?? "m365ClientSecret";
      const keyVaultSecretResponse = await this.getKeyVaultSecrets(
        this.keyVault.vaultUri,
        m365ClientSecretName,
        tokenToGetSecret as string
      );
      chai.assert.exists(keyVaultSecretResponse);
      const expectedM365ClientSecret = await CliHelper.getUserSettings(
        `${PluginId.Aad}.${StateConfigKey.clientSecret}`,
        this.projectPath,
        this.env
      );
      chai.assert.equal(keyVaultSecretResponse, expectedM365ClientSecret);
    }

    if (activeResourcePlugins.includes(PluginId.Bot)) {
      const botClientSecretName =
        (await getProvisionParameterValueByKey(
          this.projectPath,
          this.env,
          provisionParametersKey.botClientSecretName
        )) ?? "botClientSecret";
      const keyVaultSecretResponse = await this.getKeyVaultSecrets(
        this.keyVault.vaultUri,
        botClientSecretName,
        tokenToGetSecret as string
      );
      chai.assert.exists(keyVaultSecretResponse);
      const expectedBotClientSecret = await CliHelper.getUserSettings(
        `${PluginId.Bot}.${StateConfigKey.botPassword}`,
        this.projectPath,
        this.env
      );
      chai.assert.equal(keyVaultSecretResponse, expectedBotClientSecret);
    }

    // Remove "get secret" permission for test account
    await this.updateKeyVaultGetSecretPermission(
      this.subscriptionId,
      this.rg,
      this.keyVault.name,
      token,
      getAzureTenantId(),
      getAzureAccountObjectId(),
      AccessPolicyUpdateKind.Remove
    );

    console.log("Successfully validate key vault secrets.");
  }

  private async getKeyVaultSecrets(
    vaultUri: string,
    secretName: string,
    token: string,
    secretVersion = ""
  ) {
    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const getResponse = await axios.get(
        baseUrlVaultSecrets(vaultUri, secretName, secretVersion)
      );

      if (getResponse && getResponse.data && getResponse.data.value) {
        return getResponse.data.value;
      }
    } catch (error) {
      console.log(error);
    }

    return undefined;
  }

  private async updateKeyVaultGetSecretPermission(
    subscriptionId: string,
    rg: string,
    keyVaultName: string,
    token: string,
    tenantId: string,
    objectId: string,
    updateKind: AccessPolicyUpdateKind
  ) {
    console.log(
      `${updateKind} key vault "get secret" permission for ${objectId} on key vault ${keyVaultName}, subscription id: ${subscriptionId}.`
    );

    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const body = {
        properties: {
          accessPolicies: [
            {
              tenantId: `${tenantId}`,
              objectId: `${objectId}`,
              permissions: {
                secrets: ["get"],
              },
            },
          ],
        },
      };
      const getResponse = await axios.put(
        baseUrlVaultUpdateAccessPolicy(
          subscriptionId,
          rg,
          keyVaultName,
          updateKind
        ),
        body
      );
      chai.assert.equal(getResponse.status, 200);
      console.log(
        `Successfully ${updateKind} key vault "get secret" permission`
      );
    } catch (error) {
      console.log(error);
    }
  }

  private async getKeyVault(
    subscriptionId: string,
    rg: string,
    keyVaultName: string,
    token: string
  ) {
    try {
      axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      const getResponse = await axios.get(
        baseUrlVaults(subscriptionId, rg, keyVaultName)
      );
      if (getResponse && getResponse.data && getResponse.data.properties) {
        return getResponse.data.properties;
      }
    } catch (error) {
      console.log(error);
      return undefined;
    }
  }
}
