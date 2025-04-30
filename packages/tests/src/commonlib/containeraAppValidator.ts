// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AzureScopes } from "@microsoft/teamsfx-core";
import * as chai from "chai";

import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { EnvConstants } from "./constants";

import { Env } from "../utils/env";
import { Executor } from "../utils/executor";
import {
  getContainerAppProperties,
  getResourceGroupNameFromResourceId,
  getSubscriptionIdFromResourceId,
} from "./utilities";

export class ContainerAppValidator {
  private ctx: any;
  private subscriptionId: string;
  private rg: string;
  private containerAppNames: string[];

  constructor(ctx: any) {
    console.log("Start to init validator for Azure Container App.");

    this.ctx = ctx;

    const resourceId = ctx[EnvConstants.AZURE_CONTAINER_APP_RESOURCE_ID];
    chai.assert.exists(resourceId);
    this.subscriptionId = getSubscriptionIdFromResourceId(resourceId);
    chai.assert.exists(this.subscriptionId);
    this.rg = getResourceGroupNameFromResourceId(resourceId);
    chai.assert.exists(this.rg);
    this.containerAppNames = [
      EnvConstants.AZURE_CONTAINER_APP_NAME,
      EnvConstants.BACKEND_APP_NAME,
      EnvConstants.FRONTEND_APP_NAME,
    ]
      .map((name) => this.ctx[name])
      .filter((value) => !!value);

    chai.assert.isTrue(
      this.containerAppNames && this.containerAppNames.length > 0
    );

    console.log("Successfully init validator for Azure Container App.");
  }

  public async validateProvision(includeAAD = true): Promise<void> {
    console.log("Start to validate Azure Container App Provision.");

    const tokenProvider = MockAzureAccountProvider;
    const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
    const token = (await tokenCredential?.getToken(AzureScopes))?.token;

    for (const containerAppName of this.containerAppNames) {
      const response = await getContainerAppProperties(
        this.subscriptionId,
        this.rg,
        containerAppName,
        token as string
      );
      chai.assert.exists(
        response,
        `Response for ${containerAppName} should exist`
      );
    }

    console.log("Successfully validate Azure Container App Provision.");
  }

  public async validateContainerAppStatus(): Promise<void> {
    for (const containerAppName of this.containerAppNames) {
      const command = `az containerapp show --name ${containerAppName} --resource-group ${Env["azureResourceGroup"]} --subscription ${Env["azureSubscriptionId"]}`;
      const { stdout, success } = await Executor.execute(
        command,
        process.cwd()
      );
      chai.assert.isTrue(success);
      const result = JSON.parse(stdout);
      const status = result.properties?.runningStatus;
      chai.assert.strictEqual(status, "Running", "Status should be 'Running'");
    }
  }
}
