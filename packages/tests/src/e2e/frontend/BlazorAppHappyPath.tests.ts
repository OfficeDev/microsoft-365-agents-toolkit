// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhijie Huang <zhijie.huang@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes } from "@microsoft/teamsfx-core";
import { environmentNameManager } from "@microsoft/teamsfx-core/build/core/environmentName";
import axios from "axios";
import * as chai from "chai";
import fs from "fs-extra";
import { describe } from "mocha";
import path from "path";
import { FrontendWebAppConfig } from "../../commonlib";
import { CliHelper } from "../../commonlib/cliHelper";
import { EnvConstants } from "../../commonlib/constants";
import {
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getWebappSettings,
} from "../../commonlib/utilities";
import { Capability } from "../../utils/constants";
import {
  cleanUp,
  createResourceGroup,
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
  setProvisionParameterValueV3,
} from "../commonUtils";

describe("Blazor App", function () {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const subscription = getSubscriptionId();
  const projectPath = path.resolve(testFolder, appName);
  const envName = environmentNameManager.getDefaultEnvName();
  const resourceGroupName = `${appName}-rg`;
  const env = Object.assign({}, process.env);
  env["TEAMSFX_CLI_DOTNET"] = "true";

  after(async () => {
    // clean up
    await cleanUp(appName, projectPath, false, false, false);
  });
  it(
    `Create Blazor app`,
    { testPlanCaseId: 15686028, author: "zhijie.huang@microsoft.com" },
    async () => {
      await CliHelper.createDotNetProject(
        appName,
        testFolder,
        Capability.TabNonSso,
        env
      );
      const programCsPath = path.join(
        testFolder,
        appName,
        "Components",
        "App.razor"
      );
      chai.assert.isTrue(await fs.pathExists(programCsPath));
    }
  );

  it(
    `Provision Resource`,
    { testPlanCaseId: 15686030, author: "zhijie.huang@microsoft.com" },
    async () => {
      const result = await createResourceGroup(resourceGroupName, "westus");
      chai.assert.isTrue(result);

      await setProvisionParameterValueV3(projectPath, envName, {
        key: "webAppSKU",
        value: "B1",
      });
      await CliHelper.provisionProject(projectPath, "", envName as "dev", {
        ...env,
        AZURE_RESOURCE_GROUP_NAME: resourceGroupName,
      });

      const tokenProvider = MockAzureAccountProvider;
      const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
      const token = (await tokenCredential?.getToken(AzureScopes))?.token;
      chai.assert.exists(token);

      const context = await readContextMultiEnvV3(projectPath, envName);
      const resourceId =
        context[EnvConstants.TAB_AZURE_APP_SERVICE_RESOURCE_ID];
      chai.assert.exists(context);
      chai.assert.exists(resourceId);
      const response = await getWebappSettings(
        subscription,
        getResourceGroupNameFromResourceId(resourceId),
        getSiteNameFromResourceId(resourceId),
        token as string
      );
      chai.assert.exists(response);
      chai.assert.equal(
        response[FrontendWebAppConfig.clientId],
        context[EnvConstants.AAD_APP_CLIENT_ID]
      );
      chai.assert.equal(
        response[FrontendWebAppConfig.authority],
        context[EnvConstants.AAD_APP_OAUTH_AUTHORITY]
      );
    }
  );

  it(
    "Deploy Blazor app to Azure Web APP",
    { testPlanCaseId: 15686031, author: "zhijie.huang@microsoft.com" },
    async () => {
      await CliHelper.deployAll(projectPath, "", envName as "dev", env);

      const context = await readContextMultiEnvV3(projectPath, envName);
      const endpoint = context[EnvConstants.TAB_ENDPOINT];
      chai.assert.exists(endpoint);
      const axiosInstance = axios.create();
      try {
        // wait until the web app starts
        setTimeout(async () => {
          const response = await axiosInstance.get(endpoint);
          chai.assert.equal(response.status, 200);
        }, 30000);
      } catch (e) {
        chai.assert.notExists(e);
      }
    }
  );
});
