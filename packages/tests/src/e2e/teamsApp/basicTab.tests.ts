// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Zhijie Huang <zhihuan@microsoft.com>, Ning Tang
 */

import { it } from "@microsoft/extra-shot-mocha";
import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes, environmentNameManager } from "@microsoft/teamsfx-core";
import axios from "axios";
import { assert } from "chai";
import fs from "fs-extra";
import path from "path";
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
  removeTeamsAppExtendToM365,
} from "../commonUtils";
import { getTeamsApp } from "../debug/utility";

describe("Basic Tab", function () {
  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  const appName = getUniqueAppName();
  const resourceGroupName = `${appName}-rg`;
  const projectPath = path.resolve(testFolder, appName);
  const envName = environmentNameManager.getDefaultEnvName();

  afterEach(async () => {
    await cleanUp(appName, projectPath, false, false, false);
  });

  it(
    "typescript template",
    {
      testPlanCaseId: 24137515,
      author: "Ning.Tang@microsoft.com",
    },
    async function () {
      // Scaffold
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        Capability.TabNonSso,
        process.env,
        `--programming-language typescript`
      );

      // Validate Scaffold
      const indexFile = path.join(projectPath, "src", "index.ts");
      fs.access(indexFile, fs.constants.F_OK, (err) => {
        assert.notExists(err, "index.ts should exist");
      });

      // remove teamsApp/extendToM365 in case it fails
      removeTeamsAppExtendToM365(path.join(projectPath, "m365agents.yml"));
      removeTeamsAppExtendToM365(
        path.join(projectPath, "m365agents.local.yml")
      );

      // Local Debug (Provision)
      await CliHelper.provisionProject(projectPath, "", "local");
      console.log(`[Successfully] provision for ${projectPath}`);

      let context = await readContextMultiEnvV3(projectPath, "local");
      assert.isDefined(context, "local env file should exist");

      // validate aad
      assert.isUndefined(context.AAD_APP_OBJECT_ID, "AAD should not exist");

      // validate teams app
      assert.isDefined(context.TEAMS_APP_ID, "teams app id should be defined");
      const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
      assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

      // Local Debug (Deploy)
      await CliHelper.deployAll(projectPath, "", "local");
      console.log(`[Successfully] deploy for ${projectPath}`);

      context = await readContextMultiEnvV3(projectPath, "local");
      assert.isDefined(context);

      // validate ssl cert
      assert.isDefined(context.SSL_CRT_FILE, "ssl cert should be defined");
      assert.isNotEmpty(context.SSL_CRT_FILE);
      assert.isDefined(context.SSL_KEY_FILE, "ssl key should be defined");
      assert.isNotEmpty(context.SSL_KEY_FILE);

      // validate .localConfigs
      assert.isTrue(
        await fs.pathExists(path.join(projectPath, ".localConfigs")),
        ".localConfigs should exist"
      );

      // Remote Provision
      const result = await createResourceGroup(resourceGroupName, "westus");
      assert.isTrue(
        result,
        `failed to create resource group: ${resourceGroupName}`
      );

      await CliHelper.provisionProject(projectPath, "", "dev", {
        ...process.env,
        AZURE_RESOURCE_GROUP_NAME: resourceGroupName,
      });

      context = await readContextMultiEnvV3(projectPath, envName);
      assert.exists(context, "env file should exist");

      const appServiceResourceId =
        context[EnvConstants.TAB_AZURE_APP_SERVICE_RESOURCE_ID];
      assert.exists(
        appServiceResourceId,
        "Azure App Service resource ID should exist"
      );

      const tokenProvider = MockAzureAccountProvider;
      const tokenCredential = await tokenProvider.getIdentityCredentialAsync();
      const token = (await tokenCredential?.getToken(AzureScopes))?.token;
      assert.exists(token);

      const response = await getWebappSettings(
        subscription,
        getResourceGroupNameFromResourceId(appServiceResourceId),
        getSiteNameFromResourceId(appServiceResourceId),
        token as string
      );
      assert.exists(response, "Web app settings should exist");
      assert.equal(
        response["WEBSITE_NODE_DEFAULT_VERSION"],
        "~22",
        "Node version should be 22"
      );
      assert.equal(
        response["WEBSITE_RUN_FROM_PACKAGE"],
        "1",
        "Run from package should be 1"
      );
      assert.equal(
        response["RUNNING_ON_AZURE"],
        "1",
        "Running on azure should be 1"
      );

      // Remote Deploy
      await CliHelper.deployAll(projectPath);

      // Validate Deploy
      context = await readContextMultiEnvV3(projectPath, envName);
      assert.exists(context, "env file should exist");

      const endpoint = context[EnvConstants.TAB_ENDPOINT];
      assert.exists(endpoint, "Tab endpoint should exist");

      const axiosInstance = axios.create();
      try {
        const response = await axiosInstance.get(endpoint);
        assert.equal(response.status, 200, "tab endpoint is not reachable");
      } catch (e) {
        assert.fail(JSON.stringify(e));
      }
    }
  );
});
