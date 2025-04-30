// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Zhijie Huang <zhihuan@microsoft.com>
 */

import { assert } from "chai";
import fs from "fs-extra";
import path from "path";
import { it } from "@microsoft/extra-shot-mocha";
import {
  getTestFolder,
  cleanUp,
  getSubscriptionId,
  getUniqueAppName,
  readContextMultiEnvV3,
  setProvisionParameterValueV3,
  removeTeamsAppExtendToM365,
  createResourceGroup,
} from "../commonUtils";
import { EnvConstants } from "../../commonlib/constants";
import { Capability } from "../../utils/constants";
import { CliHelper } from "../../commonlib/cliHelper";
import { AzureScopes, environmentNameManager } from "@microsoft/teamsfx-core";
import {
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getWebappSettings,
} from "../../commonlib/utilities";
import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import axios from "axios";

describe("Basic Tab", function () {
  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  const appName = getUniqueAppName();
  const resourceGroupName = `${appName}-rg`;
  const projectPath = path.resolve(testFolder, appName);
  const envName = environmentNameManager.getDefaultEnvName();
  afterEach(async () => {
    await cleanUp(appName, projectPath, true, false, false);
  });

  const files = ["app.js", "app.ts"];
  const languages = ["javascript", "typescript"];
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const language = languages[i];
    it(
      `Create & Provision & Deploy Basic Tab (${language})`,
      {
        testPlanCaseId: language === "javascript" ? 9426074 : 24137515,
        author: "zhihuan@microsoft.com",
      },
      async function () {
        // Scaffold
        await CliHelper.createProjectWithCapability(
          appName,
          testFolder,
          Capability.TabNonSso,
          process.env,
          `--programming-language ${language}`
        );

        // Validate Scaffold
        const indexFile = path.join(projectPath, "src", file);
        fs.access(indexFile, fs.constants.F_OK, (err) => {
          assert.notExists(err);
        });

        // remove teamsApp/extendToM365 in case it fails
        removeTeamsAppExtendToM365(path.join(projectPath, "teamsapp.yml"));

        // Provision
        const result = await createResourceGroup(resourceGroupName, "westus");
        assert.isTrue(result);

        await setProvisionParameterValueV3(projectPath, envName, {
          key: "webAppSku",
          value: "B1",
        });
        await CliHelper.provisionProject(projectPath, "", envName as "dev", {
          ...process.env,
          AZURE_RESOURCE_GROUP_NAME: resourceGroupName,
        });

        // Validate Provision
        let context = await readContextMultiEnvV3(projectPath, envName);
        assert.exists(context);

        const resourceId =
          context[EnvConstants.TAB_AZURE_APP_SERVICE_RESOURCE_ID];
        assert.exists(resourceId);

        const tokenProvider = MockAzureAccountProvider;
        const tokenCredential =
          await tokenProvider.getIdentityCredentialAsync();
        const token = (await tokenCredential?.getToken(AzureScopes))?.token;
        assert.exists(token);

        const response = await getWebappSettings(
          subscription,
          getResourceGroupNameFromResourceId(resourceId),
          getSiteNameFromResourceId(resourceId),
          token as string
        );
        assert.exists(response);
        assert.equal(response["WEBSITE_NODE_DEFAULT_VERSION"], "~18");
        assert.equal(response["WEBSITE_RUN_FROM_PACKAGE"], "1");
        assert.equal(response["RUNNING_ON_AZURE"], "1");

        // deploy
        await CliHelper.deployAll(projectPath);

        // Validate Deploy
        context = await readContextMultiEnvV3(projectPath, envName);
        assert.exists(context);

        const endpoint = context[EnvConstants.TAB_ENDPOINT];
        assert.exists(endpoint);

        const axiosInstance = axios.create();
        try {
          const response = await axiosInstance.get(endpoint);
          assert.equal(response.status, 200);
        } catch (e) {
          assert.fail(JSON.stringify(e));
        }
      }
    );
  }
});
