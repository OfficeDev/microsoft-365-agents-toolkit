// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Quke <quke@microsoft.com>
 */

import { describe } from "mocha";
import * as path from "path";

import { it } from "@microsoft/extra-shot-mocha";

import MockAzureAccountProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/azureLoginUserPassword";
import { AzureScopes, environmentNameManager } from "@microsoft/teamsfx-core";
import { assert } from "chai";
import fs from "fs-extra";
import { CliHelper } from "../../commonlib/cliHelper";
import { EnvConstants } from "../../commonlib/constants";
import {
  getBotServiceProperties,
  getResourceGroupNameFromResourceId,
  getSiteNameFromResourceId,
  getWebappSettings,
} from "../../commonlib/utilities";
import { Capability } from "../../utils/constants";
import {
  cleanUpLocalProject,
  createResourceGroup,
  deleteResourceGroupByName,
  getSubscriptionId,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
} from "../commonUtils";
import {
  deleteAadAppByClientId,
  deleteBot,
  deleteTeamsApp,
  getAadAppByClientId,
  getTeamsApp,
} from "../debug/utility";

describe("Foundry Proxy Agent for csharp version", function () {
  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  const appName = getUniqueAppName();
  const resourceGroupName = `${appName}-rg`;
  const localResourceGroupName = `${appName}-local-rg`;
  const projectPath = path.resolve(testFolder, appName);
  const envName = environmentNameManager.getDefaultEnvName();

  after(async () => {
    // clean up
    let context = await readContextMultiEnvV3(projectPath, "local");
    if (context?.TEAMS_APP_ID) {
      await deleteTeamsApp(context.TEAMS_APP_ID);
    }
    if (context?.BOT_ID) {
      await deleteBot(context.BOT_ID);
      await deleteAadAppByClientId(context.BOT_ID);
    }

    context = await readContextMultiEnvV3(projectPath, "dev");
    if (context?.TEAMS_APP_ID) {
      await deleteTeamsApp(context.TEAMS_APP_ID);
    }
    await deleteResourceGroupByName(localResourceGroupName);
    await deleteResourceGroupByName(resourceGroupName);
    await cleanUpLocalProject(projectPath);
  });

  describe("Declarative Agent", function () {
    it(
      "Validate scaffold and manifest for declarative agent capabilities",
      {
        testPlanCaseId: 35527260,
        author: "quke@microsoft.com",
      },
      async function () {
        // Scaffold
        const env = Object.assign({}, process.env);
        env["TEAMSFX_CLI_DOTNET"] = "true";
        const options = `--foundry-endpoint https://test.ai.azure.com --foundry-agent-id test-agent-id`;
        await CliHelper.createDotNetProject(
          appName,
          testFolder,
          Capability.FoundryProxyAgent,
          env,
          options
        );

        // Validate Scaffold - Program.cs
        const programFile = path.join(testFolder, appName, "Program.cs");
        assert.isTrue(
          await fs.pathExists(programFile),
          "Program.cs should exist"
        );

        // Validate Scaffold - FoundryAgent.cs
        const agentFile = path.join(
          testFolder,
          appName,
          "Agents",
          "FoundryAgent.cs"
        );
        assert.isTrue(
          await fs.pathExists(agentFile),
          "Agents/FoundryAgent.cs should exist"
        );

        // Validate Scaffold - manifest.json
        const manifestFile = path.join(
          testFolder,
          appName,
          "appPackage",
          "manifest.json"
        );
        assert.isTrue(
          await fs.pathExists(manifestFile),
          "appPackage/manifest.json should exist"
        );
        const manifest = await fs.readJson(manifestFile);
        assert.isDefined(
          manifest.copilotAgents,
          "manifest should contain copilotAgents"
        );
        assert.isDefined(
          manifest.copilotAgents?.customEngineAgents,
          "manifest should contain customEngineAgents"
        );
        assert.isArray(
          manifest.copilotAgents?.customEngineAgents,
          "customEngineAgents should be an array"
        );
        assert.isAbove(
          manifest.copilotAgents?.customEngineAgents?.length,
          0,
          "customEngineAgents should have at least one entry"
        );

        // Validate Scaffold - appsettings.json
        const appsettingsFile = path.join(
          testFolder,
          appName,
          "appsettings.json"
        );
        assert.isTrue(
          await fs.pathExists(appsettingsFile),
          "appsettings.json should exist"
        );
        const appsettings = await fs.readJson(appsettingsFile);
        assert.isDefined(
          appsettings.AIServices,
          "appsettings should contain AIServices section"
        );
        assert.isDefined(
          appsettings.AIServices?.AzureAIFoundryProjectEndpoint,
          "AIServices should contain AzureAIFoundryProjectEndpoint"
        );
        assert.isDefined(
          appsettings.AIServices?.AgentID,
          "AIServices should contain AgentID"
        );

        // Validate Scaffold - infra files for Azure deploy
        const azureBicepFile = path.join(
          testFolder,
          appName,
          "infra",
          "azure.bicep"
        );
        assert.isTrue(
          await fs.pathExists(azureBicepFile),
          "infra/azure.bicep should exist"
        );

        const azureLocalBicepFile = path.join(
          testFolder,
          appName,
          "infra",
          "azure-local.bicep"
        );
        assert.isTrue(
          await fs.pathExists(azureLocalBicepFile),
          "infra/azure-local.bicep should exist"
        );

        // Validate m365agents yml files
        const m365agentsYmlFile = path.join(
          testFolder,
          appName,
          "m365agents.yml"
        );
        assert.isTrue(
          await fs.pathExists(m365agentsYmlFile),
          "m365agents.yml should exist"
        );

        const m365agentsLocalYmlFile = path.join(
          testFolder,
          appName,
          "m365agents.local.yml"
        );
        assert.isTrue(
          await fs.pathExists(m365agentsLocalYmlFile),
          "m365agents.local.yml should exist"
        );
      }
    );
  });

  describe("Teams Agent", function () {
    it(
      "Provision and deploy foundry proxy agent",
      {
        testPlanCaseId: 35527261,
        author: "quke@microsoft.com",
      },
      async function () {
        // Create resource group for local provision (needed for Bot Service + OAuth)
        const localResourceGroupName = `${appName}-local-rg`;
        const localRgResult = await createResourceGroup(
          localResourceGroupName,
          "australiacentral"
        );
        assert.isTrue(
          localRgResult,
          `failed to create local resource group: ${localResourceGroupName}`
        );

        // Local Debug (Provision)
        await CliHelper.provisionProject(projectPath, "", "local", {
          ...process.env,
          BOT_DOMAIN: "test.ngrok.io",
          BOT_ENDPOINT: "https://test.ngrok.io",
          AZURE_RESOURCE_GROUP_NAME: localResourceGroupName,
        });
        console.log(`[Successfully] local provision for ${projectPath}`);

        let context = await readContextMultiEnvV3(projectPath, "local");
        assert.isDefined(context, "local env file should exist");

        // validate teams app
        assert.isDefined(
          context.TEAMS_APP_ID,
          "teams app id should be defined"
        );
        const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
        assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

        // validate bot
        assert.isDefined(context.BOT_ID);
        assert.isNotEmpty(context.BOT_ID);
        const aadApp = await getAadAppByClientId(context.BOT_ID);
        assert.isDefined(aadApp);
        assert.equal(aadApp?.appId, context.BOT_ID);

        // Validate bot service via ARM (same-tenant account required)
        const localTokenProvider = MockAzureAccountProvider;
        const localCredential =
          await localTokenProvider.getIdentityCredentialAsync();
        const localToken = (await localCredential?.getToken(AzureScopes()))
          ?.token;
        assert.exists(
          localToken,
          "ARM token should be available for bot service validation"
        );
        const botService = await getBotServiceProperties(
          subscription,
          localResourceGroupName,
          localToken as string,
          context.BOT_ID
        );
        assert.isDefined(botService, "Bot service should exist in ARM");
        assert.equal(botService?.msaAppId, context.BOT_ID);
        assert.equal(
          botService?.endpoint,
          "https://test.ngrok.io/api/messages"
        );

        // validate SSO app
        assert.isDefined(
          context.SSO_APP_ID,
          "SSO app id should be defined for foundry proxy agent"
        );

        // Remote Provision
        const result = await createResourceGroup(
          resourceGroupName,
          "australiacentral"
        );
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

        // validate teams app
        assert.isDefined(context.TEAMS_APP_ID);
        const remoteTeamsApp = await getTeamsApp(context.TEAMS_APP_ID);
        assert.equal(remoteTeamsApp?.teamsAppId, context.TEAMS_APP_ID);

        const appServiceResourceId =
          context[EnvConstants.BOT_AZURE_APP_SERVICE_RESOURCE_ID];
        assert.exists(
          appServiceResourceId,
          "Azure App Service resource ID should exist"
        );

        const tokenProvider = MockAzureAccountProvider;
        const tokenCredential =
          await tokenProvider.getIdentityCredentialAsync();
        const token = (await tokenCredential?.getToken(AzureScopes()))?.token;
        assert.exists(token);

        const response = await getWebappSettings(
          subscription,
          getResourceGroupNameFromResourceId(appServiceResourceId),
          getSiteNameFromResourceId(appServiceResourceId),
          token as string
        );
        assert.exists(response, "Web app settings should exist");
        assert.equal(
          response["WEBSITE_RUN_FROM_PACKAGE"],
          "1",
          "Run from package should be 1"
        );
        assert.equal(
          response["ASPNETCORE_ENVIRONMENT"],
          "Production",
          "ASPNETCORE_ENVIRONMENT should be Production"
        );

        // Remote Deploy
        await CliHelper.deployAll(projectPath);

        // Validate Deploy
        context = await readContextMultiEnvV3(projectPath, envName);
        assert.exists(context, "env file should exist");
      }
    );
  });
});
