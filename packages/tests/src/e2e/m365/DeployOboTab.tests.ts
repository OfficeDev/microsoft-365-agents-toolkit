// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Kuojian Lu <kuojianlu@microsoft.com>
 */

import { describe } from "mocha";
import * as chai from "chai";
import * as path from "path";

import { it } from "@microsoft/extra-shot-mocha";

import m365Provider from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365LoginUserPassword";
import { AadValidator } from "../../commonlib/aadValidate";
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability } from "../../utils/constants";
import {
  cleanUpLocalProject,
  createResourceGroup,
  deleteResourceGroupByName,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
  setStaticWebAppSkuNameToStandardBicep,
} from "../commonUtils";
import {
  deleteAadAppByObjectId,
  deleteTeamsApp,
  getTeamsApp,
} from "../debug/utility";
import { StaticSiteValidator } from "../../commonlib";

describe("Deploy V3 m365-tab template", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const resourceGroupName = `${appName}-rg`;

  afterEach(async function () {
    // clean up
    const context = await readContextMultiEnvV3(projectPath, "dev");
    if (context?.TEAMS_APP_ID) {
      await deleteTeamsApp(context.TEAMS_APP_ID);
    }
    if (context?.AAD_APP_OBJECT_ID) {
      await deleteAadAppByObjectId(context.AAD_APP_OBJECT_ID);
    }
    await deleteResourceGroupByName(resourceGroupName);
    await cleanUpLocalProject(projectPath);
  });

  it(
    "happy path: provision and deploy",
    { testPlanCaseId: 17449539, author: "kuojianlu@microsoft.com" },
    async function () {
      // create
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        Capability.M365SsoLaunchPage
      );
      console.log(`[Successfully] scaffold to ${projectPath}`);

      // provision
      const result = await createResourceGroup(resourceGroupName, "westus");
      chai.assert.isTrue(result);

      // workaround free tier quota
      await setStaticWebAppSkuNameToStandardBicep(projectPath, "dev");

      await CliHelper.provisionProject(projectPath, "", "dev", {
        ...process.env,
        AZURE_RESOURCE_GROUP_NAME: resourceGroupName,
      });
      console.log(`[Successfully] provision for ${projectPath}`);

      let context = await readContextMultiEnvV3(projectPath, "dev");
      chai.assert.isDefined(context);

      // validate aad
      chai.assert.isDefined(context.AAD_APP_OBJECT_ID);
      chai.assert.isNotEmpty(context.AAD_APP_OBJECT_ID);
      const aad = AadValidator.init(context, false, m365Provider);
      await AadValidator.validate(aad);

      // validate teams app
      chai.assert.isDefined(context.TEAMS_APP_ID);
      const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
      chai.assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

      // validate m365
      chai.assert.isDefined(context.M365_TITLE_ID);
      chai.assert.isNotEmpty(context.M365_TITLE_ID);
      chai.assert.isDefined(context.M365_APP_ID);
      chai.assert.isNotEmpty(context.M365_APP_ID);

      // deploy
      await CliHelper.deployAll(projectPath, "", "dev");
      console.log(`[Successfully] deploy for ${projectPath}`);

      context = await readContextMultiEnvV3(projectPath, "dev");
      chai.assert.isDefined(context);

      // Validate deployment
      const staticSite = StaticSiteValidator.init(context);
      await StaticSiteValidator.validateDeploy(staticSite);
    }
  );
});
