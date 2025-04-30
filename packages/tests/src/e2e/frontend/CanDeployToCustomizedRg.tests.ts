// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhijie Huang <zhijie.huang@microsoft.com>
 */

import path from "path";

import { it } from "@microsoft/extra-shot-mocha";
import { environmentNameManager } from "@microsoft/teamsfx-core";
import { describe } from "mocha";
import M365Login from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365Login";
import { AadValidator, StaticSiteValidator } from "../../commonlib";
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability } from "../../utils/constants";
import {
  cleanUp,
  createResourceGroup,
  deleteResourceGroupByName,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
  removeTeamsAppExtendToM365,
  setStaticWebAppSkuNameToStandardBicep,
} from "../commonUtils";

describe("Deploy to customized resource group", function () {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const env = environmentNameManager.getDefaultEnvName();

  after(async () => {
    await cleanUp(appName, projectPath, true, false, false);
  });

  it(
    `tab project can deploy frontend hosting resource to customized resource group and successfully provision / deploy`,
    { testPlanCaseId: 17449539, author: "zhijie.huang@microsoft.com" },
    async function () {
      // Create new tab project
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        Capability.M365SsoLaunchPage
      );

      // remove teamsApp/extendToM365 in case it fails
      removeTeamsAppExtendToM365(path.join(projectPath, "teamsapp.yml"));

      // Create empty resource group
      const customizedRgName = `${appName}-customized-rg`;
      await createResourceGroup(customizedRgName, "westus");

      // workaround free tier quota
      await setStaticWebAppSkuNameToStandardBicep(projectPath, "dev");

      await CliHelper.provisionProject(projectPath, undefined, "dev", {
        ...process.env,
        AZURE_RESOURCE_GROUP_NAME: customizedRgName,
      });
      await CliHelper.deployAll(projectPath);

      // Assert
      {
        const context = await readContextMultiEnvV3(projectPath, env);

        // Validate Aad App
        const aad = AadValidator.init(context, false, M365Login);
        await AadValidator.validate(aad);

        // Validate deployment
        const staticSite = StaticSiteValidator.init(context);
        await StaticSiteValidator.validateDeploy(staticSite);
      }

      await deleteResourceGroupByName(customizedRgName);
    }
  );
});
