// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Wenyu Tang <wenyutang@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { expect } from "chai";
import fs from "fs-extra";
import path from "path";
import M365Login from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365Login";
import { AadValidator } from "../../commonlib";
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability } from "../../utils/constants";
import {
  cleanUp,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
  setAadManifestIdentifierUrisV3,
  createResourceGroup,
  setStaticWebAppSkuNameToStandardBicep,
} from "../commonUtils";
import { Executor } from "../../utils/executor";

describe("SSO Tab with aad manifest enabled", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  const env = Object.assign({}, process.env);
  env["TEAMSFX_AAD_MANIFEST"] = "true";
  env["TEAMSFX_CONFIG_UNIFY"] = "true";

  after(async () => {
    await cleanUp(appName, projectPath, true, false, false);
  });

  it(
    "SSO Tab E2E test with aad manifest enabled",
    { testPlanCaseId: 24137775, author: "wenyutang@microsoft.com" },
    async () => {
      // Arrange
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        Capability.M365SsoLaunchPage,
        env
      );
      // Assert
      expect(fs.pathExistsSync(path.join(projectPath, "infra", "azure.bicep")))
        .to.be.true;
      expect(
        fs.pathExistsSync(
          path.join(projectPath, "infra", "azure.parameters.json")
        )
      ).to.be.true;
      expect(fs.pathExistsSync(path.join(projectPath, "teamsapp.yml"))).to.be
        .true;
      expect(fs.pathExistsSync(path.join(projectPath, "aad.manifest.json"))).to
        .be.true;

      {
        // provision
        const result = await createResourceGroup(appName + "-rg", "westus");
        expect(result).to.be.true;
        process.env["AZURE_RESOURCE_GROUP_NAME"] = appName + "-rg";
        // workaround free tier quota
        await setStaticWebAppSkuNameToStandardBicep(projectPath, "dev");
        const { success } = await Executor.provision(projectPath);
        expect(success).to.be.true;
        console.log(`[Successfully] provision for ${projectPath}`);
      }

      const context = await readContextMultiEnvV3(projectPath, "dev");

      // Validate Aad App
      const aad = AadValidator.init(context, false, M365Login);
      await AadValidator.validate(aad);

      const firstIdentifierUri =
        "api://first.com/291fc1b5-1146-4d33-b7b8-ec4c441b6b33";
      await setAadManifestIdentifierUrisV3(projectPath, firstIdentifierUri);

      {
        // Deploy all resources without aad manifest
        const { success } = await Executor.provision(projectPath);
        expect(success).to.be.true;
      }
      await AadValidator.validate(aad);

      // Deploy all resources include aad manifest
      await CliHelper.updateAadManifest(projectPath, "--env dev", env);
      await AadValidator.validate(aad, firstIdentifierUri);

      const secondIdentifierUri =
        "api://second.com/291fc1b5-1146-4d33-b7b8-ec4c441b6b33";
      await setAadManifestIdentifierUrisV3(projectPath, secondIdentifierUri);

      // Only deploy aad manifest
      await CliHelper.updateAadManifest(projectPath, "--env dev", env);
      await AadValidator.validate(aad, secondIdentifierUri);
    }
  );
});
