// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhaofeng Xu <zhaofengxu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import * as chai from "chai";
import { describe } from "mocha";
import * as path from "path";
import { Cleaner } from "../../../commonlib/cleaner";
import { CliHelper } from "../../../commonlib/cliHelper";
import { TemplateProjectFolder } from "../../../utils/constants";
import { Executor } from "../../../utils/executor";
import {
  createResourceGroup,
  getTestFolder,
  getUniqueAppName,
} from "../../commonUtils";
import { checkYmlHeader } from "./utils";
// As there is build errors for v4 template prject, skip the test for now
describe.skip("upgrade", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    await Cleaner.clean(projectPath);
  });

  it(
    "upgrade project",
    { testPlanCaseId: 17184119, author: "zhaofengxu@microsoft.com" },
    async function () {
      {
        await Executor.installCLI(testFolder, "1.2.5", false);
        const env = Object.assign({}, process.env);
        // new a project ( tab only )
        await CliHelper.createTemplateProject(
          appName,
          testFolder,
          TemplateProjectFolder.HelloWorldTabBackEnd,
          env,
          true,
          true,
          false
        );
      }

      {
        // upgrade
        const result = await Executor.upgrade(projectPath);
        chai.assert.isTrue(result.success);
        const ymlFile = path.join(projectPath, "teamsapp.yml");
        await checkYmlHeader(ymlFile);
      }

      // {
      //   // preview
      //   const result = await Executor.preview(projectPath);
      //   chai.assert.isTrue(result.success);
      // }

      {
        // provision
        const rgResult = await createResourceGroup(appName + "-rg", "westus");
        chai.assert.isTrue(rgResult);
        process.env["AZURE_RESOURCE_GROUP_NAME"] = appName + "-rg";
        const result = await Executor.provision(projectPath);
        chai.assert.isTrue(result.success);
      }

      {
        // deploy
        const result = await Executor.deploy(projectPath);
        chai.assert.isTrue(result.success);
      }
    }
  );
});
