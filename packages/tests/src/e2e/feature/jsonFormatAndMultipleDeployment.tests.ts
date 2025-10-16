// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhaofeng Xu <zhaofengxu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import { describe } from "mocha";
import * as path from "path";
import { Cleaner } from "../../commonlib/cleaner";
import { Capability } from "../../utils/constants";
import {
  addJsonFileAndParamtersFile,
  updateYml,
} from "../../utils/armDriverUtils";
import { Executor } from "../../utils/executor";
import {
  createResourceGroup,
  getTestFolder,
  getUniqueAppName,
} from "../commonUtils";

describe("version check", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    await Cleaner.clean(projectPath);
  });

  it(
    "arm json format and multiple deployment",
    { testPlanCaseId: 16835373, author: "zhaofengxu@microsoft.com" },
    async function () {
      {
        const result = await Executor.createProject(
          testFolder,
          appName,
          Capability.TabNonSso,
          ProgrammingLanguage.TS
        );
        chai.assert.isTrue(result.success);
      }
      await addJsonFileAndParamtersFile(projectPath);
      await updateYml(projectPath);
      {
        // provision
        const rgResult = await createResourceGroup(appName + "-rg", "westus");
        chai.assert.isTrue(rgResult);
        process.env["AZURE_RESOURCE_GROUP_NAME"] = appName + "-rg";
        const result = await Executor.provision(projectPath);
        chai.assert.isTrue(result.success);
      }
    }
  );
});
