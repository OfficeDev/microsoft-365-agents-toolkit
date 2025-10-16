// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhaofeng Xu <zhaofengxu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import * as chai from "chai";
import { describe } from "mocha";
import * as path from "path";
import { CliHelper } from "../../commonlib/cliHelper";
import { TemplateProjectFolder } from "../../utils/constants";
import { Executor } from "../../utils/executor";
import { getTestFolder, getUniqueAppName } from "../commonUtils";
import fs from "fs-extra";

describe.skip("upgrade", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    await fs.remove(projectPath);
  });

  it(
    "sample incoming webhook notification",
    { testPlanCaseId: 19298763, author: "zhaofengxu@microsoft.com" },
    async function () {
      {
        await Executor.installCLI(testFolder, "1.2.5", false);
        const env = Object.assign({}, process.env);
        // new projiect
        await CliHelper.createTemplateProject(
          appName,
          testFolder,
          TemplateProjectFolder.IncomingWebhook,
          env,
          true,
          true,
          false
        );
      }

      {
        // provision
        const result = await Executor.provision(projectPath, "dev", true);
        chai.assert.include(
          result.stderr,
          "This command only works for project created by Teams Toolkit"
        );
      }
    }
  );
});
