// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhiyu You <zhiyou@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { expect } from "chai";
import fs from "fs-extra";
import path from "path";
import { cleanUpLocalProject, execAsync, getTestFolder } from "../commonUtils";

describe("atk new sample", function () {
  const testFolder = getTestFolder();
  const sampleName = "todo-list-with-Azure-backend";
  const projectPath = path.resolve(testFolder, sampleName);

  it(
    `${sampleName}`,
    { testPlanCaseId: 24137474, author: "zhiyou@microsoft.com" },
    async function () {
      const env = Object.assign({}, process.env);
      env.TEAMSFX_SAMPLE_CONFIG_BRANCH = "dev";
      await execAsync(`atk new sample ${sampleName} -i false`, {
        cwd: testFolder,
        env: env,
        timeout: 0,
      });

      expect(fs.pathExistsSync(projectPath)).to.be.true;
      expect(fs.pathExistsSync(path.resolve(projectPath, "teamsapp.yml"))).to.be
        .true;
    }
  );

  after(async () => {
    // clean up
    await cleanUpLocalProject(projectPath);
  });
});
