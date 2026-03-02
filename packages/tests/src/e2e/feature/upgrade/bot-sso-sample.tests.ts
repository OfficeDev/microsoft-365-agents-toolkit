// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhaofeng Xu <zhaofengxu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { MetadataV3 } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import fs from "fs-extra";
import { describe } from "mocha";
import * as path from "path";
import { Cleaner } from "../../../commonlib/cleaner";
import { CliHelper } from "../../../commonlib/cliHelper";
import { TemplateProjectFolder } from "../../../utils/constants";
import { Executor } from "../../../utils/executor";
import { getTestFolder, getUniqueAppName } from "../../commonUtils";

describe("upgrade", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    await Cleaner.clean(projectPath);
  });

  it(
    "sample bot sso",
    { testPlanCaseId: 19314244, author: "zhaofengxu@microsoft.com" },
    async function () {
      {
        await Executor.installCLI(testFolder, "1.2.5", false);
        const env = Object.assign({}, process.env);
        // new projiect
        await CliHelper.createTemplateProject(
          appName,
          testFolder,
          TemplateProjectFolder.HelloWorldBotSSO,
          env,
          true,
          true,
          false
        );
      }

      {
        // upgrade
        const result = await Executor.upgrade(projectPath, true);
        chai.assert.isTrue(result.success);
        const manifestPath = path.join(
          projectPath,
          MetadataV3.teamsManifestFolder,
          MetadataV3.teamsManifestFileName
        );
        const content = await fs.readFile(manifestPath, { encoding: "utf-8" });
        const res = JSON.parse(content);
        chai.assert.isTrue(
          res.validDomains.includes(
            "${{PROVISIONOUTPUT__BOTOUTPUT__VALIDDOMAIN}}"
          )
        );
      }
    }
  );
});
