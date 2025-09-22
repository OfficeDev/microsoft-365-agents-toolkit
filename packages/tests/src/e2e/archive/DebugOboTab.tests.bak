// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Kuojian Lu <kuojianlu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import {
  environmentNameManager,
  ProgrammingLanguage,
} from "@microsoft/teamsfx-core";
import * as chai from "chai";
import { expect } from "chai";
import * as fs from "fs-extra";
import { describe } from "mocha";
import * as path from "path";
import { AadValidator } from "../../commonlib/aadValidate";
import { Cleaner } from "../../commonlib/cleaner";
import { Capability } from "../../utils/constants";
import { Executor } from "../../utils/executor";
import { ProjectEnvReader } from "../../commonlib/projectEnvReader";
import { getTestFolder, getUniqueAppName } from "../commonUtils";
import { getTeamsApp } from "../debug/utility";

describe("Debug V3 m365-tab template", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    await Cleaner.clean(projectPath);
  });

  it(
    "happy path: provision and deploy",
    { testPlanCaseId: 17449535, author: "kuojianlu@microsoft.com" },
    async function () {
      {
        const result = await Executor.createProject(
          testFolder,
          appName,
          Capability.M365SsoLaunchPage,
          ProgrammingLanguage.TS
        );
        expect(result.success).to.be.true;
      }

      {
        // provision
        const result = await Executor.provision(
          projectPath,
          environmentNameManager.getLocalEnvName()
        );
        expect(result.success).to.be.true;

        const context = await ProjectEnvReader.readEnvFile(
          projectPath,
          environmentNameManager.getLocalEnvName()
        );
        chai.assert.isDefined(context);

        // validate aad
        chai.assert.isDefined(context!.AAD_APP_OBJECT_ID);
        chai.assert.isNotEmpty(context!.AAD_APP_OBJECT_ID);
        const aad = AadValidator.init(context, false);
        await AadValidator.validate(aad);

        // validate teams app
        chai.assert.isDefined(context!.TEAMS_APP_ID);
        const teamsApp = await getTeamsApp(context!.TEAMS_APP_ID);
        chai.assert.equal(teamsApp?.teamsAppId, context!.TEAMS_APP_ID);

        // validate m365
        chai.assert.isDefined(context!.M365_TITLE_ID);
        chai.assert.isNotEmpty(context!.M365_TITLE_ID);
        chai.assert.isDefined(context!.M365_APP_ID);
        chai.assert.isNotEmpty(context!.M365_APP_ID);
      }

      {
        // deploy
        const result = await Executor.deploy(
          projectPath,
          environmentNameManager.getLocalEnvName()
        );
        expect(result.success).to.be.true;

        const context = await ProjectEnvReader.readEnvFile(
          projectPath,
          environmentNameManager.getLocalEnvName()
        );
        chai.assert.isDefined(context);

        // validate ssl cert
        chai.assert.isDefined(context!.SSL_CRT_FILE);
        chai.assert.isNotEmpty(context!.SSL_CRT_FILE);
        chai.assert.isDefined(context!.SSL_KEY_FILE);
        chai.assert.isNotEmpty(context!.SSL_KEY_FILE);

        // validate .localConfigs
        chai.assert.isTrue(
          await fs.pathExists(path.join(projectPath, ".localConfigs"))
        );
      }
    }
  );
});
