// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Huihui Wu <huihuiwu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import {
  ProgrammingLanguage,
  environmentNameManager,
} from "@microsoft/teamsfx-core";
import { assert, expect } from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import { AppStudioValidator, SharepointValidator } from "../../commonlib";
import { Capability } from "../../utils/constants";
import { Executor } from "../../utils/executor";
import {
  cleanUpLocalProject,
  cleanupSharePointPackage,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
} from "../commonUtils";

describe("Start a new project", function () {
  let appId: string;
  let appName: string;
  let testFolder: string;
  let projectPath: string;
  let teamsAppId: string | undefined;
  beforeEach(async () => {
    testFolder = getTestFolder();
    appName = getUniqueAppName();
    projectPath = path.resolve(testFolder, appName);
  });

  it(
    "Create, provision and run SPFx project with React framework",
    { testPlanCaseId: 15687302, author: "huihuiwu@microsoft.com" },
    async function () {
      {
        const result = await Executor.createProject(
          testFolder,
          appName,
          Capability.Spfx,
          ProgrammingLanguage.TS
        );

        // check specified files
        const files: string[] = [
          "config/config.json",
          "config/deploy-azure-storage.json",
          "config/package-solution.json",
          "config/serve.json",
          "config/write-manifests.json",
          "src/webparts/helloworld/HelloworldWebPart.manifest.json",
          "src/webparts/helloworld/HelloworldWebPart.ts",
          "src/webparts/helloworld/loc/en-us.js",
          "src/webparts/helloworld/loc/mystrings.d.ts",
          "src/webparts/helloworld/assets/welcome-dark.png",
          "src/webparts/helloworld/assets/welcome-light.png",
          "src/webparts/helloworld/components/Helloworld.module.scss",
          "src/webparts/helloworld/components/Helloworld.tsx",
          "src/webparts/helloworld/components/IHelloworldProps.ts",
          "src/index.ts",
          ".gitignore",
          ".npmignore",
          ".yo-rc.json",
          "gulpfile.js",
          "package.json",
          "README.md",
          "tsconfig.json",
        ];
        for (const file of files) {
          const filePath = path.join(testFolder, appName, `src`, file);
          expect(fs.existsSync(filePath), `${filePath} must exist.`).to.eq(
            true
          );
        }
        expect(result.success).to.be.true;
      }

      {
        // validation succeed without provision
        const result = await Executor.validate(
          projectPath,
          environmentNameManager.getDefaultEnvName()
        );
        expect(result.success).to.be.true;
      }

      {
        // validation local env succeed without local debug
        const result = await Executor.validate(
          projectPath,
          environmentNameManager.getLocalEnvName()
        );
        expect(result.success).to.be.true;
      }

      {
        // provision
        const result = await Executor.provision(
          projectPath,
          environmentNameManager.getDefaultEnvName()
        );
        expect(result.success).to.be.true;
      }

      {
        // Get context
        const context = await readContextMultiEnvV3(
          projectPath,
          environmentNameManager.getDefaultEnvName()
        );

        assert.exists(context.TEAMS_APP_ID);
        teamsAppId = context.TEAMS_APP_ID;
        AppStudioValidator.setE2ETestProvider();
      }

      {
        // deploy
        const result = await Executor.deploy(
          projectPath,
          environmentNameManager.getDefaultEnvName()
        );
        expect(result.success).to.be.true;
      }

      {
        // Validate sharepoint package
        const solutionConfig = await fs.readJson(
          `${projectPath}/src/config/package-solution.json`
        );
        const sharepointPackage = `${projectPath}/src/sharepoint/${solutionConfig.paths.zippedPackage}`;
        appId = solutionConfig["solution"]["id"];
        expect(appId).to.not.be.empty;
        expect(await fs.pathExists(sharepointPackage)).to.be.true;

        // Check if package exsist in App Catalog
        SharepointValidator.init();
        SharepointValidator.validateDeploy(appId);
      }

      {
        // publish
        const result = await Executor.publish(
          projectPath,
          environmentNameManager.getDefaultEnvName()
        );
        expect(result.success).to.be.true;
      }

      {
        // Validate publish result
        await AppStudioValidator.validatePublish(teamsAppId!);
      }

      {
        // preview
        const result = await Executor.preview(
          projectPath,
          environmentNameManager.getDefaultEnvName()
        );
        expect(result.success).to.be.true;
      }
    }
  );

  afterEach(async () => {
    // clean up
    await cleanUpLocalProject(projectPath);
    await cleanupSharePointPackage(appId);
    await AppStudioValidator.cancelStagedAppInTeamsAppCatalog(teamsAppId);
  });
});
