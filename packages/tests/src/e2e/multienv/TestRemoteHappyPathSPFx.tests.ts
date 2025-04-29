// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yuan Tian <tianyuan@microsoft.com>
 */

import * as fs from "fs-extra";
import * as path from "path";
import { expect, assert } from "chai";

import {
  cleanUpLocalProject,
  cleanupSharePointPackage,
  execAsync,
  execAsyncWithRetry,
  getTestFolder,
  getUniqueAppName,
  mockTeamsfxMultiEnvFeatureFlag,
  readContextMultiEnvV3,
  removeTeamsAppExtendToM365,
} from "../commonUtils";
import { AppPackageFolderName, BuildFolderName } from "@microsoft/teamsfx-api";
import { AppStudioValidator, SharepointValidator } from "../../commonlib";
import { it } from "@microsoft/extra-shot-mocha";

describe("Multi Env Happy Path for SPFx", function () {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const type = "none";
  const processEnv = mockTeamsfxMultiEnvFeatureFlag();
  const env = "e2e";
  let appId: string;
  let teamsAppId: string | undefined;

  it(
    "Can create/provision/deploy/validate/package/publish an SPFx project",
    { testPlanCaseId: 24137702, author: "tianyuan@microsoft.com" },
    async function () {
      const command = `atk new --interactive false --app-name ${appName} --capability tab-spfx --spfx-framework-type ${type} --spfx-webpart-name helloworld --programming-language typescript`;
      let result = await execAsync(command, {
        cwd: testFolder,
        env: processEnv,
        timeout: 0,
      });

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
        expect(fs.existsSync(filePath), `${filePath} must exist.`).to.eq(true);
      }

      expect(result.stderr).to.eq("");

      // add env
      result = await execAsync(`atk env add ${env} --env dev`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });
      console.log(
        `[Successfully] env add, stdout: '${result.stdout}', stderr: '${result.stderr}'`
      );

      // list env
      result = await execAsync(`atk env list`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });
      const envs = result.stdout.trim().split(/\r?\n/).sort();
      expect(envs).to.include("dev");
      expect(envs).to.include("e2e");
      expect(result.stderr).to.be.empty;
      console.log(
        `[Successfully] env list, stdout: '${result.stdout}', stderr: '${result.stderr}'`
      );

      // remove teamsApp/extendToM365 in case it fails
      removeTeamsAppExtendToM365(path.join(projectPath, "teamsapp.yml"));

      // provision
      result = await execAsyncWithRetry(`atk provision --env ${env}`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });
      console.log(
        `[Successfully] provision, stdout: '${result.stdout}', stderr: '${result.stderr}'`
      );

      {
        // Get context
        const context = await readContextMultiEnvV3(projectPath, env);

        teamsAppId = context.TEAMS_APP_ID;
        AppStudioValidator.setE2ETestProvider();
      }

      // deploy
      result = await execAsyncWithRetry(`atk deploy --env ${env}`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });
      console.log(
        `[Successfully] deploy, stdout: '${result.stdout}', stderr: '${result.stderr}'`
      );

      {
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

      // validate manifest
      result = await execAsyncWithRetry(`atk validate --env ${env}`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });
      console.log(
        `[Successfully] validation, stdout: '${result.stdout}', stderr: '${result.stderr}'`
      );

      {
        // Validate validate manifest
        expect(result.stderr).to.be.empty;
      }

      // package
      await execAsyncWithRetry(`atk package --env ${env}`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });

      {
        // Validate package
        const file = `${projectPath}/${AppPackageFolderName}/${BuildFolderName}/appPackage.${env}.zip`;
        expect(await fs.pathExists(file)).to.be.true;
      }

      // publish
      result = await execAsyncWithRetry(`atk publish --env ${env}`, {
        cwd: projectPath,
        env: processEnv,
        timeout: 0,
      });

      {
        // Validate publish result
        await AppStudioValidator.validatePublish(teamsAppId!);
      }
    }
  );

  after(async () => {
    // clean up
    await cleanUpLocalProject(projectPath);
    await cleanupSharePointPackage(appId);
    await AppStudioValidator.cancelStagedAppInTeamsAppCatalog(teamsAppId);
  });
});
