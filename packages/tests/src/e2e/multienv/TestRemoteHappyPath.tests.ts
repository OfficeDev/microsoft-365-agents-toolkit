// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yuan Tian <tianyuan@microsoft.com>
 */

import { AppPackageFolderName, BuildFolderName } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import fs from "fs-extra";
import { describe } from "mocha";
import path from "path";
import M365Login from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365Login";
import { AppStudioValidator, BotValidator } from "../../commonlib";
import { CliHelper } from "../../commonlib/cliHelper";
import {
  cleanUp,
  execAsync,
  execAsyncWithRetry,
  getTestFolder,
  getUniqueAppName,
  mockTeamsfxMultiEnvFeatureFlag,
  readContextMultiEnvV3,
  createResourceGroup,
} from "../commonUtils";
import { expect } from "chai";
import { Executor } from "../../utils/executor";
import { it } from "@microsoft/extra-shot-mocha";
import { EnvConstants } from "../../commonlib/constants";

describe("Multi Env Happy Path for Azure", function () {
  const env = "e2e";
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const processEnv = mockTeamsfxMultiEnvFeatureFlag();
  let teamsAppId: string | undefined;

  it(
    `Can create/provision/deploy/build/validate/launch remote a azure bot project`,
    { testPlanCaseId: 24137694, author: "tianyuan@microsoft.com" },
    async function () {
      try {
        let result;
        result = await execAsync(
          `teamsapp new --interactive false --app-name ${appName} --capability notification --bot-host-type-trigger http-functions --programming-language javascript`,
          {
            cwd: testFolder,
            env: processEnv,
            timeout: 0,
          }
        );
        console.log(
          `[Successfully] scaffold to ${projectPath}, stdout: '${result.stdout}', stderr: '${result.stderr}''`
        );
        // add env
        await CliHelper.addEnv(env, projectPath, processEnv);

        // list env
        result = await execAsync(`teamsapp env list`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });
        const envs = result.stdout.trim().split(/\r?\n/).sort();
        expect(envs).to.include("dev");
        expect(envs).to.include("e2e");
        chai.expect(result.stderr).to.be.empty;
        console.log(
          `[Successfully] env list, stdout: '${result.stdout}', stderr: '${result.stderr}'`
        );

        {
          // provision
          const result = await createResourceGroup(appName + "-rg", "westus");
          expect(result).to.be.true;
          process.env["AZURE_RESOURCE_GROUP_NAME"] = appName + "-rg";
          const { success } = await Executor.provision(projectPath, env);
          expect(success).to.be.true;
          console.log(`[Successfully] provision for ${projectPath}`);
        }

        {
          // Validate provision
          // Get context
          const context = await readContextMultiEnvV3(projectPath, env);

          // Validate Bot Provision
          const bot = new BotValidator(context, projectPath, env);
          await bot.validateProvisionV3(false);
        }

        {
          // deploy
          const { success } = await Executor.deploy(projectPath, env);
          expect(success).to.be.true;
        }

        {
          // Validate deployment
          // Get context
          const context = await readContextMultiEnvV3(projectPath, env);
          // Validate Bot Deploy
          const bot = new BotValidator(context, projectPath, env);
          await bot.validateDeploy();
        }

        // validate manifest
        result = await execAsyncWithRetry(`teamsapp validate --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate validate manifest
          chai.expect(result.stderr).to.be.empty;
        }

        // update manifest
        const updateManifestCmd = `teamsapp update --env ${env}`;
        result = await execAsyncWithRetry(updateManifestCmd, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate update manifest
          chai.expect(result.stderr).to.be.empty;
        }

        // package
        await execAsyncWithRetry(`teamsapp package --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate package
          const file = `${projectPath}/${AppPackageFolderName}/${BuildFolderName}/appPackage.${env}.zip`;
          chai.expect(await fs.pathExists(file)).to.be.true;
        }

        // publish
        await execAsyncWithRetry(`teamsapp publish --env ${env}`, {
          cwd: projectPath,
          env: processEnv,
          timeout: 0,
        });

        {
          // Validate publish result
          const context = await readContextMultiEnvV3(projectPath, env);
          teamsAppId = context[EnvConstants.TEAMS_APP_ID];
          chai.assert.isNotNull(teamsAppId);
          AppStudioValidator.provider = M365Login;
          await AppStudioValidator.validatePublish(teamsAppId!);
        }
      } catch (e: any) {
        console.log("Unexpected exception is thrown when running test: " + e);
        console.log(e.stack);
        throw e;
      }
    }
  );

  after(async () => {
    // clean up
    await cleanUp(appName, projectPath, true, true, false, env, teamsAppId);
  });
});
