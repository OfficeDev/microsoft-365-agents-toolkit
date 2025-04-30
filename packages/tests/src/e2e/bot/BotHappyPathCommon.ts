// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Siglud <fanhu@microsoft.com>
 **/

import * as path from "path";
import { BotValidator } from "../../commonlib";
import {
  cleanUp,
  execAsync,
  execAsyncWithRetry,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
  createResourceGroup,
} from "../commonUtils";
import { environmentNameManager } from "@microsoft/teamsfx-core";
import {
  Runtime,
  CliCapabilities,
  CliTriggerType,
} from "../../commonlib/constants";
import { Executor } from "../../utils/executor";
import { expect } from "chai";

export async function happyPathTest(
  runtime: Runtime,
  capabilities: CliCapabilities,
  trigger?: CliTriggerType[],
  isTdpIntegrationTemplatesOnly?: boolean
): Promise<void> {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const envName = environmentNameManager.getDefaultEnvName();

  const env = Object.assign({}, process.env);
  if (runtime === Runtime.Dotnet) {
    env["TEAMSFX_CLI_DOTNET"] = "true";
  }

  if (isTdpIntegrationTemplatesOnly) {
    env["TEAMSFX_TDP_TEMPLATE_CLI_TEST"] = "true";
  }
  const triggerStr =
    trigger === undefined
      ? ""
      : `--bot-host-type-trigger ${trigger.join(" ")} `;
  const cmdBase = `atk new --interactive false --app-name ${appName} --capability ${capabilities} ${triggerStr}`;
  const cmd =
    runtime === Runtime.Dotnet
      ? `${cmdBase} --runtime dotnet`
      : `${cmdBase} --programming-language typescript`;
  console.log(`ready to run CMD: ${cmd}`);
  await execAsync(cmd, {
    cwd: testFolder,
    env: env,
    timeout: 0,
  });
  console.log(`[Successfully] scaffold to ${projectPath}`);

  // set subscription
  // await CliHelper.setSubscription(subscription, projectPath, env);

  console.log(`[Successfully] set subscription for ${projectPath}`);

  {
    // provision
    const result = await createResourceGroup(appName + "-rg", "westus");
    expect(result).to.be.true;
    process.env["AZURE_RESOURCE_GROUP_NAME"] = appName + "-rg";
    const { success } = await Executor.provision(projectPath, envName);
    expect(success).to.be.true;
    console.log(`[Successfully] provision for ${projectPath}`);
  }

  {
    // Validate provision
    // Get context
    const context = await readContextMultiEnvV3(projectPath, envName);

    // Validate Bot Provision
    const bot = new BotValidator(context, projectPath, envName);
    await bot.validateProvisionV3(false);
  }

  // deploy
  const cmdStr = "atk deploy";
  await execAsyncWithRetry(cmdStr, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });
  console.log(`[Successfully] deploy for ${projectPath}`);

  {
    // Validate deployment

    // Get context
    const context = await readContextMultiEnvV3(projectPath, envName);

    // Validate Bot Deploy
    const bot = new BotValidator(context, projectPath, envName);
    await bot.validateDeploy();
  }

  // test (validate)
  await execAsyncWithRetry(`atk validate --env ${envName}`, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });

  // package
  await execAsyncWithRetry(`atk package --env ${envName}`, {
    cwd: projectPath,
    env: env,
    timeout: 0,
  });

  console.log(`[Successfully] start to clean up for ${projectPath}`);
  await cleanUp(appName, projectPath, false, true, false);
}
