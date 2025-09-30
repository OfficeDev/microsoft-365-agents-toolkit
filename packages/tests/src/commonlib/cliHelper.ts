// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  execAsync,
  execAsyncWithRetry,
  editDotEnvFile,
  getKeyVaultNameFromResourceId,
  getProvisionParameterValueByKey,
  getKeyVaultSecretReference,
} from "./utilities";
import {
  PluginId,
  provisionParametersKey,
  Resource,
  ResourceToDeploy,
  StateConfigKey,
} from "./constants";
import { TemplateProjectFolder } from "../utils/constants";
import { Capability } from "../utils/constants";
import * as path from "path";

export class CliHelper {
  static async addEnv(
    env: string,
    projectPath: string,
    processEnv?: NodeJS.ProcessEnv
  ) {
    const command = `atk env add ${env} --env dev`;
    const timeout = 100000;

    try {
      const result = await execAsync(command, {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });
      if (result.stderr) {
        console.error(
          `[Failed] add environment for ${projectPath}. Error message: ${result.stderr}`
        );
      } else {
        console.log(`[Successfully] add environment for ${projectPath}`);
      }
    } catch (e) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async provisionProject(
    projectPath: string,
    option = "",
    env: "dev" | "local" = "dev",
    processEnv?: NodeJS.ProcessEnv
  ) {
    const result = await execAsyncWithRetry(
      `atk provision --env ${env} --interactive false --verbose ${option}`,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: 0,
      }
    );

    if (result.stderr) {
      console.error(
        `[Failed] provision ${projectPath}. Error message: ${result.stderr}`
      );
    } else {
      console.log(`[Successfully] provision ${projectPath}`);
    }
  }

  static async updateAadManifest(
    projectPath: string,
    option = "",
    processEnv?: NodeJS.ProcessEnv,
    retries?: number,
    newCommand?: string
  ) {
    const result = await execAsyncWithRetry(
      `atk entra-app update ${option} --interactive false`,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: 0,
      },
      retries,
      newCommand
    );
    const message = `update aad-app manifest template for ${projectPath}`;
    if (result.stderr) {
      console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
    } else {
      console.log(`[Successfully] ${message}`);
    }
  }

  static async deployAll(
    projectPath: string,
    option = "",
    env: "dev" | "local" = "dev",
    processEnv?: NodeJS.ProcessEnv,
    retries?: number,
    newCommand?: string
  ) {
    const result = await execAsyncWithRetry(
      `atk deploy --env ${env} --interactive false --verbose ${option}`,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: 0,
      },
      retries,
      newCommand
    );
    const message = `deploy all resources for ${projectPath}`;
    if (result.stderr) {
      console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
    } else {
      console.log(`[Successfully] ${message}`);
    }
  }

  static async deployProject(
    resourceToDeploy: ResourceToDeploy,
    projectPath: string,
    option = "",
    processEnv?: NodeJS.ProcessEnv,
    retries?: number,
    newCommand?: string
  ) {
    const result = await execAsyncWithRetry(
      `atk deploy ${resourceToDeploy} ${option}`,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: 0,
      },
      retries,
      newCommand
    );
    const message = `deploy ${resourceToDeploy} for ${projectPath}`;
    if (result.stderr) {
      console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
    } else {
      console.log(`[Successfully] ${message}`);
    }
  }

  static async createDotNetProject(
    appName: string,
    testFolder: string,
    capability: Capability,
    processEnv?: NodeJS.ProcessEnv,
    options = ""
  ): Promise<void> {
    const command = `atk new --interactive false --runtime dotnet --app-name ${appName} --capability ${capability} ${options}`;
    const timeout = 100000;
    try {
      const result = await execAsync(command, {
        cwd: testFolder,
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });
      const message = `scaffold project to ${path.resolve(
        testFolder,
        appName
      )} with capability ${capability}`;
      if (result.stderr) {
        console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
      } else {
        console.log(`[Successfully] ${message}`);
      }
    } catch (e) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async createProjectWithCapability(
    appName: string,
    testFolder: string,
    capability: Capability,
    processEnv?: NodeJS.ProcessEnv,
    options = "",
    npx = false
  ) {
    const npxCommand = npx ? "npx" : "";
    const command = `${npxCommand} atk new --interactive false --app-name ${appName} --capability ${capability} ${options}`;
    const timeout = 100000;
    try {
      const result = await execAsync(command, {
        cwd: testFolder,
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });
      const message = `scaffold project to ${path.resolve(
        testFolder,
        appName
      )} with capability ${capability}`;
      if (result.stderr) {
        console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
      } else {
        console.log(`[Successfully] ${message}`);
      }
    } catch (e) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async openTemplateProject(
    appName: string,
    testFolder: string,
    template: TemplateProjectFolder,
    processEnv?: NodeJS.ProcessEnv
  ) {
    const timeout = 100000;
    const oldPath = path.resolve("./resource", template);
    const newPath = path.resolve(testFolder, appName);
    try {
      await execAsync(`mv ${oldPath} ${newPath}`, {
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });
    } catch (error) {
      console.log(error);
      throw new Error(`Failed to open project: ${newPath}`);
    }
    const localEnvPath = path.resolve(testFolder, appName, "env", ".env.local");
    const remoteEnvPath = path.resolve(testFolder, appName, "env", ".env.dev");
    editDotEnvFile(localEnvPath, "TEAMS_APP_NAME", appName);
    editDotEnvFile(remoteEnvPath, "TEAMS_APP_NAME", appName);
  }

  static async createTemplateProject(
    appName: string,
    testFolder: string,
    template: TemplateProjectFolder,
    processEnv?: NodeJS.ProcessEnv,
    npx = false,
    oldNewCommand = false,
    isV3 = true
  ) {
    const npxCommand = npx ? "npx" : "";
    const newCommand = oldNewCommand ? "template" : "sample";
    const cliPrefix = isV3 ? "teamsapp" : "teamsfx";
    const command = `${npxCommand} ${cliPrefix} new ${newCommand} ${template} --interactive false `;
    const timeout = 100000;
    try {
      const result = await execAsync(command, {
        cwd: testFolder,
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });

      //  change original template name to appName
      await execAsync(`mv ./${template} ./${appName}`, {
        cwd: testFolder,
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });

      const localEnvPath = path.resolve(
        testFolder,
        appName,
        "env",
        ".env.local"
      );
      const remoteEnvPath = path.resolve(
        testFolder,
        appName,
        "env",
        ".env.dev"
      );
      editDotEnvFile(localEnvPath, "TEAMS_APP_NAME", appName);
      editDotEnvFile(remoteEnvPath, "TEAMS_APP_NAME", appName);

      const message = `scaffold project to ${path.resolve(
        testFolder,
        appName
      )} with template ${template}`;
      if (result.stderr) {
        console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
      } else {
        console.log(`[Successfully] ${message}`);
      }
    } catch (e) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async addCapabilityToProject(
    projectPath: string,
    capabilityToAdd: Capability
  ) {
    const command = `atk add ${capabilityToAdd}`;
    const timeout = 100000;
    try {
      const result = await execAsync(command, {
        cwd: projectPath,
        env: process.env,
        timeout: timeout,
      });
      const message = `add capability ${capabilityToAdd} to ${projectPath}`;
      if (result.stderr) {
        console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
      } else {
        console.log(`[Successfully] ${message}`);
      }
    } catch (e) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async getUserSettings(
    key: string,
    projectPath: string,
    env: string
  ): Promise<string> {
    const value = "";
    // config not supported
    // const command = `atk config get ${key} --env ${env}`;
    // const timeout = 100000;
    // try {
    //   const result = await execAsync(command, {
    //     cwd: projectPath,
    //     env: process.env,
    //     timeout: timeout,
    //   });

    //   const message = `get user settings in ${projectPath}. Key: ${key}`;
    //   if (result.stderr) {
    //     console.error(`[Failed] ${message}. Error message: ${result.stderr}`);
    //   } else {
    //     const arr = (result.stdout as string).split(":");
    //     if (!arr || arr.length <= 1) {
    //       console.error(
    //         `[Failed] ${message}. Failed to get value from cli result. result: ${result.stdout}`
    //       );
    //     } else {
    //       value = arr[1].trim() as string;
    //       console.log(`[Successfully] ${message}.`);
    //     }
    //   }
    // } catch (e) {
    //   console.log(
    //     `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
    //   );
    //   if (e.killed && e.signal == "SIGTERM") {
    //     console.log(`Command ${command} killed due to timeout ${timeout}`);
    //   }
    // }
    return value;
  }
}

export async function getExpectedM365ClientSecret(
  ctx: any,
  projectPath: string,
  env: string,
  activeResourcePlugins: string[]
): Promise<string> {
  let m365ClientSecret: string;
  if (activeResourcePlugins.includes(PluginId.KeyVault)) {
    const vaultName = getKeyVaultNameFromResourceId(
      ctx[PluginId.KeyVault][StateConfigKey.keyVaultResourceId]
    );
    const secretName =
      (await getProvisionParameterValueByKey(
        projectPath,
        env,
        provisionParametersKey.m365ClientSecretName
      )) ?? "m365ClientSecret";
    m365ClientSecret = getKeyVaultSecretReference(vaultName, secretName);
  } else {
    m365ClientSecret = await CliHelper.getUserSettings(
      `${PluginId.Aad}.${StateConfigKey.clientSecret}`,
      projectPath,
      env
    );
  }
  return m365ClientSecret;
}

export async function getExpectedBotClientSecret(
  ctx: any,
  projectPath: string,
  env: string,
  activeResourcePlugins: string[]
): Promise<string> {
  let botClientSecret: string;
  if (activeResourcePlugins.includes(PluginId.KeyVault)) {
    const vaultName = getKeyVaultNameFromResourceId(
      ctx[PluginId.KeyVault][StateConfigKey.keyVaultResourceId]
    );
    const secretName =
      (await getProvisionParameterValueByKey(
        projectPath,
        env,
        provisionParametersKey.botClientSecretName
      )) ?? "botClientSecret";
    botClientSecret = getKeyVaultSecretReference(vaultName, secretName);
  } else {
    botClientSecret = await CliHelper.getUserSettings(
      `${PluginId.Bot}.${StateConfigKey.botPassword}`,
      projectPath,
      env
    );
  }
  return botClientSecret;
}
