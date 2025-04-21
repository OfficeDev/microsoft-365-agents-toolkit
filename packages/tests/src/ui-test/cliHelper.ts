// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  execAsync,
  execAsyncWithRetry,
  timeoutPromise,
  killPort,
  spawnCommand,
  killNgrok,
} from "../utils/commonUtils";
import {
  TemplateProjectFolder,
  ResourceToDeploy,
  Capability,
} from "../utils/constants";
import path from "path";
import * as chai from "chai";
import { Executor } from "../utils/executor";
import * as os from "os";
import { ChildProcess, ChildProcessWithoutNullStreams } from "child_process";

export class CliHelper {
  static async addEnv(
    env: string,
    projectPath: string,
    processEnv?: NodeJS.ProcessEnv
  ) {
    const command = `atk env add ${env} --env dev --telemetry false`;
    const timeout = 100000;

    try {
      const result = await execAsync(command, {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: timeout,
      });
      if (result.stderr) {
        console.log(
          `[Failed] add environment for ${projectPath}. Error message: ${result.stderr}`
        );
      } else {
        console.log(`[Successfully] add environment for ${projectPath}`);
      }
    } catch (e: any) {
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
    env: "local" | "dev" = "local",
    v3 = true,
    processEnv: NodeJS.ProcessEnv = process.env,
    delay: number = 10 * 60 * 1000
  ) {
    console.log(`[Provision] ${projectPath}`);
    const timeout = timeoutPromise(delay);
    let command = "";
    if (v3) {
      command = `npx atk -v`;
    } else {
      command = `npx teamsfx -v`;
    }
    const version = await execAsyncWithRetry(
      command,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
      },
      1
    );
    console.log(`[Provision] cli version: ${version.stdout}`);

    if (v3) {
      const childProcess = spawnCommand(
        os.type() === "Windows_NT" ? "npx.cmd" : "npx",
        [
          "teamsapp",
          "provision",
          "--env",
          env,
          "--verbose",
          "--telemetry",
          "false",
        ],
        {
          cwd: projectPath,
          env: processEnv ? processEnv : process.env,
        },
        (data) => {
          console.log(data);
        },
        (error) => {
          console.log(error);
        }
      );
      await Promise.all([timeout, childProcess]);
      // close process
      childProcess.kill("SIGKILL");
    } else {
      const childProcess = spawnCommand(
        os.type() === "Windows_NT" ? "npx.cmd" : "npx",
        [
          "teamsfx",
          "provision",
          "--env",
          env,
          "--resource-group",
          processEnv?.AZURE_RESOURCE_GROUP_NAME
            ? processEnv.AZURE_RESOURCE_GROUP_NAME
            : "",
          "--verbose",
          "--interactive",
          "false",
        ],
        {
          cwd: projectPath,
          env: processEnv ? processEnv : process.env,
        },
        (data) => {
          console.log(data);
        },
        (error) => {
          console.log(error);
        }
      );
      await Promise.all([timeout, childProcess]);
      // close process
      childProcess.kill("SIGKILL");
    }
  }

  static async provisionProject2(
    projectPath: string,
    option = "",
    env: "dev" | "local" = "dev",
    processEnv?: NodeJS.ProcessEnv
  ) {
    const result = await execAsyncWithRetry(
      `atk provision --env ${env} --interactive false --verbose ${option} --telemetry false`,
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

  static async showVersion(
    projectPath: string,
    processEnv?: NodeJS.ProcessEnv
  ) {
    const result = await execAsyncWithRetry(`atk --version`, {
      cwd: projectPath,
      env: processEnv ? processEnv : process.env,
      timeout: 0,
    });

    console.log(`Cli Version: ${result.stdout}`);
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
      `atk deploy --env ${env} --interactive false --verbose ${option} --telemetry false`,
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

  static async publishProject(
    projectPath: string,
    env: "local" | "dev" = "local",
    option = "",
    processEnv?: NodeJS.ProcessEnv
  ) {
    console.log(`[publish] ${projectPath}`);
    const result = await execAsyncWithRetry(
      `atk publish --env ${env} --verbose  ${option} --telemetry false`,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
        timeout: 0,
      }
    );

    if (result.stderr) {
      console.log(
        `[Failed] publish ${projectPath}. Error message: ${result.stderr}`
      );
    } else {
      console.log(`[Successfully] publish ${projectPath}`);
    }
  }

  static async addFeature(feature: string, cwd: string) {
    console.log(`[start] add feature ${feature}... `);
    const { success } = await Executor.execute(
      `teamsfx add ${feature} --verbose --interactive false`,
      cwd
    );
    chai.expect(success).to.be.true;
    const message = `[success] add ${feature} successfully !!!`;
    console.log(message);
  }

  static async updateAadManifest(
    projectPath: string,
    option = "",
    processEnv?: NodeJS.ProcessEnv,
    retries?: number,
    newCommand?: string
  ) {
    const result = await execAsyncWithRetry(
      `tamsapp entra-app update ${option} --interactive false --telemetry false`,
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
      console.log(`[Failed] ${message}. Error message: ${result.stderr}`);
    } else {
      console.log(`[Successfully] ${message}`);
    }
  }

  static async deploy(
    projectPath: string,
    env: "local" | "dev" = "local",
    v3 = true,
    processEnv: NodeJS.ProcessEnv = process.env,
    delay: number = 10 * 60 * 1000
  ) {
    console.log(`[Deploy] ${projectPath}`);
    const timeout = timeoutPromise(delay);

    let command = "";
    if (v3) {
      command = `npx atk -v`;
    } else {
      command = `npx teamsfx -v`;
    }
    const version = await execAsyncWithRetry(
      command,
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
      },
      1
    );
    console.log(`[Deploy] cli version: ${version.stdout}`);

    if (v3) {
      const childProcess = spawnCommand(
        os.type() === "Windows_NT" ? "npx.cmd" : "npx",
        [
          "teamsapp",
          "deploy",
          "--env",
          env,
          "--verbose",
          "--telemetry",
          "false",
        ],
        {
          cwd: projectPath,
          env: processEnv ? processEnv : process.env,
        },
        (data) => {
          console.log(data);
        },
        (error) => {
          console.log(error);
        }
      );
      await Promise.all([timeout, childProcess]);
      // close process
      childProcess.kill("SIGKILL");
    } else {
      const childProcess = spawnCommand(
        os.type() === "Windows_NT" ? "npx.cmd" : "npx",
        [
          "teamsfx",
          "deploy",
          "--env",
          env,
          "--verbose",
          "--interactive",
          "false",
        ],
        {
          cwd: projectPath,
          env: processEnv ? processEnv : process.env,
        },
        (data) => {
          console.log(data);
        },
        (error) => {
          console.log(error);
        }
      );
      await Promise.all([timeout, childProcess]);
      // close process
      childProcess.kill("SIGKILL");
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
    console.log("add command is not supported in v3");
  }

  static async createDotNetProject(
    appName: string,
    testFolder: string,
    capability: "tab" | "bot",
    processEnv?: NodeJS.ProcessEnv,
    options = ""
  ): Promise<void> {
    const command = `atk new --interactive false --runtime dotnet --app-name ${appName} --capability ${capability} ${options} --telemetry false`;
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
        console.log(`[Failed] ${message}. Error message: ${result.stderr}`);
      } else {
        console.log(`[Successfully] ${message}`);
      }
    } catch (e: any) {
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
    lang: "javascript" | "typescript" = "javascript",
    options = "",
    processEnv?: NodeJS.ProcessEnv
  ) {
    const command = `atk new --interactive false --app-name ${appName} --capability ${capability} --programming-language ${lang} ${options} --telemetry false`;
    const timeout = 100000;
    try {
      await Executor.execute("atk -v", testFolder);
      await Executor.execute(command, testFolder);
      const message = `scaffold project to ${path.resolve(
        testFolder,
        appName
      )} with capability ${capability}`;
      console.log(`[Successfully] ${message}`);
    } catch (e: any) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async createProjectWithCapabilityMigration(
    appName: string,
    testFolder: string,
    capability: Capability,
    lang: "javascript" | "typescript" = "javascript",
    options = "",
    processEnv?: NodeJS.ProcessEnv
  ) {
    let command;
    console.log("cli version is V3 or not: " + CliHelper.getVersionFlag());
    const versionFlag = JSON.parse(CliHelper.getVersionFlag() as string);
    if (versionFlag) {
      command = `atk new --interactive false --app-name ${appName} --capability ${capability} --programming-language ${lang} ${options} --telemetry false`;
    } else {
      command = `teamsfx new --interactive false --app-name ${appName} --capabilities ${capability} --programming-language ${lang} ${options}`;
    }
    const timeout = 100000;
    try {
      if (versionFlag) {
        const { stdout } = await Executor.execute("atk -v", testFolder);
        console.log(stdout);
      } else {
        const { stdout } = await Executor.execute("teamsfx -v", testFolder);
        console.log(stdout);
      }
      await Executor.execute(command, testFolder);
      const message = `scaffold project to ${path.resolve(
        testFolder,
        appName
      )} with capability ${capability}`;
      console.log(`[Successfully] ${message}`);
    } catch (e: any) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout ${timeout}`);
      }
    }
  }

  static async createTemplateProject(
    testFolder: string,
    template: TemplateProjectFolder,
    V3: boolean,
    processEnv?: NodeJS.ProcessEnv
  ) {
    process.env["TEAMSFX_V3"] = V3 ? "true" : "false";
    process.env["TEAMSFX_V3_MIGRATION"] = V3 ? "true" : "false";

    console.log("TEAMSFX_V3: " + process.env["TEAMSFX_V3"]);
    console.log(await Executor.execute("atk -v", testFolder));

    const command = `atk new sample ${template} --interactive false --telemetry false`;
    const timeout = 100000;
    try {
      const result = await Executor.execute(command, testFolder);

      const message = `scaffold project to ${path.resolve(
        template
      )} with template ${template}`;
      if (result.stderr) {
        console.log(`[Failed] ${message}. Error message: ${result.stderr}`);
      } else {
        console.log(`[Successfully] ${message}`);
      }
    } catch (e: any) {
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
    // const command = `teamsfx config get ${key} --env ${env}`;
    // const timeout = 100000;
    // try {
    //   const result = await execAsync(command, {
    //     cwd: projectPath,
    //     env: process.env,
    //     timeout: timeout,
    //   });

    //   const message = `get user settings in ${projectPath}. Key: ${key}`;
    //   if (result.stderr) {
    //     console.log(`[Failed] ${message}. Error message: ${result.stderr}`);
    //   } else {
    //     const arr = (result.stdout as string).split(":");
    //     if (!arr || arr.length <= 1) {
    //       console.log(
    //         `[Failed] ${message}. Failed to get value from cli result. result: ${result.stdout}`
    //       );
    //     } else {
    //       value = arr[1].trim() as string;
    //       console.log(`[Successfully] ${message}.`);
    //     }
    //   }
    // } catch (e: any) {
    //   console.log(
    //     `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
    //   );
    //   if (e.killed && e.signal == "SIGTERM") {
    //     console.log(`Command ${command} killed due to timeout ${timeout}`);
    //   }
    // }
    return value;
  }

  static async installCLI(version: string, global: boolean, cwd = "./") {
    console.log(`install CLI with version ${version}`);
    if (global) {
      const { success } = await Executor.execute(
        `npm install -g @microsoft/m365agentstoolkit-cli@${version}`,
        cwd
      );
      //chai.expect(success).to.be.true;
    } else {
      const { success } = await Executor.execute(
        `npm install @microsoft/m365agentstoolkit-cli@${version}`,
        cwd
      );
      //chai.expect(success).to.be.true;
    }
    console.log("install CLI successfully");
  }

  static setV3Enable() {
    process.env["TEAMSFX_V3"] = "true";
  }

  static setV2Enable() {
    process.env["TEAMSFX_V3"] = "false";
  }

  static getVersionFlag() {
    return process.env["TEAMSFX_V3"];
  }

  static async debugProject(
    projectPath: string,
    env: "local" | "dev" = "local",
    v3 = true,
    processEnv: NodeJS.ProcessEnv = process.env,
    delay: number = 8 * 60 * 1000
  ) {
    console.log(`[start] ${env} debug ... `);
    const timeout = timeoutPromise(delay);
    const childProcess = spawnCommand(
      os.type() === "Windows_NT"
        ? v3
          ? "teamsapp.cmd"
          : "teamsfx.cmd"
        : v3
        ? "teamsapp"
        : "teamsfx",
      v3
        ? ["preview", "--env", env, "--telemetry", "false"]
        : ["preview", `--${env}`],
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
      },
      (data) => {
        console.log(data);
      },
      (error) => {
        console.log(error);
        if (error.includes("Error:")) {
          chai.assert.fail(error);
        }
      }
    );
    await Promise.all([timeout, childProcess]);
    try {
      // close process & port
      childProcess.kill("SIGKILL");
    } catch (error) {
      console.log(`kill process failed, cause by: `, error);
    }
    try {
      const result = await killPort(53000);
      console.log(`close port 53000 successfully, `, result.stdout);
    } catch (error) {
      console.log(`close port 53000 failed, cause by: `, error);
    }
    try {
      const result = await killPort(7071);
      console.log(`close port 7071 successfully, `, result.stdout);
    } catch (error) {
      console.log(`close port 7071 failed, cause by: `, error);
    }
    try {
      const result = await killPort(9229);
      console.log(`close port 9229 successfully, `, result.stdout);
    } catch (error) {
      console.log(`close port 9229 failed, cause by: `, error);
    }
    try {
      const result = await killPort(3978);
      console.log(`close port 3978 successfully, `, result.stdout);
    } catch (error) {
      console.log(`close port 3978 failed, cause by: `, error);
    }
    try {
      const result = await killPort(9239);
      console.log(`close port 9239 successfully, `, result.stdout);
    } catch (error) {
      console.log(`close port 9239 failed, cause by: `, error);
    }
    console.log("[success] debug successfully !!!");
  }

  static async dockerBuild(
    projectPath: string,
    folder: string,
    path = "./",
    processEnv: NodeJS.ProcessEnv = process.env,
    delay: number = 3 * 60 * 1000
  ): Promise<ChildProcessWithoutNullStreams> {
    console.log(`[start] docker build ... `);
    const timeout = timeoutPromise(delay);
    const childProcess = spawnCommand(
      "docker",
      ["build", "-t", folder, path],
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
      },
      (data) => {
        console.log(data);
      },
      (error) => {
        console.log(error);
        if (error.includes("Error:")) {
          chai.assert.fail(error);
        }
      }
    );
    await Promise.all([timeout, childProcess]);
    console.log("[success] docker build successfully !!!");
    return childProcess;
  }

  static async dockerRun(
    projectPath: string,
    folder: string,
    processEnv: NodeJS.ProcessEnv = process.env,
    delay: number = 30 * 1000
  ): Promise<ChildProcessWithoutNullStreams> {
    console.log(`[start] docker run ... `);
    const timeout = timeoutPromise(delay);
    const childProcess = spawnCommand(
      "docker",
      ["run", "-p", "3978:80", "--env-file", ".localConfigs", folder],
      {
        cwd: projectPath,
        env: processEnv ? processEnv : process.env,
      },
      (data) => {
        console.log(data);
      },
      (error) => {
        console.log(error);
        if (error.includes("Error:")) {
          chai.assert.fail(error);
        }
      }
    );
    await Promise.all([timeout, childProcess]);
    console.log("[success] docker run successfully !!!");
    return childProcess;
  }

  static async stopAllDocker() {
    console.log(`[start] docker stop all ... `);
    let cmd = "";
    if (os.type() === "Windows_NT") {
      cmd = "docker ps -q | ForEach-Object { docker stop $_ }";
    } else {
      cmd = "docker stop $(docker ps -q)";
    }
    const { stderr, stdout } = await execAsync(cmd);
    if (stderr) {
      console.log(stderr);
    }
    console.log(stdout);
    console.log("[success] docker stop all successfully !!!");
  }
}
