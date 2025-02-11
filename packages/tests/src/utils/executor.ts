// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { execAsync, editDotEnvFile, editSWASku } from "./commonUtils";
import {
  TemplateProjectFolder,
  Capability,
  LocalDebugError,
} from "./constants";
import path from "path";
import fs from "fs-extra";
import * as os from "os";
import {
  spawn,
  ChildProcessWithoutNullStreams,
  ChildProcess,
} from "child_process";
import { expect } from "chai";
import { Env } from "./env";

export class Executor {
  static async execute(
    command: string,
    cwd: string,
    processEnv?: NodeJS.ProcessEnv,
    timeout?: number,
    skipErrorMessage?: string | undefined
  ) {
    let retryCount = 0;
    const maxRetries = 2;

    while (retryCount < maxRetries) {
      // if failed, retry. 2 times at most.
      try {
        console.log(`[Start] "${command}" in ${cwd}.`);
        const options = {
          cwd,
          env: processEnv ?? process.env,
          timeout: timeout ?? 0,
        };
        const result = await execAsync(command, options);

        if (result.stderr) {
          if (skipErrorMessage) {
            if (
              result.stderr.includes(skipErrorMessage) ||
              result.stderr.toLowerCase().includes(skipErrorMessage) ||
              result.stderr
                .toLowerCase()
                .includes(LocalDebugError.DeprecatedError)
            ) {
              console.log(`[Skip Warning] ${result.stderr}`);
              return { success: true, ...result };
            }
          }
          // the command exit with 0
          console.log(
            `[Pending] "${command}" in ${cwd} with some stderr: ${result.stderr}`
          );
          return { success: false, ...result };
        } else {
          console.log(`[Success] "${command}" in ${cwd}.`);
          return { success: true, ...result };
        }
      } catch (e: any) {
        if (e.killed && e.signal == "SIGTERM") {
          console.error(`[Failed] "${command}" in ${cwd}. Timeout and killed.`);
        } else {
          console.error(
            `[Failed] "${command}" in ${cwd} with error: ${e.message}`
          );
        }
        retryCount++;
        if (retryCount >= maxRetries) {
          return { success: false, stdout: "", stderr: e.message as string };
        }

        console.log(
          `Retrying "${command}" in ${cwd}. Attempt ${retryCount} of ${maxRetries}.`
        );
      }
    }
    console.log(`[Failed] Not executed command ${command}`);
    return { success: false, stdout: "", stderr: "" };
  }

  static async login() {
    const command = `az login --username ${Env["azureAccountName"]} --password '${Env["azureAccountPassword"]}' --tenant ${Env["azureTenantId"]}`;
    await Executor.execute(command, process.cwd());

    // set subscription
    const subscription = Env["azureSubscriptionId"];
    const setSubscriptionCommand = `az account set --subscription ${subscription}`;
    return await Executor.execute(setSubscriptionCommand, process.cwd());
  }

  static concatProcessEnv(
    processEnv: NodeJS.ProcessEnv,
    env: Record<string, string>
  ) {
    return Object.assign({}, processEnv, env);
  }

  static async createProject(
    workspace: string,
    appName: string,
    capability: Capability,
    language: ProgrammingLanguage,
    customized: Record<string, string> = {}
  ) {
    let languageParam = "";
    if (language === ProgrammingLanguage.CSharp) {
      languageParam = "--runtime dotnet";
    } else if (
      language !== ProgrammingLanguage.Common &&
      language !== ProgrammingLanguage.None
    ) {
      languageParam = `--programming-language ${language}`;
    }

    const command =
      `teamsapp new --interactive false --debug --app-name ${appName} --capability ${capability} ${languageParam} ` +
      Object.entries(customized)
        .map(([key, value]) => "--" + key + " " + value)
        .join(" ");
    return this.execute(command, workspace);
  }

  static async addEnv(workspace: string, newEnv: string, env = "dev") {
    const command = `teamsapp env add ${newEnv} --env ${env}`;
    return this.execute(command, workspace);
  }

  static async addSPFxWebPart(
    workspace: string,
    spfxFolder: string,
    webpartName: string,
    manifestPath: string,
    localManifestPath: string
  ) {
    const command =
      `teamsapp add spfx-web-part --spfx-webpart-name ${webpartName}` +
      ` --spfx-folder ${spfxFolder} --teams-manifest-file ${manifestPath}` +
      ` --local-teams-manifest-file ${localManifestPath} --interactive false `;
    return this.execute(command, workspace);
  }

  static async upgrade(workspace: string, isV3 = true) {
    const prefix = isV3 ? "teamsapp" : "teamsfx";
    const command = `${prefix} upgrade --force`;
    return this.execute(command, workspace);
  }

  static async executeCmd(
    workspace: string,
    cmd: string,
    env = "dev",
    processEnv?: NodeJS.ProcessEnv,
    npx = false,
    isV3 = true,
    skipErrorMessage?: string
  ) {
    const npxCommand = npx ? "npx " : "";
    const cliPrefix = isV3 ? "teamsapp" : "teamsfx";
    const command = `${npxCommand} ${cliPrefix} ${cmd} --env ${env}`;
    return this.execute(
      command,
      workspace,
      processEnv,
      undefined,
      skipErrorMessage
    );
  }

  static async provision(
    workspace: string,
    env = "dev",
    isV3 = true,
    skipErrorMessage?: string
  ) {
    return this.executeCmd(
      workspace,
      "provision",
      env,
      undefined,
      false,
      isV3,
      skipErrorMessage
    );
  }

  static async provisionWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "provision", env, processEnv, npx, isV3);
  }

  static async validate(
    workspace: string,
    env = "dev",
    manifestFolderName = "appPackage"
  ) {
    return this.executeCmd(
      workspace,
      `validate --manifest-file ./${manifestFolderName}/manifest.json`,
      env
    );
  }

  static async validateWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "deploy", env, processEnv, npx, isV3);
  }

  static async deploy(workspace: string, env = "dev") {
    return this.executeCmd(
      workspace,
      "deploy",
      env,
      undefined,
      undefined,
      undefined,
      LocalDebugError.WarningCapError
    );
  }

  static async deployWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "deploy", env, processEnv, npx, isV3);
  }

  static async publish(workspace: string, env = "dev") {
    return this.executeCmd(workspace, "publish", env);
  }

  static async listAppOwners(workspace: string, env = "dev") {
    return this.executeCmd(
      workspace,
      "collaborator status --interactive false"
    );
  }

  static async addAppOwner(
    workspace: string,
    email: string,
    teamsManifestFilePath: string,
    env = "dev"
  ) {
    return this.executeCmd(
      workspace,
      `collaborator grant --email ${email} -t ${teamsManifestFilePath}  --interactive false`
    );
  }

  static async publishWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "publish", env, processEnv, npx, isV3);
  }

  static async preview(workspace: string, env = "dev") {
    const skipErrorMessage =
      "Warning: If you changed the manifest file, please run".toLowerCase();
    return this.executeCmd(
      workspace,
      "preview",
      env,
      undefined,
      undefined,
      undefined,
      skipErrorMessage
    );
  }

  static debugProject(
    projectPath: string,
    env: "local" | "dev" = "local",
    v3 = true,
    processEnv: NodeJS.ProcessEnv = process.env,
    onData?: (data: string) => void,
    onError?: (data: string) => void,
    openOnly?: boolean
  ) {
    let childProcess: ChildProcess | null = null;
    console.log(`[start] ${env} debug ... `);
    childProcess = Executor.spawnCommand(
      projectPath,
      v3 ? "teamsapp" : "teamsfx",
      ["preview", v3 ? "--env" : "", v3 ? env : `--${env}`],
      onData,
      onError
    );
    return childProcess;
  }

  static async previewWithCustomizedProcessEnv(
    workspace: string,
    processEnv: NodeJS.ProcessEnv,
    env = "dev",
    npx = false,
    isV3 = true
  ) {
    return this.executeCmd(workspace, "preview", env, processEnv, npx, isV3);
  }

  static async installCLI(workspace: string, version: string, global: boolean) {
    const packageName = version.startsWith("3.")
      ? "@microsoft/teamsapp-cli"
      : "@microsoft/teamsfx-cli";
    if (global) {
      const command = `npm install -g ${packageName}@${version}`;
      return this.execute(command, workspace);
    } else {
      const command = `npm install ${packageName}@${version}`;
      return this.execute(command, workspace);
    }
  }

  static async createTemplateProject(
    appName: string,
    testFolder: string,
    template: TemplateProjectFolder,
    processEnv?: NodeJS.ProcessEnv
  ) {
    const command = `teamsapp new sample ${template} --interactive false `;
    const timeout = 100000;
    try {
      await this.execute(command, testFolder, processEnv, timeout);

      //  change original template name to appName
      await this.execute(
        `mv ./${template} ./${appName}`,
        testFolder,
        processEnv ? processEnv : process.env,
        timeout
      );

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
      console.log(message);
    } catch (e: any) {
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
    processEnv?: NodeJS.ProcessEnv,
    subFolder?: string
  ) {
    const timeout = 100000;
    let oldPath = "";
    if (subFolder) {
      oldPath = path.resolve(
        __dirname,
        "..",
        "e2e/resource",
        subFolder,
        template
      );
    } else {
      oldPath = path.resolve(__dirname, "..", "e2e/resource", template);
    }
    const newPath = path.resolve(testFolder, appName);
    try {
      await this.execute(
        `mv ${oldPath} ${newPath}`,
        testFolder,
        processEnv,
        timeout
      );
    } catch (error) {
      console.log(error);
      throw new Error(`Failed to open project: ${newPath}`);
    }
    const localEnvPath = path.resolve(testFolder, appName, "env", ".env.local");
    const remoteEnvPath = path.resolve(testFolder, appName, "env", ".env.dev");
    const azureParameter = path.resolve(
      testFolder,
      appName,
      "infra",
      "azure.parameters.json"
    );
    editDotEnvFile(localEnvPath, "TEAMS_APP_NAME", appName);
    editDotEnvFile(remoteEnvPath, "TEAMS_APP_NAME", appName);
    editSWASku(azureParameter);
    console.log(`successfully open project: ${newPath}`);
  }

  static async package(
    workspace: string,
    env = "dev",
    manifestFolderName = "appPackage"
  ) {
    return this.executeCmd(
      workspace,
      `package --manifest-file ./${manifestFolderName}/manifest.json`,
      env
    );
  }

  static startDevtunnel(
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const child = spawn(
      os.type() === "Windows_NT"
        ? "devtunnel"
        : `${os.homedir()}/bin/devtunnel`,
      ["host", "-p", "3978", "--allow-anonymous"],
      {
        env: process.env,
      }
    );
    child.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    child.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
    return child;
  }

  static deleteTunnel(
    tunnelName: string,
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const child = spawn(
      os.type() === "Windows_NT"
        ? "devtunnel"
        : `${os.homedir()}/bin/devtunnel`,
      ["delete", tunnelName, "-f"],
      {
        env: process.env,
      }
    );
    console.log("delete tunnel: ", tunnelName);
    child.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    child.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
    return child;
  }

  static deleteAllTunnel(
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ) {
    const child = spawn(
      os.type() === "Windows_NT"
        ? "devtunnel"
        : `${os.homedir()}/bin/devtunnel`,
      ["delete-all", "-f"],
      {
        env: process.env,
      }
    );
    child.stdout.on("data", (data) => {
      const dataString = data.toString();
      if (onData) {
        onData(dataString);
      }
    });
    child.stderr.on("data", (data) => {
      const dataString = data.toString();
      if (onError) {
        onError(dataString);
      }
    });
  }

  static spawnCommand(
    projectPath: string,
    command: string,
    args: string[],
    onData?: (data: string) => void,
    onError?: (data: string) => void
  ): ChildProcessWithoutNullStreams {
    const isWindows = os.type() === "Windows_NT";

    const childProcess = spawn(command, args, {
      cwd: projectPath,
      shell: isWindows,
    });

    childProcess.stdout.on("data", (data) => {
      const dataString = data.toString();
      onData && onData(dataString);
    });

    childProcess.stderr.on("data", (data) => {
      const dataString = data.toString();
      onError && onError(dataString);
    });

    childProcess.on("error", (error) => {
      onError && onError(`Failed to start process: ${error.message}`);
    });

    return childProcess;
  }

  static debugBotFunctionPreparation(projectPath: string) {
    let envFile = "";
    let tunnelName = "";
    let envContent = "";
    try {
      envFile = path.resolve(projectPath, "env", ".env.local");
      envContent = fs.readFileSync(envFile, "utf-8");
    } catch (error) {
      console.log("read file error", error);
    }
    const domainRegex = /Connect via browser: https:\/\/(\S+)/;
    const endpointRegex = /Connect via browser: (\S+)/;
    const tunnelNameRegex = /Ready to accept connections for tunnel: (\S+)/;
    const devtunnelProcess = Executor.startDevtunnel(
      (data) => {
        if (data) {
          // start devtunnel
          const domainFound = data.match(domainRegex);
          const endpointFound = data.match(endpointRegex);
          const tunnelNameFound = data.match(tunnelNameRegex);
          if (domainFound && endpointFound) {
            if (domainFound[1] && endpointFound[1]) {
              const domain = domainFound[1];
              const endpoint = endpointFound[1];
              try {
                console.log(endpoint);
                envContent += `\nBOT_ENDPOINT=${endpoint}`;
                envContent += `\nBOT_DOMAIN=${domain}`;
                envContent += `\nBOT_FUNCTION_ENDPOINT=${endpoint}`;
                envContent += `\nPROVISIONOUTPUT__BOTOUTPUT__SITEENDPOINT=${endpoint}`;
                envContent += `\nPROVISIONOUTPUT__BOTOUTPUT__VALIDDOMAIN=${domain}`;
                fs.writeFileSync(envFile, envContent);
                console.log(envContent);
              } catch (error) {
                console.log(error);
              }
            }
          }
          if (tunnelNameFound) {
            if (tunnelNameFound[1]) {
              tunnelName = tunnelNameFound[1];
              console.log(tunnelName);
            }
          }
        }
      },
      (error) => {
        console.log(error);
      }
    );
    return { devtunnelProcess, tunnelName };
  }

  static async cliPreview(projectPath: string, includeBot: boolean) {
    console.log("======= debug with cli ========");
    console.log("botFlag: ", includeBot);
    let tunnelName = "";
    let devtunnelProcess: ChildProcessWithoutNullStreams | null = null;
    let debugProcess: ChildProcess | null = null;
    if (includeBot) {
      const tunnel = Executor.debugBotFunctionPreparation(projectPath);
      tunnelName = tunnel.tunnelName;
      devtunnelProcess = tunnel.devtunnelProcess;
      await new Promise((resolve) => setTimeout(resolve, 60 * 1000));
      {
        const { success } = await Executor.provision(projectPath, "local");
        expect(success).to.be.true;
        console.log(`[Successfully] provision for ${projectPath}`);
      }
      {
        const { success } = await Executor.deploy(projectPath, "local");
        expect(success).to.be.true;
        console.log(`[Successfully] deploy for ${projectPath}`);
      }
    }
    debugProcess = Executor.debugProject(
      projectPath,
      "local",
      true,
      process.env,
      (data) => {
        if (data) {
          console.log(data);
        }
      },
      (error) => {
        const errorMsg = error.toString();
        if (
          // skip warning messages
          errorMsg.includes(LocalDebugError.WarningError)
        ) {
          console.log("[skip error] ", error);
        } else {
          expect.fail(errorMsg);
        }
      }
    );
    await new Promise((resolve) => setTimeout(resolve, 3 * 60 * 1000));
    return {
      tunnelName,
      devtunnelProcess,
      debugProcess,
    };
  }

  static async killPort(
    port: number
  ): Promise<{ stdout: string; stderr: string }> {
    // windows
    if (process.platform === "win32") {
      const command = `for /f "tokens=5" %a in ('netstat -ano ^| find ":${port}"') do taskkill /PID %a /F`;
      console.log("run command: ", command);
      const result = await execAsync(command);
      return result;
    } else {
      const command = `kill -9 $(lsof -t -i:${port})`;
      console.log("run command: ", command);
      const result = await execAsync(command);
      return result;
    }
  }

  static async closeProcess(childProcess: ChildProcess | null) {
    if (childProcess) {
      try {
        if (os.type() === "Windows_NT") {
          process.kill(-childProcess.pid);
        } else {
          console.log("kill process", childProcess.spawnargs.join(" "));
          childProcess.kill("SIGKILL");
        }
      } catch (error) {
        console.log(error);
      }
    } else {
      console.log(childProcess);
    }
  }
}
