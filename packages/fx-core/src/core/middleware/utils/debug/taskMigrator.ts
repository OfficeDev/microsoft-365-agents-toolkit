// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  assign,
  CommentArray,
  CommentJSONValue,
  CommentObject,
  CommentSymbol,
  parse,
} from "comment-json";
import { DebugMigrationContext } from "./debugMigrationContext";
import {
  defaultNpmInstallArg,
  FolderName,
  Prerequisite,
  TaskCommand,
  TaskDefaultValue,
  TaskLabel,
  TunnelType,
} from "../../../../component/local";
import {
  createResourcesTask,
  defaultFuncSymlinkDir,
  generateLabel,
  isCommentArray,
  isCommentObject,
  OldProjectSettingsHelper,
  setUpLocalProjectsTask,
  startAuthTask,
  startBackendTask,
  startBotTask,
  startFrontendTask,
  updateLocalEnv,
  watchBackendTask,
} from "./debugV3MigrationUtils";
import { InstallToolArgs } from "../../../../component/driver/devTool/interfaces/InstallToolArgs";
import { BuildArgs } from "../../../../component/driver/interface/buildAndDeployArgs";
import { LocalCrypto } from "../../../crypto";
import * as os from "os";
import * as path from "path";
import { NodeChecker } from "../../../../component/deps-checker/internal/nodeChecker";
import { URL } from "url";

export async function migrateTransparentPrerequisite(
  context: DebugMigrationContext
): Promise<void> {
  for (const task of context.tasks) {
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.checkPrerequisites)
    ) {
      continue;
    }

    if (isCommentObject(task["args"]) && isCommentArray(task["args"]["prerequisites"])) {
      const newPrerequisites: string[] = [];
      const toolsArgs: InstallToolArgs = {};

      for (const prerequisite of task["args"]["prerequisites"]) {
        if (prerequisite === Prerequisite.nodejs) {
          newPrerequisites.push(`"${Prerequisite.nodejs}", // Validate if Node.js is installed.`);
        } else if (prerequisite === Prerequisite.m365Account) {
          newPrerequisites.push(
            `"${Prerequisite.m365Account}", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.`
          );
        } else if (prerequisite === Prerequisite.portOccupancy) {
          newPrerequisites.push(
            `"${Prerequisite.portOccupancy}", // Validate available ports to ensure those debug ones are not occupied.`
          );
        } else if (prerequisite === Prerequisite.func) {
          toolsArgs.func = {
            version: await getFuncVersion(),
            symlinkDir: defaultFuncSymlinkDir,
          };
        } else if (prerequisite === Prerequisite.devCert) {
          toolsArgs.devCert = { trust: true };
        } else if (prerequisite === Prerequisite.dotnet) {
          toolsArgs.dotnet = true;
        }
      }

      task["args"]["prerequisites"] = parse(`[
        ${newPrerequisites.join("\n  ")}
      ]`);
      if (Object.keys(toolsArgs).length > 0) {
        if (!context.appYmlConfig.deploy) {
          context.appYmlConfig.deploy = {};
        }
        context.appYmlConfig.deploy.tools = toolsArgs;
      }
    }
  }
}

export function migrateTransparentLocalTunnel(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.startLocalTunnel)
    ) {
      ++index;
      continue;
    }

    if (isCommentObject(task["args"])) {
      if (typeof task["args"]["ngrokArgs"] === "string") {
        const portNumber = getNgrokPort(task["args"]["ngrokArgs"]);
        if (portNumber) {
          task["args"] = generateLocalTunnelTaskArgs(context, portNumber);
          const comment = `{
            // Start the local tunnel service to forward public URL to local port and inspect traffic.
            // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          }`;
          const comments = task[Symbol.for("before:label") as CommentSymbol];
          comments?.splice(0, comments?.length ?? 0);
          assign(task, parse(comment));
          ++index;
          continue;
        }
      }
    }

    const comment = `{
          // Teams Toolkit now uses Dev Tunnel as default tunnel solution.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for more details.
          // If you still prefer to use ngrok, please refer to https://aka.ms/teamsfx-tasks/customize-tunnel-service to learn how to use your own tunnel service.
        }`;
    const newTask = assign(parse(comment), {
      label: task["label"],
      type: "shell",
      command:
        "echo 'Teams Toolkit now uses Dev Tunnel as default tunnel solution. For manual updates, see https://aka.ms/teamsfx-tasks/local-tunnel.' && exit 1",
      windows: {
        options: {
          shell: {
            executable: "cmd.exe",
            args: ["/d", "/c"],
          },
        },
      },
    });
    context.tasks.splice(index, 1, newTask);
    ++index;
  }
  return Promise.resolve();
}

function getNgrokPort(ngrokCommand: string): number | undefined {
  const regex = /http\s+(?<port>\d+)\s+--log=stdout\s+--log-format=logfmt\s?/gm;
  const match = regex.exec(ngrokCommand);
  if (!match) {
    return undefined;
  }
  const portNumber = Number.parseInt(match.groups?.port ?? "");
  return Number.isInteger(portNumber) ? portNumber : undefined;
}

export function migrateTransparentNpmInstall(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.npmInstall)
    ) {
      ++index;
      continue;
    }

    if (isCommentObject(task["args"]) && isCommentArray(task["args"]["projects"])) {
      for (const npmArgs of task["args"]["projects"]) {
        if (!isCommentObject(npmArgs) || !(typeof npmArgs["cwd"] === "string")) {
          continue;
        }
        const npmInstallArg: BuildArgs = { args: "install" };
        npmInstallArg.workingDirectory = npmArgs["cwd"].replace("${workspaceFolder}", ".");

        if (typeof npmArgs["npmInstallArgs"] === "string") {
          npmInstallArg.args = `install ${npmArgs["npmInstallArgs"]}`;
        } else if (
          isCommentArray(npmArgs["npmInstallArgs"]) &&
          npmArgs["npmInstallArgs"].length > 0
        ) {
          npmInstallArg.args = `install ${npmArgs["npmInstallArgs"].join(" ")}`;
        }

        if (!context.appYmlConfig.deploy) {
          context.appYmlConfig.deploy = {};
        }
        if (!context.appYmlConfig.deploy.npmCommands) {
          context.appYmlConfig.deploy.npmCommands = [];
        }
        context.appYmlConfig.deploy.npmCommands.push(npmInstallArg);
      }
    }

    if (typeof task["label"] === "string") {
      // TODO: remove preLaunchTask in launch.json
      replaceInDependsOn(task["label"], context.tasks);
    }
    context.tasks.splice(index, 1);
  }
  return Promise.resolve();
}

export function migrateSetUpTab(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.setUpTab)
    ) {
      ++index;
      continue;
    }

    let url = new URL("https://localhost:53000");
    if (isCommentObject(task["args"]) && typeof task["args"]["baseUrl"] === "string") {
      try {
        url = new URL(task["args"]["baseUrl"]);
      } catch {}
    }

    if (!context.appYmlConfig.provision.configureApp) {
      context.appYmlConfig.provision.configureApp = {};
    }
    if (!context.appYmlConfig.provision.configureApp.tab) {
      context.appYmlConfig.provision.configureApp.tab = {};
    }
    context.appYmlConfig.provision.configureApp.tab.domain = url.host;
    context.appYmlConfig.provision.configureApp.tab.endpoint = url.origin;

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    if (!context.appYmlConfig.deploy.tab) {
      context.appYmlConfig.deploy.tab = {};
    }
    context.appYmlConfig.deploy.tab.port = parseInt(url.port);

    const label = task["label"] as string;
    index = handleProvisionAndDeploy(context, index, label);
  }
  return Promise.resolve();
}

export async function migrateSetUpBot(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.setUpBot)
    ) {
      ++index;
      continue;
    }

    context.appYmlConfig.provision.bot = {
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      messagingEndpoint: `$\{{${context.placeholderMapping.botEndpoint}}}/api/messages`,
    };

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    context.appYmlConfig.deploy.bot = true;

    const envs: { [key: string]: string } = {};
    if (isCommentObject(task["args"])) {
      if (task["args"]["botId"] && typeof task["args"]["botId"] === "string") {
        envs["BOT_ID"] = task["args"]["botId"];
      }
      if (task["args"]["botPassword"] && typeof task["args"]["botPassword"] === "string") {
        const envReferencePattern = /^\$\{env:(.*)\}$/;
        const matchResult = task["args"]["botPassword"].match(envReferencePattern);
        const botPassword = matchResult ? process.env[matchResult[1]] : task["args"]["botPassword"];
        if (botPassword) {
          const cryptoProvider = new LocalCrypto(context.oldProjectSettings.projectId);
          const result = cryptoProvider.encrypt(botPassword);
          if (result.isOk()) {
            envs["SECRET_BOT_PASSWORD"] = result.value;
          }
        }
      }
      if (
        task["args"]["botMessagingEndpoint"] &&
        typeof task["args"]["botMessagingEndpoint"] === "string"
      ) {
        if (task["args"]["botMessagingEndpoint"].startsWith("http")) {
          context.appYmlConfig.provision.bot.messagingEndpoint =
            task["args"]["botMessagingEndpoint"];
        } else if (task["args"]["botMessagingEndpoint"].startsWith("/")) {
          // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
          context.appYmlConfig.provision.bot.messagingEndpoint = `$\{{${context.placeholderMapping.botEndpoint}}}${task["args"]["botMessagingEndpoint"]}`;
        }
      }
    }
    await updateLocalEnv(context.migrationContext, envs);

    const label = task["label"] as string;
    index = handleProvisionAndDeploy(context, index, label);
  }
}

export async function migrateSetUpSSO(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.setUpSSO)
    ) {
      ++index;
      continue;
    }

    if (!context.appYmlConfig.provision.registerApp) {
      context.appYmlConfig.provision.registerApp = {};
    }
    context.appYmlConfig.provision.registerApp.aad = true;

    if (!context.appYmlConfig.provision.configureApp) {
      context.appYmlConfig.provision.configureApp = {};
    }
    context.appYmlConfig.provision.configureApp.aad = true;

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    context.appYmlConfig.deploy.sso = true;

    const envs: { [key: string]: string } = {};
    if (isCommentObject(task["args"])) {
      if (task["args"]["objectId"] && typeof task["args"]["objectId"] === "string") {
        envs["AAD_APP_OBJECT_ID"] = task["args"]["objectId"];
      }
      if (task["args"]["clientId"] && typeof task["args"]["clientId"] === "string") {
        envs["AAD_APP_CLIENT_ID"] = task["args"]["clientId"];
      }
      if (task["args"]["clientSecret"] && typeof task["args"]["clientSecret"] === "string") {
        const envReferencePattern = /^\$\{env:(.*)\}$/;
        const matchResult = task["args"]["clientSecret"].match(envReferencePattern);
        const clientSecret = matchResult
          ? process.env[matchResult[1]]
          : task["args"]["clientSecret"];
        if (clientSecret) {
          const cryptoProvider = new LocalCrypto(context.oldProjectSettings.projectId);
          const result = cryptoProvider.encrypt(clientSecret);
          if (result.isOk()) {
            envs["SECRET_AAD_APP_CLIENT_SECRET"] = result.value;
          }
        }
      }
      if (
        task["args"]["accessAsUserScopeId"] &&
        typeof task["args"]["accessAsUserScopeId"] === "string"
      ) {
        envs["AAD_APP_ACCESS_AS_USER_PERMISSION_ID"] = task["args"]["accessAsUserScopeId"];
      }
    }
    await updateLocalEnv(context.migrationContext, envs);

    const label = task["label"] as string;
    index = handleProvisionAndDeploy(context, index, label);
  }
}

export function migratePrepareManifest(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === TaskCommand.prepareManifest)
    ) {
      ++index;
      continue;
    }

    let appPackagePath: string | undefined = undefined;
    if (isCommentObject(task["args"]) && typeof task["args"]["appPackagePath"] === "string") {
      appPackagePath = task["args"]["appPackagePath"];
    }

    if (!appPackagePath) {
      if (!context.appYmlConfig.provision.registerApp) {
        context.appYmlConfig.provision.registerApp = {};
      }
      context.appYmlConfig.provision.registerApp.teamsApp = true;
    }

    if (!context.appYmlConfig.provision.configureApp) {
      context.appYmlConfig.provision.configureApp = {};
    }
    if (!context.appYmlConfig.provision.configureApp.teamsApp) {
      context.appYmlConfig.provision.configureApp.teamsApp = {};
    }
    context.appYmlConfig.provision.configureApp.teamsApp.appPackagePath = appPackagePath;

    const label = task["label"] as string;
    index = handleProvisionAndDeploy(context, index, label);
  }
  return Promise.resolve();
}

export function migrateInstallAppInTeams(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(typeof task["command"] === "string") ||
      !task["command"].includes("${command:fx-extension.install-app-in-teams}")
    ) {
      ++index;
      continue;
    }

    const label = task["label"];
    if (typeof label === "string") {
      replaceInDependsOn(label, context.tasks);
    }
    context.tasks.splice(index, 1);
  }
  return Promise.resolve();
}

export async function migrateValidateDependencies(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(typeof task["command"] === "string") ||
      !task["command"].includes("${command:fx-extension.validate-dependencies}")
    ) {
      ++index;
      continue;
    }

    const newTask = generatePrerequisiteTask(task, context);

    context.tasks.splice(index, 1, newTask);
    ++index;

    const toolsArgs: InstallToolArgs = {};
    if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
      toolsArgs.devCert = {
        trust: true,
      };
      if (OldProjectSettingsHelper.includeSSO(context.oldProjectSettings)) {
        toolsArgs.dotnet = true;
      }
    }
    if (OldProjectSettingsHelper.includeFunction(context.oldProjectSettings)) {
      toolsArgs.func = {
        version: await getFuncVersion(),
        symlinkDir: defaultFuncSymlinkDir,
      };
      toolsArgs.dotnet = true;
    }
    if (Object.keys(toolsArgs).length > 0) {
      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      context.appYmlConfig.deploy.tools = toolsArgs;
    }
  }
}

export function migrateBackendExtensionsInstall(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(
        typeof task["command"] === "string" &&
        task["command"].includes("${command:fx-extension.backend-extensions-install}")
      )
    ) {
      ++index;
      continue;
    }

    if (!context.appYmlConfig.deploy) {
      context.appYmlConfig.deploy = {};
    }
    context.appYmlConfig.deploy.dotnetCommand = {
      args: "build extensions.csproj -o ./bin --ignore-failed-sources",
      workingDirectory: `${FolderName.Function}`,
      execPath: "${{DOTNET_PATH}}",
    };

    const label = task["label"];
    if (typeof label === "string") {
      replaceInDependsOn(label, context.tasks);
    }
    context.tasks.splice(index, 1);
  }
  return Promise.resolve();
}

export function migrateFrontendStart(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: frontend start") ||
        (isCommentArray(task["dependsOn"]) &&
          task["dependsOn"].includes("teamsfx: frontend start")))
    ) {
      const newLabel = generateLabel("Start frontend", getLabels(context.tasks));
      const newTask = startFrontendTask(newLabel);
      context.tasks.splice(index + 1, 0, newTask);
      replaceInDependsOn("teamsfx: frontend start", context.tasks, newLabel);

      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      context.appYmlConfig.deploy.frontendStart = {
        sso: OldProjectSettingsHelper.includeSSO(context.oldProjectSettings),
        functionName: OldProjectSettingsHelper.getFunctionName(context.oldProjectSettings),
      };
      if (!context.appYmlConfig.deploy.npmCommands) {
        context.appYmlConfig.deploy.npmCommands = [];
      }
      const existing = context.appYmlConfig.deploy.npmCommands.find(
        (value) => value.args === "install -D env-cmd"
      );
      if (!existing) {
        context.appYmlConfig.deploy.npmCommands.push({
          args: "install -D env-cmd",
          workingDirectory: ".",
        });
      }

      break;
    } else {
      ++index;
    }
  }
  return Promise.resolve();
}

export function migrateAuthStart(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: auth start") ||
        (isCommentArray(task["dependsOn"]) && task["dependsOn"].includes("teamsfx: auth start")))
    ) {
      const newLabel = generateLabel("Start auth", getLabels(context.tasks));
      const newTask = startAuthTask(newLabel);
      context.tasks.splice(index + 1, 0, newTask);
      replaceInDependsOn("teamsfx: auth start", context.tasks, newLabel);

      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      context.appYmlConfig.deploy.authStart = {
        appsettingsPath: path.join(
          os.homedir(),
          ".fx",
          "localauth",
          "appsettings.Development.json"
        ),
      };

      break;
    } else {
      ++index;
    }
  }
  return Promise.resolve();
}

export function migrateBotStart(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: bot start") ||
        (isCommentArray(task["dependsOn"]) && task["dependsOn"].includes("teamsfx: bot start")))
    ) {
      const newLabel = generateLabel("Start bot", getLabels(context.tasks));
      const newTask = startBotTask(newLabel, context.oldProjectSettings.programmingLanguage);
      context.tasks.splice(index + 1, 0, newTask);
      replaceInDependsOn("teamsfx: bot start", context.tasks, newLabel);

      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      context.appYmlConfig.deploy.botStart = {
        tab: OldProjectSettingsHelper.includeTab(context.oldProjectSettings),
        function: OldProjectSettingsHelper.includeFunction(context.oldProjectSettings),
        sso: OldProjectSettingsHelper.includeSSO(context.oldProjectSettings),
      };

      if (!context.appYmlConfig.deploy.npmCommands) {
        context.appYmlConfig.deploy.npmCommands = [];
      }
      const existing = context.appYmlConfig.deploy.npmCommands.find(
        (value) => value.args === "install -D env-cmd"
      );
      if (!existing) {
        context.appYmlConfig.deploy.npmCommands.push({
          args: "install -D env-cmd",
          workingDirectory: ".",
        });
      }

      break;
    } else {
      ++index;
    }
  }
  return Promise.resolve();
}

export function migrateBackendWatch(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: backend watch") ||
        (isCommentArray(task["dependsOn"]) && task["dependsOn"].includes("teamsfx: backend watch")))
    ) {
      const newLabel = generateLabel("Watch backend", getLabels(context.tasks));
      const newTask = watchBackendTask(newLabel);
      context.tasks.splice(index + 1, 0, newTask);
      replaceInDependsOn("teamsfx: backend watch", context.tasks, newLabel);

      break;
    } else {
      ++index;
    }
  }
  return Promise.resolve();
}

export function migrateBackendStart(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: backend start") ||
        (isCommentArray(task["dependsOn"]) && task["dependsOn"].includes("teamsfx: backend start")))
    ) {
      const newLabel = generateLabel("Start backend", getLabels(context.tasks));
      const newTask = startBackendTask(newLabel, context.oldProjectSettings.programmingLanguage);
      context.tasks.splice(index + 1, 0, newTask);
      replaceInDependsOn("teamsfx: backend start", context.tasks, newLabel);

      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      context.appYmlConfig.deploy.backendStart = true;
      if (!context.appYmlConfig.deploy.npmCommands) {
        context.appYmlConfig.deploy.npmCommands = [];
      }
      const existing = context.appYmlConfig.deploy.npmCommands.find(
        (value) => value.args === "install -D env-cmd"
      );
      if (!existing) {
        context.appYmlConfig.deploy.npmCommands.push({
          args: "install -D env-cmd",
          workingDirectory: ".",
        });
      }

      break;
    } else {
      ++index;
    }
  }
  return Promise.resolve();
}

export async function migrateValidateLocalPrerequisites(
  context: DebugMigrationContext
): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(
        typeof task["command"] === "string" &&
        task["command"].includes("${command:fx-extension.validate-local-prerequisites}")
      )
    ) {
      ++index;
      continue;
    }

    const newTask = generatePrerequisiteTask(task, context);
    context.tasks.splice(index, 1, newTask);
    ++index;

    const toolsArgs: InstallToolArgs = {};
    const npmCommands: BuildArgs[] = [];
    let dotnetCommand: BuildArgs | undefined;
    if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
      toolsArgs.devCert = {
        trust: true,
      };
      npmCommands.push({
        args: `install ${defaultNpmInstallArg}`,
        workingDirectory: `${FolderName.Frontend}`,
      });
    }

    if (OldProjectSettingsHelper.includeFunction(context.oldProjectSettings)) {
      toolsArgs.func = {
        version: await getFuncVersion(),
        symlinkDir: defaultFuncSymlinkDir,
      };
      toolsArgs.dotnet = true;
      npmCommands.push({
        args: `install ${defaultNpmInstallArg}`,
        workingDirectory: `${FolderName.Function}`,
      });
      dotnetCommand = {
        args: "build extensions.csproj -o ./bin --ignore-failed-sources",
        workingDirectory: `${FolderName.Function}`,
        execPath: "${{DOTNET_PATH}}",
      };
    }

    if (OldProjectSettingsHelper.includeBot(context.oldProjectSettings)) {
      if (OldProjectSettingsHelper.includeFuncHostedBot(context.oldProjectSettings)) {
        toolsArgs.func = {
          version: await getFuncVersion(),
          symlinkDir: defaultFuncSymlinkDir,
        };
      }
      npmCommands.push({
        args: `install ${defaultNpmInstallArg}`,
        workingDirectory: `${FolderName.Bot}`,
      });
    }

    if (Object.keys(toolsArgs).length > 0 || npmCommands.length > 0 || dotnetCommand) {
      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      if (Object.keys(toolsArgs).length > 0) {
        context.appYmlConfig.deploy.tools = toolsArgs;
      }
      if (npmCommands.length > 0) {
        context.appYmlConfig.deploy.npmCommands = npmCommands;
      }
      context.appYmlConfig.deploy.dotnetCommand = dotnetCommand;
    }
  }
}

export function migratePreDebugCheck(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "shell") ||
      !(
        typeof task["command"] === "string" &&
        task["command"].includes("${command:fx-extension.pre-debug-check}")
      )
    ) {
      ++index;
      continue;
    }

    if (!context.appYmlConfig.provision.registerApp) {
      context.appYmlConfig.provision.registerApp = {};
    }
    if (OldProjectSettingsHelper.includeSSO(context.oldProjectSettings)) {
      context.appYmlConfig.provision.registerApp.aad = true;
    }
    context.appYmlConfig.provision.registerApp.teamsApp = true;

    if (OldProjectSettingsHelper.includeBot(context.oldProjectSettings)) {
      context.appYmlConfig.provision.bot = {
        // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
        messagingEndpoint: `$\{{${context.placeholderMapping.botEndpoint}}}/api/messages`,
      };
    }

    if (!context.appYmlConfig.provision.configureApp) {
      context.appYmlConfig.provision.configureApp = {};
    }
    if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
      context.appYmlConfig.provision.configureApp.tab = {
        domain: "localhost:53000",
        endpoint: "https://localhost:53000",
      };
    }
    if (OldProjectSettingsHelper.includeSSO(context.oldProjectSettings)) {
      context.appYmlConfig.provision.configureApp.aad = true;
    }
    if (!context.appYmlConfig.provision.configureApp.teamsApp) {
      context.appYmlConfig.provision.configureApp.teamsApp = {};
    }

    const validateLocalPrerequisitesTask = context.tasks.find(
      (_task) =>
        isCommentObject(_task) &&
        _task["type"] === "shell" &&
        typeof _task["command"] === "string" &&
        _task["command"].includes("${command:fx-extension.validate-local-prerequisites}")
    );
    if (validateLocalPrerequisitesTask) {
      if (!context.appYmlConfig.deploy) {
        context.appYmlConfig.deploy = {};
      }
      if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
        context.appYmlConfig.deploy.tab = {
          port: 53000,
        };
      }
      if (OldProjectSettingsHelper.includeBot(context.oldProjectSettings)) {
        context.appYmlConfig.deploy.bot = true;
      }
      if (OldProjectSettingsHelper.includeSSO(context.oldProjectSettings)) {
        context.appYmlConfig.deploy.sso = true;
      }
    }

    const existingLabels = getLabels(context.tasks);
    const createResourcesLabel = generateLabel("Provision", existingLabels);
    const setUpLocalProjectsLabel = generateLabel("Deploy", existingLabels);
    task["dependsOn"] = new CommentArray(createResourcesLabel, setUpLocalProjectsLabel);
    task["dependsOrder"] = "sequence";
    const createResources = createResourcesTask(createResourcesLabel);
    context.tasks.splice(index + 1, 0, createResources);
    const setUpLocalProjects = setUpLocalProjectsTask(setUpLocalProjectsLabel);
    context.tasks.splice(index + 2, 0, setUpLocalProjects);
    delete task["type"];
    delete task["command"];
    delete task["presentation"];

    break;
  }
  return Promise.resolve();
}

export function migrateNgrokStartTask(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      isCommentObject(task) &&
      ((typeof task["dependsOn"] === "string" && task["dependsOn"] === "teamsfx: ngrok start") ||
        (isCommentArray(task["dependsOn"]) && task["dependsOn"].includes("teamsfx: ngrok start")))
    ) {
      const newTask = generateLocalTunnelTask(context);
      context.tasks.splice(index + 1, 0, newTask);
      break;
    } else {
      ++index;
    }
  }
  replaceInDependsOn("teamsfx: ngrok start", context.tasks, TaskLabel.StartLocalTunnel);
  return Promise.resolve();
}

export function migrateNgrokStartCommand(context: DebugMigrationContext): Promise<void> {
  let index = 0;
  while (index < context.tasks.length) {
    const task = context.tasks[index];
    if (
      !isCommentObject(task) ||
      !(task["type"] === "teamsfx") ||
      !(task["command"] === "ngrok start")
    ) {
      ++index;
      continue;
    }

    const newTask = generateLocalTunnelTask(context, task);
    context.tasks.splice(index, 1, newTask);
    ++index;
  }
  return Promise.resolve();
}

export function migrateGetFuncPathCommand(context: DebugMigrationContext): Promise<void> {
  const getFuncPathCommand = "${command:fx-extension.get-func-path}";
  const getFuncPathDelimiterCommand = "${command:fx-extension.get-path-delimiter}";
  for (const task of context.tasks) {
    if (!isCommentObject(task)) {
      continue;
    }

    const generateNewValue = (oldStr: string): string => {
      const newStr = oldStr.startsWith(getFuncPathCommand)
        ? oldStr.replace(
            getFuncPathCommand,
            `\${workspaceFolder}/devTools/func${getFuncPathDelimiterCommand}`
          )
        : oldStr;
      return newStr.replace(
        /\${command:fx-extension.get-func-path}/g,
        `${getFuncPathDelimiterCommand}\${workspaceFolder}/devTools/func${getFuncPathDelimiterCommand}`
      );
    };

    if (isCommentObject(task["options"]) && isCommentObject(task["options"]["env"])) {
      for (const [key, value] of Object.entries(task["options"]["env"])) {
        if (typeof value === "string") {
          task["options"]["env"][key] = generateNewValue(value);
        }
      }
    }

    const platforms = ["windows", "linux", "osx"];
    platforms.forEach((platform) => {
      if (
        isCommentObject(task[platform]) &&
        isCommentObject((task[platform] as CommentObject)["options"]) &&
        isCommentObject(((task[platform] as CommentObject)["options"] as CommentObject)["env"])
      ) {
        const envObj = ((task[platform] as CommentObject)["options"] as CommentObject)[
          "env"
        ] as CommentObject;

        for (const [key, value] of Object.entries(envObj)) {
          if (typeof value === "string") {
            envObj[key] = generateNewValue(value);
          }
        }
      }
    });
  }
  return Promise.resolve();
}

function generatePrerequisiteTask(
  task: CommentObject,
  context: DebugMigrationContext
): CommentObject {
  const comment = `{
    // Check if all required prerequisites are installed and will install them if not.
    // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
  }`;
  const newTask: CommentObject = assign(parse(comment), task) as CommentObject;

  newTask["type"] = "teamsfx";
  newTask["command"] = "debug-check-prerequisites";

  const prerequisites = [
    `"${Prerequisite.nodejs}", // Validate if Node.js is installed.`,
    `"${Prerequisite.m365Account}", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.`,
    `"${Prerequisite.portOccupancy}", // Validate available ports to ensure those debug ones are not occupied.`,
  ];
  const prerequisitesComment = `
  [
    ${prerequisites.join("\n  ")}
  ]`;

  const ports: string[] = [];
  if (OldProjectSettingsHelper.includeTab(context.oldProjectSettings)) {
    ports.push(`${TaskDefaultValue.checkPrerequisites.ports.tabService}, // tab service port`);
  }
  if (OldProjectSettingsHelper.includeBot(context.oldProjectSettings)) {
    ports.push(`${TaskDefaultValue.checkPrerequisites.ports.botService}, // bot service port`);
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.botDebug}, // bot inspector port for Node.js debugger`
    );
  }
  if (OldProjectSettingsHelper.includeFunction(context.oldProjectSettings)) {
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.backendService}, // backend service port`
    );
    ports.push(
      `${TaskDefaultValue.checkPrerequisites.ports.backendDebug}, // backend inspector port for Node.js debugger`
    );
  }
  const portsComment = `
  [
    ${ports.join("\n  ")}
  ]
  `;

  const args: { [key: string]: CommentJSONValue } = {
    prerequisites: parse(prerequisitesComment),
    portOccupancy: parse(portsComment),
  };

  newTask["args"] = args as CommentJSONValue;
  return newTask;
}

function generateLocalTunnelTask(context: DebugMigrationContext, task?: CommentObject) {
  const comment = `{
        // Start the local tunnel service to forward public URL to local port and inspect traffic.
        // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
    }`;
  const newTask = assign(task ?? parse(`{"label": "${TaskLabel.StartLocalTunnel}"}`), {
    type: "teamsfx",
    command: TaskCommand.startLocalTunnel,
    args: generateLocalTunnelTaskArgs(context),
    isBackground: true,
    problemMatcher: "$teamsfx-local-tunnel-watch",
  });
  return assign(parse(comment), newTask);
}

function generateLocalTunnelTaskArgs(
  context: DebugMigrationContext,
  portNumnber = TaskDefaultValue.startLocalTunnel.devTunnel.bot.port
): CommentJSONValue {
  const placeholderComment = `
    {
      // Keep consistency with upgraded configuration.
    }
  `;
  return assign(parse("{}"), {
    type: TunnelType.devTunnel,
    ports: [
      {
        portNumber: portNumnber,
        protocol: TaskDefaultValue.startLocalTunnel.devTunnel.bot.protocol,
        access: TaskDefaultValue.startLocalTunnel.devTunnel.bot.access,
        writeToEnvironmentFile: assign(parse(placeholderComment), {
          endpoint: context.placeholderMapping.botEndpoint,
          domain: context.placeholderMapping.botDomain,
        }),
      },
    ],
    env: "local",
  });
}

function handleProvisionAndDeploy(
  context: DebugMigrationContext,
  index: number,
  label: string
): number {
  context.tasks.splice(index, 1);

  const existingLabels = getLabels(context.tasks);

  const generatedBefore = context.generatedLabels.find((value) => value.startsWith("Provision"));
  const createResourcesLabel = generatedBefore || generateLabel("Provision", existingLabels);

  const setUpLocalProjectsLabel =
    context.generatedLabels.find((value) => value.startsWith("Deploy")) ||
    generateLabel("Deploy", existingLabels);

  if (!generatedBefore) {
    context.generatedLabels.push(createResourcesLabel);
    const createResources = createResourcesTask(createResourcesLabel);
    context.tasks.splice(index, 0, createResources);
    ++index;

    context.generatedLabels.push(setUpLocalProjectsLabel);
    const setUpLocalProjects = setUpLocalProjectsTask(setUpLocalProjectsLabel);
    context.tasks.splice(index, 0, setUpLocalProjects);
    ++index;
  }

  replaceInDependsOn(label, context.tasks, createResourcesLabel, setUpLocalProjectsLabel);

  return index;
}

function replaceInDependsOn(
  label: string,
  tasks: CommentArray<CommentJSONValue>,
  ...replacements: string[]
): void {
  for (const task of tasks) {
    if (isCommentObject(task) && task["dependsOn"]) {
      if (typeof task["dependsOn"] === "string") {
        if (task["dependsOn"] === label) {
          if (replacements.length > 0) {
            task["dependsOn"] = new CommentArray(...replacements);
          } else {
            delete task["dependsOn"];
          }
        }
      } else if (Array.isArray(task["dependsOn"])) {
        const index = task["dependsOn"].findIndex((value) => value === label);
        if (index !== -1) {
          if (replacements.length > 0 && !task["dependsOn"].includes(replacements[0])) {
            task["dependsOn"].splice(index, 1, ...replacements);
          } else {
            task["dependsOn"].splice(index, 1);
          }
        }
      }
    }
  }
}

function getLabels(tasks: CommentArray<CommentJSONValue>): string[] {
  const labels: string[] = [];
  for (const task of tasks) {
    if (isCommentObject(task) && typeof task["label"] === "string") {
      labels.push(task["label"]);
    }
  }

  return labels;
}

async function getFuncVersion(): Promise<string> {
  const nodeVersion = (await NodeChecker.getInstalledNodeVersion())?.majorVersion;
  return !nodeVersion || Number.parseInt(nodeVersion) >= 18 ? "~4.0.5174" : "4";
}
