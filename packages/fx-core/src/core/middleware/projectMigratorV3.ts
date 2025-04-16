// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import { Middleware, NextFunction } from "@feathersjs/hooks/lib";
import { AppPackageFolderName, FxError, Platform, err, ok } from "@microsoft/teamsfx-api";
import * as commentJson from "comment-json";
import * as fs from "fs-extra";
import { EOL } from "os";
import * as path from "path";
import { TOOLS } from "../../common/globalVars";
import { getLocalizedString } from "../../common/localizeUtils";
import {
  Component,
  TelemetryEvent,
  sendTelemetryErrorEvent,
  sendTelemetryEvent,
} from "../../common/telemetry";
import {
  MetadataV2,
  MetadataV3,
  MetadataV4,
  VersionSource,
  VersionState,
} from "../../common/versionMetadata";
import { MANIFEST_TEMPLATE_CONSOLIDATE } from "../../component/driver/teamsApp/constants";
import { manifestUtils } from "../../component/driver/teamsApp/utils/ManifestUtils";
import {
  AbandonedProjectError,
  MigrationError,
  NotAllowedMigrationError,
  UpgradeV3CanceledError,
  assembleError,
} from "../../error";
import { getTemplatesFolder } from "../../folder";
import { CoreHookContext } from "../types";
import { loadProjectSettingsByProjectPathV2 } from "./projectSettingsLoader";
import { VersionForMigration } from "./types";
import { FileType, replacePlaceholdersForV3 } from "./utils/MigrationUtils";
import { AppYmlGenerator } from "./utils/appYmlGenerator";
import { AppLocalYmlGenerator } from "./utils/debug/appLocalYmlGenerator";
import { HubName, LaunchBrowser, LaunchUrl } from "./utils/debug/constants";
import { DebugMigrationContext } from "./utils/debug/debugMigrationContext";
import {
  OldProjectSettingsHelper,
  getPlaceholderMappings,
  ignoreDevToolsDir,
  isCommentObject,
  launchRemote,
  readJsonCommentFile,
} from "./utils/debug/debugV3MigrationUtils";
import {
  migrateAuthStart,
  migrateBackendExtensionsInstall,
  migrateBackendStart,
  migrateBackendWatch,
  migrateBotStart,
  migrateFrontendStart,
  migrateGetFuncPathCommand,
  migrateInstallAppInTeams,
  migrateNgrokStartCommand,
  migrateNgrokStartTask,
  migratePreDebugCheck,
  migratePrepareManifest,
  migrateSetUpBot,
  migrateSetUpSSO,
  migrateSetUpTab,
  migrateTransparentLocalTunnel,
  migrateTransparentNpmInstall,
  migrateTransparentPrerequisite,
  migrateValidateDependencies,
  migrateValidateLocalPrerequisites,
} from "./utils/debug/taskMigrator";
import { MigrationContext, backupFolder } from "./utils/migrationContext";
import {
  addMissingValidDomainForManifest,
  buildEnvFileName,
  buildEnvUserFileName,
  fsReadDirSync,
  generateAppIdUri,
  getCapabilityStatus,
  getParameterFromCxt,
  getProjectVersion,
  getTemplateFolderPath,
  getTrackingIdFromPath,
  getVersionState,
  isValidDomainForBotOutputKey,
  jsonObjectNamesConvertV3,
  migrationNotificationMessage,
  outputCancelMessage,
  readAndConvertUserdata,
  readBicepContent,
  readJsonFile,
  replaceAppIdUri,
  tryExtractEnvFromUserdata,
  updateAndSaveManifestForSpfx,
} from "./utils/v3MigrationUtils";

const Constants = {
  vscodeProvisionBicepPath: "./templates/azure/provision.bicep",
  launchJsonPath: ".vscode/launch.json",
  tasksJsonPath: ".vscode/tasks.json",
  reportName: "upgradeReport.md",
  envWriteOption: {
    // .env.{env} file might be already exist, use append mode (flag: a+)
    encoding: "utf8",
    flag: "a+",
  },
  envFilePrefix: ".env.",
};

const Parameters = {
  skipUserConfirm: "skipUserConfirm",
  isNonmodalMessage: "isNonmodalMessage",
  confirmOnly: "confirmOnly",
};

const TelemetryPropertyKey = {
  button: "button",
  mode: "mode",
  upgradeVersion: "upgrade-version",
  projectId: "project-id",
  reason: "reason",
};

const TelemetryPropertyValue = {
  ok: "ok",
  learnMore: "learn-more",
  cancel: "cancel",
  modal: "modal",
  nonmodal: "nonmodal",
  confirmOnly: "confirm-only",
  skipUserConfirm: "skip-user-confirm",
  upgradeVersion: "5.0",
};

export const learnMoreLink = "https://aka.ms/teams-toolkit-5.0-upgrade";

// MigrationError provides learnMoreLink as helplink for user. Remember add related error message in learnMoreLink when adding new error.
export const errorNames = {
  manifestTemplateNotExist: "ManifestTemplateNotExist",
  manifestTemplateInvalid: "ManifestTemplateInvalid",
  aadManifestTemplateNotExist: "AadManifestTemplateNotExist",
};
const upgradeButton = () => getLocalizedString("core.option.upgrade");
export const moreInfoButton: () => string = () => getLocalizedString("core.option.moreInfo");
const migrationMessageButtons = () => [upgradeButton(), moreInfoButton()];

const telemetryProperties = {
  [TelemetryPropertyKey.upgradeVersion]: TelemetryPropertyValue.upgradeVersion,
};

type Migration = (context: MigrationContext) => Promise<void>;
export const subMigrations: Array<Migration> = [
  preMigration,
  manifestsMigration,
  generateAppYml,
  configsMigration,
  statesMigration,
  userdataMigration,
  generateApimPluginEnvContent,
  updateLaunchJson,
  azureParameterMigration,
  debugMigration,
  updateGitignore,
];

export const ProjectMigratorMWV3: Middleware = async (ctx: CoreHookContext, next: NextFunction) => {
  await setTelemetryProjectId(ctx);
  const versionForMigration = await checkVersionForMigration(ctx);
  // abandoned v3 project which will not be supported. Show user the message to create new project.
  if (versionForMigration.source === VersionSource.settings) {
    await TOOLS?.ui.showMessage(
      "warn",
      getLocalizedString("core.migrationV3.abandonedProject"),
      true
    );
    ctx.result = err(new AbandonedProjectError());
    return;
  }
  const projectPath = getParameterFromCxt(ctx, "projectPath", "");
  const isValid = await checkActiveResourcePlugins(projectPath);
  const isUpgradeable = isValid && versionForMigration.state === VersionState.upgradeable;
  if (isUpgradeable) {
    // in cli non interactive scenario, migration will return an error instead of popup dialog.
    const nonInteractive = getParameterFromCxt(ctx, "nonInteractive");
    if (nonInteractive) {
      ctx.result = err(new NotAllowedMigrationError());
      return;
    }

    const isRunMigration = await showNotification(ctx, versionForMigration);
    if (isRunMigration) {
      const isNonmodalMessage = getParameterFromCxt(ctx, Parameters.isNonmodalMessage);
      if (isNonmodalMessage) {
        const versionForMigration = await checkVersionForMigration(ctx);
        if (versionForMigration.state !== VersionState.upgradeable) {
          ctx.result = ok(undefined);
          return;
        }
      }
      const migrationContext = await MigrationContext.create(ctx);
      await wrapRunMigration(migrationContext, migrate);
      ctx.result = ok(undefined);
    }
    return;
  } else {
    // continue next step only when:
    // 1. no need to upgrade the project;
    // 2. no need to update Microsoft 365 Agents Toolkit version;
    await next();
  }
};

export async function wrapRunMigration(context: MigrationContext, exec: Migration): Promise<void> {
  try {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorMigrateStart);
    await exec(context);
    await showSummaryReport(context);
    sendTelemetryEventForUpgrade(
      Component.core,
      TelemetryEvent.ProjectMigratorMigrate,
      context.telemetryProperties
    );
  } catch (error: any) {
    const errorMessage = buildErrorMessage(error, context.currentStep);
    const fxError = assembleError(error, Component.core);
    fxError.message = errorMessage;
    sendTelemetryErrorEventForUpgrade(
      Component.core,
      TelemetryEvent.ProjectMigratorError,
      fxError,
      context.telemetryProperties
    );
    await rollbackMigration(context);
    throw error;
  }
  await context.removeFxV2();
}

export function buildErrorMessage(error: any, step?: string): string {
  let message = error.message as string;
  if (error.code === "ENOENT" && error.path) {
    const fileName = path.basename(error.path);
    message = `Missing file: ${fileName}\n${message}`;
  }
  if (step) {
    message = `MigrationStep: ${step}\n${message}`;
  }
  return message;
}

export async function rollbackMigration(context: MigrationContext): Promise<void> {
  await context.cleanModifiedPaths();
  await context.restoreBackup();
  await context.cleanBackup();
}

async function showSummaryReport(context: MigrationContext): Promise<void> {
  const summaryPath = path.join(context.backupPath, Constants.reportName);
  const templatePath = path.join(getTemplatesFolder(), "core/v3Migration", Constants.reportName);

  const content = await fs.readFile(templatePath);
  await fs.writeFile(summaryPath, content);
  await TOOLS?.ui?.openFile?.(summaryPath);
}

export async function migrate(context: MigrationContext): Promise<void> {
  for (const subMigration of subMigrations) {
    context.currentStep = subMigration.name;
    await subMigration(context);
  }
}

export async function preMigration(context: MigrationContext): Promise<void> {
  await context.backup(MetadataV2.configFolder);
}

export async function checkVersionForMigration(ctx: CoreHookContext): Promise<VersionForMigration> {
  const versionInfo = await getProjectVersion(ctx);
  const versionState = getVersionState(versionInfo);
  const platform = getParameterFromCxt(ctx, "platform", Platform.VSCode) as Platform;

  return {
    currentVersion: versionInfo.version,
    source: versionInfo.source,
    state: versionState,
    platform: platform,
  };
}

export async function generateAppYml(context: MigrationContext): Promise<void> {
  const bicepContent: string = await readBicepContent(context);
  const oldProjectSettings = await loadProjectSettings(context.projectPath);
  const appYmlGenerator = new AppYmlGenerator(
    oldProjectSettings,
    bicepContent,
    context.projectPath
  );
  const appYmlString: string = await appYmlGenerator.generateAppYml();
  await context.fsWriteFile(MetadataV4.configFile, appYmlString);
  if (oldProjectSettings.programmingLanguage?.toLowerCase() === "csharp") {
    const placeholderMappings = await getPlaceholderMappings(context);
    const appLocalYmlString: string = await appYmlGenerator.generateAppLocalYml(
      placeholderMappings
    );
    await context.fsWriteFile(MetadataV4.localConfigFile, appLocalYmlString);
  }
}

export async function updateLaunchJson(context: MigrationContext): Promise<void> {
  const launchJsonPath = path.join(context.projectPath, Constants.launchJsonPath);
  if (await fs.pathExists(launchJsonPath)) {
    await context.backup(Constants.launchJsonPath);
    let launchJsonContent = await fs.readFile(launchJsonPath, "utf8");
    const oldProjectSettings = await loadProjectSettings(context.projectPath);
    if (oldProjectSettings.isM365) {
      const jsonObject = JSON.parse(launchJsonContent);
      jsonObject.configurations.push(
        launchRemote(HubName.teams, LaunchBrowser.edge, "Edge", LaunchUrl.teamsRemote, 1)
      );
      jsonObject.configurations.push(
        launchRemote(HubName.teams, LaunchBrowser.chrome, "Chrome", LaunchUrl.teamsRemote, 1)
      );
      if (OldProjectSettingsHelper.includeTab(oldProjectSettings)) {
        jsonObject.configurations.push(
          launchRemote(HubName.outlook, LaunchBrowser.edge, "Edge", LaunchUrl.outlookRemoteTab, 2)
        );
        jsonObject.configurations.push(
          launchRemote(
            HubName.outlook,
            LaunchBrowser.chrome,
            "Chrome",
            LaunchUrl.outlookRemoteTab,
            2
          )
        );
        jsonObject.configurations.push(
          launchRemote(HubName.office, LaunchBrowser.edge, "Edge", LaunchUrl.officeRemoteTab, 3)
        );
        jsonObject.configurations.push(
          launchRemote(HubName.office, LaunchBrowser.chrome, "Chrome", LaunchUrl.officeRemoteTab, 3)
        );
      } else if (OldProjectSettingsHelper.includeBot(oldProjectSettings)) {
        jsonObject.configurations.push(
          launchRemote(HubName.outlook, LaunchBrowser.edge, "Edge", LaunchUrl.outlookRemoteBot, 2)
        );
        jsonObject.configurations.push(
          launchRemote(
            HubName.outlook,
            LaunchBrowser.chrome,
            "Chrome",
            LaunchUrl.outlookRemoteBot,
            2
          )
        );
      }
      launchJsonContent = JSON.stringify(jsonObject, null, 4);
    }
    const result = launchJsonContent
      .replace(/\${teamsAppId}/g, "${{TEAMS_APP_ID}}")
      .replace(/\${teamsAppInternalId}/g, "${{M365_APP_ID}}") // For M365 apps
      .replace(/\${localTeamsAppId}/g, "${{local:TEAMS_APP_ID}}")
      .replace(/\${localTeamsAppInternalId}/g, "${{local:M365_APP_ID}}"); // For M365 apps
    await context.fsWriteFile(Constants.launchJsonPath, result);
  }
}

async function loadProjectSettings(projectPath: string): Promise<any> {
  const oldProjectSettings = await loadProjectSettingsByProjectPathV2(projectPath);
  if (oldProjectSettings.isOk()) {
    return oldProjectSettings.value;
  } else {
    throw oldProjectSettings.error;
  }
}
export function isSPFxProject(projectSettings?: any): boolean {
  const solutionSettings = projectSettings?.solutionSettings;
  if (solutionSettings) {
    const selectedPlugins = solutionSettings.activeResourcePlugins;
    return selectedPlugins && selectedPlugins.indexOf("fx-resource-spfx") !== -1;
  }
  return false;
}
export async function manifestsMigration(context: MigrationContext): Promise<void> {
  // Check manifest existing
  const oldAppPackageFolderPath = path.join(getTemplateFolderPath(context), AppPackageFolderName);
  const oldManifestPath = path.join(oldAppPackageFolderPath, MANIFEST_TEMPLATE_CONSOLIDATE);
  const oldManifestAbsolutePath = path.join(context.projectPath, oldManifestPath);
  const oldManifestExists = await fs.pathExists(oldManifestAbsolutePath);
  if (!oldManifestExists) {
    // templates/appPackage/manifest.template.json does not exist
    throw new MigrationError(
      new Error(getLocalizedString("core.migrationV3.manifestNotExist")),
      errorNames.manifestTemplateNotExist,
      learnMoreLink
    );
  }

  // Validate manifest template
  {
    const res = await manifestUtils._readAppManifest(
      path.join(context.projectPath, oldManifestPath)
    );
    if (res.isErr()) {
      throw new MigrationError(
        new Error(getLocalizedString("core.migrationV3.manifestInvalid")),
        errorNames.manifestTemplateInvalid
      );
    }
  }

  // Backup templates/appPackage
  await context.backup(oldAppPackageFolderPath);

  // Read Bicep
  const bicepContent = await readBicepContent(context);
  if (isValidDomainForBotOutputKey(bicepContent)) {
    context.isBotValidDomain = true;
  }

  // Ensure appPackage
  await context.fsEnsureDir(AppPackageFolderName);

  // copy other files from old path to new except manifest.template.json and aad.template.json
  const oldAppPackageFiles = await fs.readdir(
    path.join(context.projectPath, oldAppPackageFolderPath)
  );
  for (const file of oldAppPackageFiles) {
    if (file !== MANIFEST_TEMPLATE_CONSOLIDATE && file !== MetadataV2.aadTemplateFileName) {
      await context.fsCopy(
        path.join(oldAppPackageFolderPath, file),
        path.join(AppPackageFolderName, file)
      );
    }
  }

  // Read capability project settings
  const projectSettings = await loadProjectSettings(context.projectPath);
  const capabilities = getCapabilityStatus(projectSettings);
  const appIdUri = generateAppIdUri(capabilities);
  const isSpfx = isSPFxProject(projectSettings);

  if (!isSpfx) {
    await addMissingValidDomainForManifest(
      path.join(context.projectPath, oldManifestPath),
      capabilities.Tab,
      capabilities.BotSso,
      context.isBotValidDomain
    );
  }
  // Read Teams app manifest and save to templates/appPackage/manifest.json
  const manifestPath = path.join(AppPackageFolderName, MetadataV3.teamsManifestFileName);
  let oldManifest = await fs.readFile(path.join(context.projectPath, oldManifestPath), "utf8");
  oldManifest = replaceAppIdUri(oldManifest, appIdUri);
  const manifest = replacePlaceholdersForV3(oldManifest, bicepContent);
  if (isSpfx) {
    await updateAndSaveManifestForSpfx(context, manifest);
  } else {
    await context.fsWriteFile(manifestPath, manifest);
  }

  // Read Microsoft Entra app manifest and save to ./aad.manifest.json
  const oldAadManifestPath = path.join(oldAppPackageFolderPath, MetadataV2.aadTemplateFileName);
  const oldAadManifestExists = await fs.pathExists(
    path.join(context.projectPath, oldAadManifestPath)
  );

  const activeResourcePlugins = projectSettings.solutionSettings.activeResourcePlugins as any[];
  const component = projectSettings.components as any[];
  const aadRequired =
    (activeResourcePlugins && activeResourcePlugins.includes("fx-resource-aad-app-for-teams")) ||
    (component &&
      component.findIndex((component, index, obj) => {
        return component.name == "aad-app";
      }) >= 0);

  if (oldAadManifestExists && aadRequired) {
    let oldAadManifest = await fs.readFile(
      path.join(context.projectPath, oldAadManifestPath),
      "utf-8"
    );
    oldAadManifest = replaceAppIdUri(oldAadManifest, appIdUri);
    const aadManifest = replacePlaceholdersForV3(oldAadManifest, bicepContent);
    await context.fsWriteFile(MetadataV3.aadManifestFileName, aadManifest);
  } else if (aadRequired && !oldAadManifestExists) {
    throw new MigrationError(
      new Error(getLocalizedString("core.migrationV3.aadManifestNotExist")),
      errorNames.aadManifestTemplateNotExist,
      learnMoreLink
    );
  }

  await context.fsRemove(oldAppPackageFolderPath);
}

export async function azureParameterMigration(context: MigrationContext): Promise<void> {
  // Ensure `.fx/configs` exists
  const configFolderPath = path.join(".fx", "configs");
  const configFolderPathExists = await context.fsPathExists(configFolderPath);
  if (!configFolderPathExists) {
    // Keep same practice now. Needs dicussion whether to throw error.
    return;
  }

  // Read Bicep
  const azureFolderPath = path.join(getTemplateFolderPath(context), "azure");
  const bicepContent = await readBicepContent(context);

  const fileNames = fsReadDirSync(context, configFolderPath);
  for (const fileName of fileNames) {
    if (!fileName.startsWith("azure.parameters.")) {
      continue;
    }

    const parameterPath = path.join(context.projectPath, configFolderPath, fileName);
    // double confirm for file existence as fsReadDirSync may return deleted file names
    if (!fs.pathExistsSync(parameterPath)) {
      continue;
    }
    const content = await fs.readFile(parameterPath, "utf-8");

    const newContent = replacePlaceholdersForV3(content, bicepContent);
    await context.fsWriteFile(path.join(azureFolderPath, fileName), newContent);
  }
}

export async function showNotification(
  ctx: CoreHookContext,
  versionForMigration: VersionForMigration
): Promise<boolean> {
  const isNonmodalMessage = getParameterFromCxt(ctx, Parameters.isNonmodalMessage);
  if (isNonmodalMessage) {
    return await showNonmodalNotification(ctx, versionForMigration);
  }
  const confirmOnly = getParameterFromCxt(ctx, Parameters.confirmOnly);
  if (confirmOnly) {
    return await showConfirmOnlyNotification(ctx);
  }
  const skipUserConfirm = getParameterFromCxt(ctx, Parameters.skipUserConfirm);
  if (skipUserConfirm) {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
      [TelemetryPropertyKey.button]: TelemetryPropertyValue.ok,
      [TelemetryPropertyKey.mode]: TelemetryPropertyValue.skipUserConfirm,
    });
    return true;
  }
  return await askUserConfirm(ctx, versionForMigration);
}

async function askUserConfirm(
  ctx: CoreHookContext,
  versionForMigration: VersionForMigration
): Promise<boolean> {
  sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotificationStart);
  let answer;
  do {
    answer = await popupMessageModal(versionForMigration);
    if (answer === moreInfoButton()) {
      TOOLS?.ui.openUrl(learnMoreLink);
      sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
        [TelemetryPropertyKey.button]: TelemetryPropertyValue.learnMore,
        [TelemetryPropertyKey.mode]: TelemetryPropertyValue.modal,
      });
    }
  } while (answer === moreInfoButton());
  if (!answer || !migrationMessageButtons().includes(answer)) {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
      [TelemetryPropertyKey.button]: TelemetryPropertyValue.cancel,
      [TelemetryPropertyKey.mode]: TelemetryPropertyValue.modal,
    });
    ctx.result = err(new UpgradeV3CanceledError());
    outputCancelMessage(versionForMigration.currentVersion, versionForMigration.platform);
    return false;
  }
  sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
    [TelemetryPropertyKey.button]: TelemetryPropertyValue.ok,
    [TelemetryPropertyKey.mode]: TelemetryPropertyValue.modal,
  });
  return true;
}

async function showNonmodalNotification(
  ctx: CoreHookContext,
  versionForMigration: VersionForMigration
): Promise<boolean> {
  sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotificationStart);
  const answer = await popupMessageNonmodal(versionForMigration);
  if (answer === moreInfoButton()) {
    TOOLS?.ui.openUrl(learnMoreLink);
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
      [TelemetryPropertyKey.button]: TelemetryPropertyValue.learnMore,
      [TelemetryPropertyKey.mode]: TelemetryPropertyValue.nonmodal,
    });
    return false;
  } else if (answer === upgradeButton()) {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
      [TelemetryPropertyKey.button]: TelemetryPropertyValue.ok,
      [TelemetryPropertyKey.mode]: TelemetryPropertyValue.nonmodal,
    });
    return true;
  }
  return false;
}

async function showConfirmOnlyNotification(ctx: CoreHookContext): Promise<boolean> {
  sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotificationStart);
  const res = await TOOLS?.ui.showMessage(
    "info",
    getLocalizedString("core.migrationV3.confirmOnly.Message"),
    true,
    "OK"
  );
  if (res?.isOk() && res.value === "OK") {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
      [TelemetryPropertyKey.button]: TelemetryPropertyValue.ok,
      [TelemetryPropertyKey.mode]: TelemetryPropertyValue.confirmOnly,
    });
    return true;
  } else {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorNotification, {
      [TelemetryPropertyKey.button]: TelemetryPropertyValue.cancel,
      [TelemetryPropertyKey.mode]: TelemetryPropertyValue.confirmOnly,
    });
    return false;
  }
}

async function popupMessageModal(
  versionForMigration: VersionForMigration
): Promise<string | undefined> {
  return await popupMessage(versionForMigration, true);
}

async function popupMessageNonmodal(
  versionForMigration: VersionForMigration
): Promise<string | undefined> {
  return await popupMessage(versionForMigration, false);
}

async function popupMessage(
  versionForMigration: VersionForMigration,
  isModal: boolean
): Promise<string | undefined> {
  const res = await TOOLS?.ui.showMessage(
    "warn",
    migrationNotificationMessage(versionForMigration),
    isModal,
    ...migrationMessageButtons()
  );
  return res?.isOk() ? res.value : undefined;
}

async function setTelemetryProjectId(context: CoreHookContext): Promise<void> {
  const projectPath = getParameterFromCxt(context, "projectPath", "");
  try {
    const projectId = await getTrackingIdFromPath(projectPath);
    telemetryProperties[TelemetryPropertyKey.projectId] = projectId;
  } catch (error) {
    // do not set trackingId if error happens
  }
}

export async function configsMigration(context: MigrationContext): Promise<void> {
  // general
  if (await context.fsPathExists(path.join(".fx", "configs"))) {
    // if ./fx/states/ exists
    const fileNames = fsReadDirSync(context, path.join(".fx", "configs")); // search all files, get file names
    for (const fileName of fileNames)
      if (fileName.startsWith("config.")) {
        const fileRegex = new RegExp("(config\\.)([a-zA-Z0-9_-]*)(\\.json)$", "g"); // config.*.json
        const fileNamesArray = fileRegex.exec(fileName);
        if (fileNamesArray != null) {
          // get envName
          const envName = fileNamesArray[2];
          // create .env.{env} file if not exist
          await context.fsEnsureDir(MetadataV3.defaultEnvironmentFolder);
          if (
            !(await context.fsPathExists(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName)
            ))
          ) {
            // create env file
            await context.fsCreateFile(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName)
            );
            // add first line of comment
            await context.fsWriteFile(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName),
              envName === "local" ? MetadataV3.envFileLocalComment : MetadataV3.envFileDevComment,
              Constants.envWriteOption
            );
          }
          const obj = await readJsonFile(
            context,
            path.join(".fx", "configs", "config." + envName + ".json")
          );
          if (obj["manifest"]) {
            const bicepContent = await readBicepContent(context);
            const teamsfx_env = fs
              .readFileSync(
                path.join(
                  context.projectPath,
                  MetadataV3.defaultEnvironmentFolder,
                  Constants.envFilePrefix + envName
                )
              )
              .toString()
              .includes("TEAMSFX_ENV=")
              ? ""
              : "TEAMSFX_ENV=" + envName + EOL + "APP_NAME_SUFFIX=" + envName + EOL;
            // convert every name and add the env name at the first line
            const envData =
              teamsfx_env +
              jsonObjectNamesConvertV3(
                obj["manifest"],
                "manifest.",
                "",
                FileType.CONFIG,
                bicepContent
              );
            await context.fsWriteFile(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName),
              envData,
              Constants.envWriteOption
            );
          }
        }
      }
  }
}

export async function statesMigration(context: MigrationContext): Promise<void> {
  // general
  if (await context.fsPathExists(path.join(".fx", "states"))) {
    // if ./fx/states/ exists
    const fileNames = fsReadDirSync(context, path.join(".fx", "states")); // search all files, get file names
    for (const fileName of fileNames)
      if (fileName.startsWith("state.")) {
        const fileRegex = new RegExp("(state\\.)([a-zA-Z0-9_-]*)(\\.json)$", "g"); // state.*.json
        const fileNamesArray = fileRegex.exec(fileName);
        if (fileNamesArray != null) {
          // get envName
          const envName = fileNamesArray[2];
          // create .env.{env} file if not exist
          await context.fsEnsureDir(MetadataV3.defaultEnvironmentFolder);
          if (
            !(await context.fsPathExists(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName)
            ))
          ) {
            // create env file
            await context.fsCreateFile(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName)
            );
            // add first line of comment
            await context.fsWriteFile(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName),
              envName === "local" ? MetadataV3.envFileLocalComment : MetadataV3.envFileDevComment,
              Constants.envWriteOption
            );
          }
          const obj = await readJsonFile(
            context,
            path.join(".fx", "states", "state." + envName + ".json")
          );
          if (obj) {
            const bicepContent = await readBicepContent(context);
            // convert every name
            const envData = jsonObjectNamesConvertV3(
              obj,
              "state.",
              "",
              FileType.STATE,
              bicepContent
            );
            await context.fsWriteFile(
              path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName),
              envData,
              Constants.envWriteOption
            );
          }
        }
      }
  }
}

export async function userdataMigration(context: MigrationContext): Promise<void> {
  const stateFolder = path.join(MetadataV2.configFolder, MetadataV2.stateFolder);
  if (!(await context.fsPathExists(stateFolder))) {
    return;
  }
  await context.fsEnsureDir(MetadataV3.defaultEnvironmentFolder);
  const stateFiles = fsReadDirSync(context, stateFolder); // search all files, get file names
  for (const stateFile of stateFiles) {
    const envName = tryExtractEnvFromUserdata(stateFile);
    if (envName) {
      // get envName
      const envFileName = buildEnvUserFileName(envName);
      const bicepContent = await readBicepContent(context);
      const envData =
        MetadataV3.secretFileComment +
        EOL +
        MetadataV3.secretComment +
        (await readAndConvertUserdata(context, path.join(stateFolder, stateFile), bicepContent));
      await context.fsWriteFile(
        path.join(MetadataV3.defaultEnvironmentFolder, envFileName),
        envData,
        Constants.envWriteOption
      );
    }
  }
}

export async function debugMigration(context: MigrationContext): Promise<void> {
  // Backup vscode/tasks.json
  await context.backup(Constants.tasksJsonPath);

  // Read .vscode/tasks.json
  const tasksJsonContent = await readJsonCommentFile(
    path.join(context.projectPath, Constants.tasksJsonPath)
  );
  if (!isCommentObject(tasksJsonContent) || !Array.isArray(tasksJsonContent["tasks"])) {
    // Invalid tasks.json content
    return;
  }

  // Migrate .vscode/tasks.json
  const migrateTaskFuncs = [
    migrateTransparentPrerequisite,
    migrateTransparentNpmInstall,
    migrateTransparentLocalTunnel,
    migrateSetUpTab,
    migrateSetUpBot,
    migrateSetUpSSO,
    migratePrepareManifest,
    migrateInstallAppInTeams,
    migrateValidateDependencies,
    migrateBackendExtensionsInstall,
    migrateFrontendStart,
    migrateAuthStart,
    migrateBotStart,
    migrateBackendWatch,
    migrateBackendStart,
    migratePreDebugCheck,
    migrateValidateLocalPrerequisites,
    migrateNgrokStartTask,
    migrateNgrokStartCommand,
    migrateGetFuncPathCommand,
  ];

  const oldProjectSettings = await loadProjectSettings(context.projectPath);
  const placeholderMappings = await getPlaceholderMappings(context);

  const debugContext = new DebugMigrationContext(
    context,
    tasksJsonContent["tasks"],
    oldProjectSettings,
    placeholderMappings
  );

  for (const func of migrateTaskFuncs) {
    await func(debugContext);
  }

  // Write .vscode/tasks.json
  await context.fsWriteFile(
    Constants.tasksJsonPath,
    commentJson.stringify(tasksJsonContent, null, 4)
  );

  // Generate app.local.yml
  const appYmlGenerator = new AppLocalYmlGenerator(
    oldProjectSettings,
    debugContext.appYmlConfig,
    placeholderMappings
  );
  const appYmlString: string = await appYmlGenerator.generateAppYml();
  await context.fsWriteFile(MetadataV3.localConfigFile, appYmlString);
}

export function checkapimPluginExists(pjSettings: any): boolean {
  if (pjSettings && pjSettings["components"]) {
    for (const obj of pjSettings["components"])
      if (Object.keys(obj).includes("name") && obj["name"] === "apim") return true;
    return false;
  } else {
    return false;
  }
}

export async function generateApimPluginEnvContent(context: MigrationContext): Promise<void> {
  // general
  if (await context.fsPathExists(path.join(".fx", "configs", "projectSettings.json"))) {
    const projectSettingsContent = fs.readJsonSync(
      path.join(context.projectPath, ".fx", "configs", "projectSettings.json")
    );
    // judge if apim plugin exists
    if (checkapimPluginExists(projectSettingsContent)) {
      const fileNames = fsReadDirSync(context, path.join(".fx", "configs"));
      for (const fileName of fileNames)
        if (fileName.startsWith("config.")) {
          const fileRegex = new RegExp("(config.)([a-zA-Z0-9_-]*)(.json)$", "g"); // config.*.json
          const fileNamesArray = fileRegex.exec(fileName);
          if (fileNamesArray != null) {
            // get envName
            const envName = fileNamesArray[2];
            if (envName != "local") {
              await context.fsEnsureDir(MetadataV3.defaultEnvironmentFolder);
              if (
                !(await context.fsPathExists(
                  path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName)
                ))
              )
                await context.fsCreateFile(
                  path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName)
                );
              const apimPluginAppendContent =
                "APIM__PUBLISHEREMAIL= # To protect your privacy, Microsoft 365 Agents Toolkit does not store your email address. Please enter your email address below before provisioning to avoid failures." +
                EOL +
                "APIM__PUBLISHERNAME= # To protect your privacy, Microsoft 365 Agents Toolkit does not store your name. Please enter your name below before provisioning to avoid failures." +
                EOL;
              await context.fsWriteFile(
                path.join(MetadataV3.defaultEnvironmentFolder, Constants.envFilePrefix + envName),
                apimPluginAppendContent,
                Constants.envWriteOption
              );
            }
          }
        }
    }
  }
}

export async function checkActiveResourcePlugins(projectPath: string): Promise<boolean> {
  try {
    const projectSettings = await loadProjectSettings(projectPath);
    const solutionSettings = projectSettings.solutionSettings;
    if (projectSettings && solutionSettings && solutionSettings.activeResourcePlugins) {
      return true;
    } else {
      sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorPrecheckFailed, {
        [TelemetryPropertyKey.reason]: "solutionSettings invalid",
      });
      return false;
    }
  } catch (error: any) {
    sendTelemetryEventForUpgrade(Component.core, TelemetryEvent.ProjectMigratorPrecheckFailed, {
      [TelemetryPropertyKey.reason]: error.message,
    });
    return false;
  }
}

export async function updateGitignore(context: MigrationContext): Promise<void> {
  const gitignoreFile = ".gitignore";
  const ignoreFileExist: boolean = await context.backup(gitignoreFile);
  if (!ignoreFileExist) {
    await context.fsCreateFile(gitignoreFile);
  }

  let ignoreFileContent: string = await fs.readFile(
    path.join(context.projectPath, gitignoreFile),
    "utf8"
  );
  ignoreFileContent += EOL + `${MetadataV3.defaultEnvironmentFolder}/${buildEnvUserFileName("*")}`;
  ignoreFileContent += EOL + `${MetadataV3.defaultEnvironmentFolder}/${buildEnvFileName("local")}`;
  ignoreFileContent += EOL + `${backupFolder}/*`;
  ignoreFileContent += EOL + ignoreDevToolsDir;

  await context.fsWriteFile(gitignoreFile, ignoreFileContent);
}

function sendTelemetryEventForUpgrade(
  component: string,
  eventName: string,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number }
): void {
  const mergedProperties = { ...properties, ...telemetryProperties };
  sendTelemetryEvent(component, eventName, mergedProperties, measurements);
}

function sendTelemetryErrorEventForUpgrade(
  component: string,
  eventName: string,
  fxError: FxError,
  properties?: { [p: string]: string }
): void {
  const mergedProperties = { ...properties, ...telemetryProperties };
  sendTelemetryErrorEvent(component, eventName, fxError, mergedProperties);
}
