// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import fs from "fs-extra";
import { MigrationContext } from "./migrationContext";
import { isObject } from "lodash";
import { FileType, namingConverterV3 } from "./MigrationUtils";
import { EOL } from "os";
import { AppPackageFolderName, Inputs, Platform } from "@microsoft/teamsfx-api";
import { CoreHookContext } from "../../types";
import semver from "semver";
import { getProjectSettingPathV2, getProjectSettingsPath } from "../projectSettingsLoader";
import {
  MetadataV2,
  MetadataV3,
  MetadataV3Abandoned,
  VersionInfo,
  VersionSource,
  VersionState,
} from "../../../common/versionMetadata";
import { VersionForMigration } from "../types";
import { getLocalizedString } from "../../../common/localizeUtils";
import { TOOLS } from "../../../common/globalVars";
import { settingsUtil } from "../../../component/utils/settingsUtil";
import * as dotenv from "dotenv";
import { manifestUtils } from "../../../component/driver/teamsApp/utils/ManifestUtils";

// read json files in states/ folder
export async function readJsonFile(context: MigrationContext, filePath: string): Promise<any> {
  const filepath = path.join(context.projectPath, filePath);
  if (await fs.pathExists(filepath)) {
    const obj = fs.readJson(filepath);
    return obj;
  }
}

// read bicep file content
export async function readBicepContent(context: MigrationContext): Promise<any> {
  const bicepFilePath = path.join(getTemplateFolderPath(context), "azure", "provision.bicep");
  const bicepFileExists = await context.fsPathExists(bicepFilePath);
  return bicepFileExists
    ? fs.readFileSync(path.join(context.projectPath, bicepFilePath), "utf8")
    : "";
}

// get template folder path
export function getTemplateFolderPath(context: MigrationContext): string {
  const inputs: Inputs = context.arguments[context.arguments.length - 1];
  return inputs.platform === Platform.VS ? "Templates" : "templates";
}

// read file names list under the given path
export function fsReadDirSync(context: MigrationContext, _path: string): string[] {
  const dirPath = path.join(context.projectPath, _path);
  return fs.readdirSync(dirPath);
}

// env variables in this list will be only convert into .env.{env} when migrating {env}.userdata
const skipList = [
  "state.fx-resource-aad-app-for-teams.clientSecret",
  "state.fx-resource-bot.botPassword",
  "state.fx-resource-apim.apimClientAADClientSecret",
  "state.fx-resource-azure-sql.adminPassword",
];

export const validDomain = {
  botWithValid: "{{state.teams-bot.validDomain}}",
  tab: "{{state.teams-tab.domain}}",
  bot: "{{state.teams-bot.domain}}",
};

// convert any obj names if can be converted (used in states and configs migration)
export function jsonObjectNamesConvertV3(
  obj: any,
  prefix: string,
  parentKeyName: string,
  filetype: FileType,
  bicepContent: any
): string {
  let returnData = "";
  if (isObject(obj)) {
    for (const keyName of Object.keys(obj)) {
      returnData +=
        parentKeyName === ""
          ? jsonObjectNamesConvertV3(obj[keyName], prefix, prefix + keyName, filetype, bicepContent)
          : jsonObjectNamesConvertV3(
              obj[keyName],
              prefix,
              parentKeyName + "." + keyName,
              filetype,
              bicepContent
            );
    }
  } else if (!skipList.includes(parentKeyName)) {
    const res = namingConverterV3(parentKeyName, filetype, bicepContent);
    if (res.isOk()) {
      let stateValue = obj;
      if (typeof obj === "string" && obj.includes("#")) {
        stateValue = `"${obj}"`;
      }
      return res.value + "=" + (stateValue.toString() as string) + EOL;
    }
  } else return "";
  return returnData;
}

export async function getProjectVersion(ctx: CoreHookContext): Promise<VersionInfo> {
  const projectPath = getParameterFromCxt(ctx, "projectPath", "");
  return await getProjectVersionFromPath(projectPath);
}

export function migrationNotificationMessage(versionForMigration: VersionForMigration): string {
  if (versionForMigration.platform === Platform.VS) {
    return getLocalizedString("core.migrationV3.VS.Message");
  }
  const res = getLocalizedString("core.migrationV3.Message");
  return res;
}

export function outputCancelMessage(version: string, platform: Platform): void {
  TOOLS?.logProvider.warning(`Upgrade cancelled.`);
  if (platform === Platform.VSCode) {
    TOOLS?.logProvider.warning(
      `Notice upgrade to new configuration files is a must-have to continue to use current version Microsoft 365 Agents Toolkit. Get more info at ${MetadataV3.v3UpgradeWikiLink}.`
    );
    TOOLS?.logProvider.warning(
      `If you want to upgrade, please run command (Teams: Upgrade project) or click the "Upgrade project" button on Microsoft 365 Agents Toolkit sidebar to trigger the upgrade.`
    );
    TOOLS?.logProvider.warning(
      `If you are not ready to upgrade, please continue to use the old version Microsoft 365 Agents Toolkit ${MetadataV2.platformVersion[platform]}.`
    );
  } else if (platform === Platform.VS) {
    TOOLS?.logProvider.warning(
      `Notice upgrade to new configuration files is a must-have to continue to use current version Microsoft 365 Agents Toolkit. Get more info at ${MetadataV3.v3UpgradeWikiLink}.`
    );
    TOOLS?.logProvider.warning(`If you want to upgrade, please trigger this command again.`);
    TOOLS?.logProvider.warning(
      `If you are not ready to upgrade, please continue to use the old version Microsoft 365 Agents Toolkit.`
    );
  } else {
    TOOLS?.logProvider.warning(
      `Notice upgrade to new configuration files is a must-have to continue to use current version Microsoft 365 Toolkit CLI. Get more info at ${MetadataV3.v3UpgradeWikiLink}.`
    );
    TOOLS?.logProvider.warning(`If you want to upgrade, please trigger this command again.`);
    TOOLS?.logProvider.warning(
      `If you are not ready to upgrade, please continue to use the old version Microsoft 365 Toolkit CLI ${MetadataV2.platformVersion[platform]}.`
    );
  }
}

export async function getProjectVersionFromPath(projectPath: string): Promise<VersionInfo> {
  const v3path = getProjectSettingsPath(projectPath);
  if (await fs.pathExists(v3path)) {
    const readSettingsResult = await settingsUtil.readSettings(projectPath, false);
    if (readSettingsResult.isOk()) {
      return {
        version: readSettingsResult.value.version || "",
        source: VersionSource.teamsapp,
      };
    } else {
      throw readSettingsResult.error;
    }
  }
  const v2path = getProjectSettingPathV2(projectPath);
  if (await fs.pathExists(v2path)) {
    const settings = await fs.readJson(v2path);
    return {
      version: settings.version || "",
      source: VersionSource.projectSettings,
    };
  }
  const abandonedPath = path.resolve(
    projectPath,
    MetadataV3Abandoned.configFolder,
    MetadataV3Abandoned.configFile
  );
  if (await fs.pathExists(abandonedPath)) {
    return {
      version: MetadataV3Abandoned.configFolder,
      source: VersionSource.settings,
    };
  }
  return {
    version: "",
    source: VersionSource.unknown,
  };
}

export async function getTrackingIdFromPath(projectPath: string): Promise<string> {
  const v3path = getProjectSettingsPath(projectPath);
  if (await fs.pathExists(v3path)) {
    const readSettingsResult = await settingsUtil.readSettings(projectPath, false);
    if (readSettingsResult.isOk()) {
      return readSettingsResult.value.trackingId;
    } else {
      return "";
    }
  }
  const v2path = getProjectSettingPathV2(projectPath);
  if (await fs.pathExists(v2path)) {
    const settings = await fs.readJson(v2path);
    if (settings.projectId) {
      return settings.projectId || "";
    }
  }
  return "";
}

export function getVersionState(info: VersionInfo): VersionState {
  if (
    info.source === VersionSource.projectSettings &&
    semver.gte(info.version, MetadataV2.projectVersion) &&
    semver.lte(info.version, MetadataV2.projectMaxVersion)
  ) {
    return VersionState.upgradeable;
  } else if (
    info.source === VersionSource.teamsapp &&
    semver.valid(info.version) &&
    semver.lt(info.version, MetadataV3.unSupprotVersion)
  ) {
    return VersionState.compatible;
  } else if (info.source === VersionSource.teamsapp && !semver.valid(info.version)) {
    return VersionState.compatible;
  }
  return VersionState.unsupported;
}

export function getParameterFromCxt(
  ctx: CoreHookContext,
  key: string,
  defaultValue?: string
): string {
  const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
  const value = (inputs[key] as string) || defaultValue || "";
  return value;
}

export function getCapabilityStatus(projectSettings: any): {
  TabSso: boolean;
  BotSso: boolean;
  Tab: boolean;
} {
  const capabilities = projectSettings.solutionSettings.capabilities;
  const tabSso = capabilities.includes("TabSSO");
  const botSso = capabilities.includes("BotSSO");
  const tab = capabilities.includes("Tab");

  return {
    TabSso: tabSso,
    BotSso: botSso,
    Tab: tab,
  };
}

export function generateAppIdUri(capabilities: { TabSso: boolean; BotSso: boolean }): string {
  if (capabilities.TabSso && !capabilities.BotSso) {
    return "api://{{state.fx-resource-frontend-hosting.domain}}/{{state.fx-resource-aad-app-for-teams.clientId}}";
  } else if (capabilities.TabSso && capabilities.BotSso) {
    return "api://{{state.fx-resource-frontend-hosting.domain}}/botid-{{state.fx-resource-bot.botId}}";
  } else if (!capabilities.TabSso && capabilities.BotSso) {
    return "api://botid-{{state.fx-resource-bot.botId}}";
  } else {
    return "api://{{state.fx-resource-aad-app-for-teams.clientId}}";
  }
}

export function replaceAppIdUri(manifest: string, appIdUri: string): string {
  const appIdUriRegex = /{{+ *state\.fx\-resource\-aad\-app\-for\-teams\.applicationIdUris *}}+/g;
  if (manifest.match(appIdUriRegex)) {
    manifest = manifest.replace(appIdUriRegex, appIdUri);
  }

  return manifest;
}

export async function readAndConvertUserdata(
  context: MigrationContext,
  filePath: string,
  bicepContent: any
): Promise<string> {
  let returnAnswer = "";

  const userdataContent = await fs.readFile(path.join(context.projectPath, filePath), "utf8");
  const secretes = dotenv.parse(userdataContent);
  for (const secreteKey of Object.keys(secretes)) {
    const res = namingConverterV3("state." + secreteKey, FileType.USERDATA, bicepContent);
    if (res.isOk()) returnAnswer += `${res.value}=${secretes[secreteKey]}${EOL}`;
  }

  return returnAnswer;
}

export async function updateAndSaveManifestForSpfx(
  context: MigrationContext,
  manifest: string
): Promise<void> {
  const remoteTemplatePath = path.join(AppPackageFolderName, MetadataV3.teamsManifestFileName);
  const localTemplatePath = path.join(AppPackageFolderName, "manifest.local.json");

  const contentRegex = /\"\{\{\^config\.isLocalDebug\}\}.*\{\{\/config\.isLocalDebug\}\}\"/g;
  const remoteRegex = /\{\{\^config\.isLocalDebug\}\}.*\{\{\/config\.isLocalDebug\}\}\{/g;
  const localRegex = /\}\{\{\#config\.isLocalDebug\}\}.*\{\{\/config\.isLocalDebug\}\}/g;

  let remoteTemplate = manifest,
    localTemplate = manifest;

  // Replace contentUrls
  const placeholders = manifest.match(contentRegex);
  if (placeholders) {
    for (const placeholder of placeholders) {
      // Replace with local and remote url
      // Will only replace if one match found
      const remoteUrl = placeholder.match(remoteRegex);
      if (remoteUrl && remoteUrl.length == 1) {
        remoteTemplate = remoteTemplate.replace(
          placeholder,
          `"${remoteUrl[0].substring(24, remoteUrl[0].length - 25)}"`
        );
      }

      const localUrl = placeholder.match(localRegex);
      if (localUrl && localUrl.length == 1) {
        localTemplate = localTemplate.replace(
          placeholder,
          `"${localUrl[0].substring(25, localUrl[0].length - 24)}"`
        );
      }
    }
  }

  await context.fsWriteFile(remoteTemplatePath, remoteTemplate);
  await context.fsWriteFile(localTemplatePath, localTemplate);
}

export function isValidDomainForBotOutputKey(bicepContent: string): boolean {
  // Match teams-bot or fx-resource-bot output obj
  const pluginRegex = new RegExp(
    "output +(\\S+) +object += +{" + // Mataches start of output declaration and capture output name. Example: output functionOutput object = {
      "[^{]*" + // Matches everything between '{' and plugin id declaration. For example: comments, extra properties. Will match multilines.
      "teamsFxPluginId: +'(teams-bot|fx-resource-bot)'" + // Matches given plugin id == teams-bot or fx-resource-bot
      "[^}]*" + // Mathches anything except '}'
      "(validDomain|domain) *:" + // Matches domain key and tries not to mismatch key and value
      "[^}]*}", // Matches until end of obj as '}'
    "g"
  );
  const outputContents = pluginRegex.exec(bicepContent);
  if (outputContents && outputContents[3] === "validDomain") {
    return true;
  } else {
    return false;
  }
}

export async function addMissingValidDomainForManifest(
  manifestPath: string,
  tab: boolean,
  bot: boolean,
  isValidDomain: boolean
): Promise<void> {
  const teamsAppManifest = (await manifestUtils._readAppManifest(manifestPath))._unsafeUnwrap();
  teamsAppManifest.validDomains = teamsAppManifest.validDomains ?? [];
  const shouldAddTabDomain = tab && !teamsAppManifest.validDomains?.includes(validDomain.tab);
  if (shouldAddTabDomain) {
    teamsAppManifest.validDomains.push(validDomain.tab);
  }
  const shouldAddBotDomain =
    bot &&
    !teamsAppManifest.validDomains?.includes(validDomain.bot) &&
    !teamsAppManifest.validDomains?.includes(validDomain.botWithValid);
  if (shouldAddBotDomain) {
    teamsAppManifest.validDomains.push(isValidDomain ? validDomain.botWithValid : validDomain.bot);
  }
  await manifestUtils._writeAppManifest(teamsAppManifest, manifestPath);
}

export function tryExtractEnvFromUserdata(filename: string): string {
  const userdataRegex = new RegExp(`([a-zA-Z0-9_-]*)\\.${MetadataV2.userdataSuffix}`, "g");
  const regRes = userdataRegex.exec(filename);
  if (regRes != null) {
    return regRes[1];
  }
  return "";
}

function buildFileName(...parts: string[]): string {
  return parts.join(".");
}
export function buildEnvFileName(envName: string): string {
  return buildFileName(MetadataV3.envFilePrefix, envName);
}

export function buildEnvUserFileName(envName: string): string {
  return buildFileName(MetadataV3.envFilePrefix, envName, MetadataV3.secretFileSuffix);
}
