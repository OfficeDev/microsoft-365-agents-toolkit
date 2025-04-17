// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { hooks } from "@feathersjs/hooks";
import {
  AppManifestUtils,
  FxError,
  IMessagingExtensionCommand,
  InputsWithProjectPath,
  ManifestCapability,
  Result,
  TeamsManifest,
  TeamsManifestConverter,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import { cloneDeep } from "lodash";
import * as path from "path";
import "reflect-metadata";
import stripBom from "strip-bom";
import { v4 } from "uuid";
import isUUID from "validator/lib/isUUID";
import { ErrorContextMW } from "../../../../common/globalVars";
import { getCapabilities as checkManifestCapabilities } from "../../../../common/projectTypeChecker";
import {
  FileNotFoundError,
  JSONSyntaxError,
  MissingEnvironmentVariablesError,
  ReadFileError,
} from "../../../../error/common";
import { CapabilityOptions } from "../../../../question/constants";
import { convertManifestTemplateToV2, convertManifestTemplateToV3 } from "../../../migrate";
import { expandEnvironmentVariable, getEnvironmentVariables } from "../../../utils/common";
import { ManifestType } from "../../../utils/envFunctionUtils";
import { DriverContext } from "../../interface/commonArgs";
import {
  Constants,
  STATIC_TABS_MAX_ITEMS,
  STATIC_TABS_TPL_EXISTING_APP,
  STATIC_TABS_TPL_V3,
} from "../constants";
import { AppStudioError } from "../errors";
import { AppStudioResultFactory } from "../results";
import { getResolvedManifest } from "./utils";

export class ManifestUtils {
  async readAppManifest(projectPath: string): Promise<Result<TeamsManifest, FxError>> {
    const filePath = this.getTeamsAppManifestPath(projectPath);
    return await this._readAppManifest(filePath);
  }

  readAppManifestSync(projectPath: string): Result<TeamsManifest, FxError> {
    const filePath = this.getTeamsAppManifestPath(projectPath);
    if (!fs.existsSync(filePath)) {
      return err(new FileNotFoundError("teamsApp", filePath));
    }
    // Be compatible with UTF8-BOM encoding
    // Avoid Unexpected token error at JSON.parse()
    let content;
    try {
      content = fs.readFileSync(filePath, { encoding: "utf-8" });
    } catch (e) {
      return err(new ReadFileError(e, "ManifestUtils"));
    }
    content = stripBom(content);
    const contentV3 = convertManifestTemplateToV3(content);
    try {
      const manifest = TeamsManifestConverter.jsonToManifest(contentV3);
      return ok(manifest);
    } catch (e) {
      return err(new JSONSyntaxError(filePath, e, "ManifestUtils"));
    }
  }

  @hooks([ErrorContextMW({ component: "ManifestUtils" })])
  async _readAppManifest(manifestTemplatePath: string): Promise<Result<TeamsManifest, FxError>> {
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new FileNotFoundError("teamsApp", manifestTemplatePath));
    }
    // Be compatible with UTF8-BOM encoding
    // Avoid Unexpected token error at JSON.parse()
    let content = await fs.readFile(manifestTemplatePath, { encoding: "utf-8" });
    content = stripBom(content);
    const contentV3 = convertManifestTemplateToV3(content);
    try {
      const manifest = TeamsManifestConverter.jsonToManifest(contentV3);
      return ok(manifest);
    } catch (e) {
      return err(new JSONSyntaxError(manifestTemplatePath, e, "ManifestUtils"));
    }
  }

  async _writeAppManifest(
    appManifest: TeamsManifest,
    manifestTemplatePath: string
  ): Promise<Result<undefined, FxError>> {
    const content = TeamsManifestConverter.manifestToJson(appManifest);
    const contentV2 = convertManifestTemplateToV2(content);
    await fs.writeFile(manifestTemplatePath, contentV2);
    return ok(undefined);
  }

  getTeamsAppManifestPath(projectPath: string): string {
    // Samples from https://github.com/OfficeDev/Microsoft-Teams-Samples have the manifest in appManifest folder
    const filePath = path.join(projectPath, "appManifest", "manifest.json");
    if (fs.existsSync(filePath)) {
      return filePath;
    }

    // Samples from https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature have the manifest in root folder
    const officeAddinManifestPath = path.join(projectPath, "manifest.json");
    if (fs.existsSync(officeAddinManifestPath)) {
      return officeAddinManifestPath;
    }

    return path.join(projectPath, "appPackage", "manifest.json");
  }

  async addCapabilities(
    inputs: InputsWithProjectPath,
    capabilities: ManifestCapability[]
  ): Promise<Result<undefined, FxError>> {
    const appManifestRes = await this._readAppManifest(inputs["addManifestPath"]);
    if (appManifestRes.isErr()) return err(appManifestRes.error);
    const appManifest = appManifestRes.value;
    for (const capability of capabilities) {
      const exceedLimit = this._capabilityExceedLimit(appManifest, capability.name);
      if (exceedLimit) {
        return err(
          AppStudioResultFactory.UserError(
            AppStudioError.CapabilityExceedLimitError.name,
            AppStudioError.CapabilityExceedLimitError.message(capability.name)
          )
        );
      }
      let staticTabIndex = appManifest.staticTabs?.length ?? 0;
      switch (capability.name) {
        case "staticTab":
          appManifest.staticTabs = appManifest.staticTabs || [];
          if (capability.snippet) {
            appManifest.staticTabs.push(capability.snippet);
          } else {
            if (capability.existingApp) {
              const template = cloneDeep(STATIC_TABS_TPL_EXISTING_APP[0]);
              // eslint-disable-next-line @typescript-eslint/restrict-plus-operands
              template.entityId = "index" + staticTabIndex;
              appManifest.staticTabs.push(template);
            } else {
              const tabManifest =
                inputs.features === CapabilityOptions.dashboardTab().id
                  ? STATIC_TABS_TPL_V3[1]
                  : STATIC_TABS_TPL_V3[0];
              const template = cloneDeep(tabManifest);
              // eslint-disable-next-line @typescript-eslint/restrict-plus-operands
              template.entityId = "index" + staticTabIndex;
              appManifest.staticTabs.push(template);
            }
            staticTabIndex++;
          }
          break;
      }
    }
    if (inputs.validDomain && !appManifest.validDomains?.includes(inputs.validDomain)) {
      appManifest.validDomains?.push(inputs.validDomain);
    }

    const content = TeamsManifestConverter.manifestToJson(appManifest);
    const contentV2 = convertManifestTemplateToV2(content);
    await fs.writeFile(inputs["addManifestPath"], contentV2);

    return ok(undefined);
  }

  _capabilityExceedLimit(manifest: TeamsManifest, capability: "staticTab"): boolean {
    switch (capability) {
      case "staticTab":
        return (
          manifest.staticTabs !== undefined && manifest.staticTabs.length >= STATIC_TABS_MAX_ITEMS
        );
    }
  }
  public getCapabilities(template: TeamsManifest): string[] {
    return checkManifestCapabilities(template);
  }

  /**
   * Get command id from composeExtensions
   * @param manifest
   */
  public getOperationIds(manifest: TeamsManifest): string[] {
    const ids: string[] = [];
    manifest.composeExtensions?.map((extension: any) => {
      extension.commands?.map((command: IMessagingExtensionCommand) => {
        ids.push(command.id);
      });
    });
    return ids;
  }

  public async getPluginFilePath(
    manifest: TeamsManifest,
    manifestPath: string
  ): Promise<Result<string, FxError>> {
    const genaric = manifest as any;
    const pluginFile = genaric.copilotExtensions
      ? genaric.copilotExtensions.plugins?.[0]?.file
      : genaric.copilotAgents?.plugins?.[0]?.file;
    if (pluginFile) {
      const plugin = path.resolve(path.dirname(manifestPath), pluginFile);
      const doesFileExist = await fs.pathExists(plugin);
      if (doesFileExist) {
        return ok(plugin);
      } else {
        return err(new FileNotFoundError("ManifestUtils", pluginFile));
      }
    } else {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.TeamsAppRequiredPropertyMissingError.name,
          AppStudioError.TeamsAppRequiredPropertyMissingError.message("plugins", manifestPath)
        )
      );
    }
  }

  async getManifestV3(
    manifestTemplatePath: string,
    context: DriverContext,
    generateIdIfNotResolved = true
  ): Promise<Result<TeamsManifest, FxError>> {
    const manifestRes = await manifestUtils._readAppManifest(manifestTemplatePath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    let manifest: TeamsManifest = manifestRes.value;

    let teamsAppId = "";
    if (generateIdIfNotResolved) {
      // Corner Case: Avoid MissingEnvironmentVariablesError for manifest.id
      teamsAppId = expandEnvironmentVariable(manifest.id);
      manifest.id = "";
    }

    const manifestTemplateString = JSON.stringify(manifest);

    // Add environment variable keys to telemetry
    const resolvedManifestRes = await getResolvedManifest(
      manifestTemplateString,
      manifestTemplatePath,
      ManifestType.TeamsManifest,
      context
    );

    if (resolvedManifestRes.isErr()) {
      return err(resolvedManifestRes.error);
    }
    const resolvedManifestString = resolvedManifestRes.value;
    manifest = JSON.parse(resolvedManifestString);

    if (generateIdIfNotResolved) {
      if (!isUUID(teamsAppId)) {
        manifest.id = v4();
      } else {
        manifest.id = teamsAppId;
      }
    }

    return ok(manifest);
  }
  extractManifestFromArchivedFile(archivedFile: Buffer): Result<TeamsManifest, FxError> {
    const zipEntries = new AdmZip(archivedFile).getEntries();
    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
    if (!manifestFile) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.FileNotFoundError.name,
          AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
        )
      );
    }
    const manifestString = manifestFile.getData().toString();
    const manifest = TeamsManifestConverter.jsonToManifest(manifestString);
    return ok(manifest);
  }

  /**
   * trim the short name in manifest to make sure it is no more than 25 length
   */
  async trimManifestShortName(
    projectPath: string,
    maxLength = 25
  ): Promise<Result<undefined, FxError>> {
    const manifestPath = this.getTeamsAppManifestPath(projectPath);
    if (await fs.pathExists(manifestPath)) {
      const manifest = await AppManifestUtils.readTeamsManifest(manifestPath);
      const shortName = manifest.name.short;
      let hasSuffix = false;
      let trimmedName = shortName;
      if (shortName.includes("${{APP_NAME_SUFFIX}}")) {
        hasSuffix = true;
        trimmedName = shortName.replace("${{APP_NAME_SUFFIX}}", "");
      }
      if (trimmedName.length <= maxLength) return ok(undefined);
      let newShortName = trimmedName.replace(/\s/g, "").slice(0, maxLength);
      if (hasSuffix) {
        newShortName += "${{APP_NAME_SUFFIX}}";
      }
      manifest.name.short = newShortName;
      await AppManifestUtils.writeTeamsManifest(manifestPath, manifest);
    }
    return ok(undefined);
  }

  async resolveLocFile(locFilePath: string): Promise<Result<string, FxError>> {
    if (!(await fs.pathExists(locFilePath))) {
      return err(new FileNotFoundError("teamsApp", locFilePath));
    }

    const locFileString = await fs.readFile(locFilePath, "utf8");
    const resolvedLocFileString = expandEnvironmentVariable(locFileString);
    const unresolvedEnvVariables = getEnvironmentVariables(resolvedLocFileString);
    if (unresolvedEnvVariables && unresolvedEnvVariables.length > 0) {
      return err(
        new MissingEnvironmentVariablesError(
          "teamsApp",
          unresolvedEnvVariables.join(","),
          locFilePath
        )
      );
    }

    return ok(resolvedLocFileString);
  }
}

export const manifestUtils = new ManifestUtils();
