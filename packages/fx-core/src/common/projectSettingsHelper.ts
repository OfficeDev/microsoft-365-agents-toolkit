// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ConfigFolderName } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import { parse } from "yaml";
import { pathUtils } from "../component/utils/pathUtils";
import { YamlFileNames } from "./versionMetadata";

export enum OfficeManifestType {
  XmlAddIn,
  MetaOsAddIn,
}

export function validateProjectSettings(projectSettings: any): string | undefined {
  if (!projectSettings) return "empty projectSettings";
  if (!projectSettings.solutionSettings) return undefined;
  const solutionSettings = projectSettings.solutionSettings;
  let validateRes = validateStringArray(solutionSettings.azureResources);
  if (validateRes) {
    return `solutionSettings.azureResources validation failed: ${validateRes}`;
  }
  validateRes = validateStringArray(solutionSettings.capabilities);
  if (validateRes) {
    return `solutionSettings.capabilities validation failed: ${validateRes}`;
  }
  validateRes = validateStringArray(solutionSettings.activeResourcePlugins);
  if (validateRes) {
    return `solutionSettings.activeResourcePlugins validation failed: ${validateRes}`;
  }

  if (projectSettings?.solutionSettings?.migrateFromV1) {
    return "The project created before v2.0.0 is only supported in the Microsoft 365 Agents Toolkit before v3.4.0.";
  }

  return undefined;
}

function validateStringArray(arr?: any, enums?: string[]) {
  if (!arr) {
    return "is undefined";
  }
  if (!Array.isArray(arr)) {
    return "is not array";
  }
  for (const element of arr) {
    if (typeof element !== "string") {
      return "array elements is not string type";
    }
    if (enums && !enums.includes(element)) {
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      return `array elements is out of scope: ${enums}`;
    }
  }
  return undefined;
}

export function isValidProject(workspacePath?: string): boolean {
  if (!workspacePath) return false;
  try {
    return isValidProjectV3(workspacePath) || isValidProjectV2(workspacePath);
  } catch (e) {
    return false;
  }
}

export function isValidOfficeAddInProject(workspacePath?: string): boolean {
  const xmlManifestList = fetchManifestList(workspacePath, OfficeManifestType.XmlAddIn);
  const metaOsManifestList = fetchManifestList(workspacePath, OfficeManifestType.MetaOsAddIn);
  try {
    if (
      xmlManifestList &&
      xmlManifestList.length > 0 &&
      (!metaOsManifestList || metaOsManifestList.length == 0)
    ) {
      return true;
    } else {
      return false;
    }
  } catch (e) {
    return false;
  }
}

export function isManifestOnlyOfficeAddinProject(workspacePath?: string): boolean {
  if (!workspacePath) return false;
  const srcPath = path.join(workspacePath, "src");
  return !fs.existsSync(srcPath);
}

export function fetchManifestList(
  workspacePath?: string,
  officeManifestType?: OfficeManifestType
): string[] | undefined {
  if (!workspacePath) return undefined;
  const list = fs.readdirSync(workspacePath);
  const manifestList = list.filter((fileName) =>
    officeManifestType == OfficeManifestType.XmlAddIn
      ? isOfficeXmlAddInManifest(fileName)
      : isOfficeMetaOsAddInManifest(fileName)
  );
  return manifestList;
}

export function isOfficeXmlAddInManifest(inputFileName: string): boolean {
  return (
    inputFileName.toLocaleLowerCase().indexOf("manifest") != -1 &&
    inputFileName.toLocaleLowerCase().endsWith(".xml")
  );
}

export function isOfficeMetaOsAddInManifest(inputFileName: string): boolean {
  return (
    inputFileName.toLocaleLowerCase().indexOf("manifest") != -1 &&
    inputFileName.toLocaleLowerCase().endsWith(".json")
  );
}

export function isValidProjectV3(workspacePath: string): boolean {
  if (isValidOfficeAddInProject(workspacePath)) {
    return false;
  }
  for (const ymlFilaName of YamlFileNames) {
    const ymlFilePath = path.join(workspacePath, ymlFilaName);
    if (fs.pathExistsSync(ymlFilePath)) {
      return true;
    }
  }
  return false;
}

export function isValidProjectV2(workspacePath: string): boolean {
  const confFolderPath = path.resolve(workspacePath, `.${ConfigFolderName}`, "configs");
  const settingsFile = path.resolve(confFolderPath, "projectSettings.json");
  if (!fs.existsSync(settingsFile)) {
    return false;
  }
  const projectSettings: any = fs.readJsonSync(settingsFile);
  if (validateProjectSettings(projectSettings)) return false;
  return true;
}

export function isVSProject(projectSettings?: any): boolean {
  return projectSettings?.programmingLanguage === "csharp";
}

export function getProjectMetadata(
  rootPath?: string | undefined
): { version?: string; projectId?: string } | undefined {
  if (!rootPath) {
    return undefined;
  }
  try {
    const ymlPath = pathUtils.getYmlFilePath(rootPath, "dev");
    if (!ymlPath || !fs.pathExistsSync(ymlPath)) {
      return undefined;
    }
    const ymlContent = fs.readFileSync(ymlPath, "utf-8");
    const ymlObject = parse(ymlContent);
    return {
      projectId: ymlObject?.projectId ? ymlObject.projectId.toString() : "",
      version: ymlObject?.version ? ymlObject.version.toString() : "",
    };
  } catch {
    return undefined;
  }
}
