// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import fs from "fs-extra";
import * as path from "path";
import { parse } from "yaml";
import { pathUtils } from "../component/utils/pathUtils";
import { YamlFileNames } from "./versionMetadata";

export enum OfficeManifestType {
  XmlAddIn,
  MetaOsAddIn,
}

export function isValidProject(workspacePath?: string): boolean {
  if (!workspacePath) return false;
  try {
    return isValidProjectV3(workspacePath);
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
  for (const ymlFilaName of YamlFileNames) {
    const ymlFilePath = path.join(workspacePath, ymlFilaName);
    if (fs.pathExistsSync(ymlFilePath)) {
      return true;
    }
  }
  return false;
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
