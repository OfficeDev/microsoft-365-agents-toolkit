// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import { Inputs, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as path from "path";
import { MigrationContext } from "../../../../src/core/middleware/utils/migrationContext";
import { buildEnvUserFileName } from "../../../../src/core/middleware/utils/v3MigrationUtils";
import * as Handlebars from "handlebars";
import { YamlParser } from "../../../../src/component/configManager/parser";

export async function mockMigrationContext(projectPath: string): Promise<MigrationContext> {
  const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
  inputs.projectPath = projectPath;
  const ctx = {
    arguments: [inputs],
  };
  return await MigrationContext.create(ctx);
}

export function getTestAssetsPath(projectName: string): string {
  return path.join("tests/core/middleware/testAssets/v3Migration", projectName.toString());
}

// Change CRLF to LF to avoid test failures in different OS
export function normalizeLineBreaks(content: string): string {
  return content.replace(/\r\n/g, "\n");
}

export async function assertFileContent(
  projectPath: string,
  actualFilePath: string,
  expectedFileName: string
): Promise<void> {
  const actualFileFullPath = path.join(projectPath, actualFilePath);
  const expectedFileFulePath = path.join(projectPath, "expectedResult", expectedFileName);
  assert.isTrue(await fs.pathExists(actualFileFullPath));
  const actualFileContent = normalizeLineBreaks(await fs.readFile(actualFileFullPath, "utf8"));
  const expectedFileContent = normalizeLineBreaks(await fs.readFile(expectedFileFulePath, "utf8"));
  assert.equal(actualFileContent, expectedFileContent);
}

export async function assertFileContentByTemplateCompose(
  projectPath: string,
  actualFilePath: string,
  expectedFileName: string,
  expectedFolder = "expectedResult"
): Promise<void> {
  const actualFileFullPath = path.join(projectPath, actualFilePath);
  assert.isTrue(await fs.pathExists(actualFileFullPath));
  const actualFileContent = normalizeLineBreaks(await fs.readFile(actualFileFullPath, "utf8"));
  const expectedFileContent = await loadExpectedYmlFile(
    path.join(projectPath, expectedFolder, expectedFileName)
  );
  assert.equal(actualFileContent, expectedFileContent);

  const parser = new YamlParser();
  const res = await parser.parse(path.join(projectPath, actualFilePath));
  assert.isTrue(res.isOk());
}

export async function loadExpectedYmlFile(filePath: string): Promise<string> {
  const originalExpectedContent = await fs.readFile(filePath, "utf8");
  const template = Handlebars.compile(originalExpectedContent, {
    noEscape: true,
  });
  const expectedFileContent = template(ymlTemplates);
  return normalizeLineBreaks(expectedFileContent);
}

export async function copyTestProject(projectName: string, targetPath: string): Promise<void> {
  await fs.copy(getTestAssetsPath(projectName), targetPath);
}

export async function readOldProjectSettings(projectPath: string): Promise<any> {
  return await fs.readJson(path.join(projectPath, Constants.oldProjectSettingsFilePath));
}

export async function readSettingJson(projectPath: string): Promise<any> {
  return await fs.readJson(path.join(projectPath, Constants.settingsFilePath));
}

export async function readEnvFile(projectPath: string, env: string): Promise<any> {
  return await fs.readFileSync(path.join(projectPath, ".env." + env)).toString();
}

export async function readEnvUserFile(projectPath: string, env: string): Promise<any> {
  return await fs.readFileSync(path.join(projectPath, buildEnvUserFileName(env))).toString();
}

export function getManifestPathV2(projectPath: string): string {
  return path.join(projectPath, "templates", "appPackage", "manifest.template.json");
}

export function getAction(lifecycleDefinition: Array<any>, actionName: string): any[] {
  if (lifecycleDefinition) {
    return lifecycleDefinition.filter((item) => item.uses === actionName);
  }
  return [];
}

export const ymlTemplates: Record<string, string> = {};
export async function getYmlTemplates(): Promise<Record<string, string>> {
  if (Object.keys(ymlTemplates).length > 0) {
    return ymlTemplates;
  }
  const templateFolder = getTestAssetsPath("ymlTemplates");
  const templateList = await fs.readdir(templateFolder);
  await Promise.all(
    templateList.map(async (file: string) => {
      const content = (await fs.readFile(path.join(templateFolder, file))).toString();
      ymlTemplates[path.parse(file).name] = content;
    })
  );
  return ymlTemplates;
}

export const Constants = {
  happyPathTestProject: "happyPath",
  settingsFilePath: "teamsfx/settings.json",
  oldProjectSettingsFilePath: ".fx/configs/projectSettings.json",
  manifestsMigrationHappyPath: "manifestsHappyPath",
  manifestsMigrationHappyPathWithoutAad: "manifestsHappyPathWithoutAad",
  manifestsMigrationHappyPathSpfx: "manifestsHappyPathSpfx",
  manifestsMigrationHappyPathOld: "manifestsMigrationHappyPathOld",
  launchJsonPath: ".vscode/launch.json",
  happyPathWithoutFx: "happyPath_for_needMigrateToAadManifest/happyPath_no_fx",
  happyPathAadManifestTemplateExist:
    "happyPath_for_needMigrateToAadManifest/happyPath_aadManifestTemplateExist",
  happyPathWithoutPermission: "happyPath_for_needMigrateToAadManifest/happyPath_no_permissionFile",
  happyPathAadPluginNotActive:
    "happyPath_for_needMigrateToAadManifest/happyPath_aadPluginNotActive",
  environmentFolder: "env",
};
