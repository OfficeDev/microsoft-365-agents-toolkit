// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import { generateAppYml } from "../../../../src/core/middleware/projectMigratorV3";
import { MigrationContext } from "../../../../src/core/middleware/utils/migrationContext";
import { randomAppName } from "../../utils";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import {
  copyTestProject,
  assertFileContent,
  readOldProjectSettings,
  Constants,
  mockMigrationContext,
  assertFileContentByTemplateCompose,
  getYmlTemplates,
} from "./utils";
import { MetadataV3 } from "../../../../src/common/versionMetadata";

describe("generateAppYml-js/ts", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    await fs.ensureDir(projectPath);
    await getYmlTemplates();
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for js SSO tab", async () => {
    await copyTestProject("jsSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts SSO tab", async () => {
    await copyTestProject("jsSsoTab", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js non SSO tab", async () => {
    await copyTestProject("jsNonSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts non SSO tab", async () => {
    await copyTestProject("jsNonSsoTab", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js tab with api", async () => {
    await copyTestProject("jsTabWithApi", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts tab with api", async () => {
    await copyTestProject("jsTabWithApi", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js function bot", async () => {
    await copyTestProject("jsFunctionBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts function bot", async () => {
    await copyTestProject("jsFunctionBot", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js webapp bot", async () => {
    await copyTestProject("jsWebappBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts webapp bot", async () => {
    await copyTestProject("jsWebappBot", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js webapp bot as resourceId eq botWebAppResourceId", async () => {
    await copyTestProject("jsWebappBot_botWebAppId", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts webapp bot as resourceId eq botWebAppResourceId", async () => {
    await copyTestProject("jsWebappBot_botWebAppId", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js function bot as resourceId eq botWebAppResourceId", async () => {
    await copyTestProject("jsFuncBot_botWebAppId", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts function bot as resourceId eq botWebAppResourceId", async () => {
    await copyTestProject("jsFuncBot_botWebAppId", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js webApp bot as resourceId eq webAppResourceId", async () => {
    await copyTestProject("jsWebappBot_webAppId", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts webApp bot as resourceId eq webAppResourceId", async () => {
    await copyTestProject("jsWebappBot_webAppId", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });

  it("should success for js function bot as resourceId eq webAppResourceId", async () => {
    await copyTestProject("jsFuncBot_webAppId", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "js.app.yml");
  });

  it("should success for ts function bot as resourceId eq webAppResourceId", async () => {
    await copyTestProject("jsFuncBot_webAppId", projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.programmingLanguage = "typescript";
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "ts.app.yml");
  });
});

describe("generateAppYml-m365", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  ["transparent-m365-tab", "transparent-m365-me"].forEach((testCase) => {
    it(testCase, async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      await copyTestProject(path.join("debug", testCase), projectPath);

      await generateAppYml(migrationContext);

      await assertFileContentByTemplateCompose(
        projectPath,
        MetadataV3.configFile,
        "app.yml",
        "expected"
      );
    });
  });
});

describe("generateAppYml-csharp", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    migrationContext.arguments.push({
      platform: "vs",
    });
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for sso tab project", async () => {
    await copyTestProject("csharpSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, MetadataV3.configFile, "app.yml");
  });

  it("should success for non-sso tab project", async () => {
    await copyTestProject("csharpNonSsoTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, MetadataV3.configFile, "app.yml");
  });

  it("should success for web app bot project", async () => {
    await copyTestProject("csharpWebappBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, MetadataV3.configFile, "app.yml");
  });

  it("should success for function bot project", async () => {
    await copyTestProject("csharpFunctionBot", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContent(projectPath, MetadataV3.configFile, "app.yml");
  });
});

describe("generateAppYml-csharp", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    migrationContext.arguments.push({
      platform: "vs",
    });
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });
});

describe("generateAppYml-spfx", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  let migrationContext: MigrationContext;

  beforeEach(async () => {
    migrationContext = await mockMigrationContext(projectPath);
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success for spfx project", async () => {
    await copyTestProject("spfxTab", projectPath);

    await generateAppYml(migrationContext);

    await assertFileContentByTemplateCompose(projectPath, MetadataV3.configFile, "app.yml");
  });
});
