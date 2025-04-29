// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import { hooks } from "@feathersjs/hooks/lib";
import { err, FxError, Inputs, ok, Platform, Result } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv from "mocked-env";
import * as os from "os";
import * as path from "path";
import * as sinon from "sinon";
import { setTools } from "../../../../src/common/globalVars";
import { MetadataV3, VersionSource, VersionState } from "../../../../src/common/versionMetadata";
import { NodeChecker } from "../../../../src/component/deps-checker/internal/nodeChecker";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { settingsUtil } from "../../../../src/component/utils/settingsUtil";
import * as MigratorV3 from "../../../../src/core/middleware/projectMigratorV3";
import {
  azureParameterMigration,
  checkapimPluginExists,
  checkVersionForMigration,
  configsMigration,
  debugMigration,
  errorNames,
  generateApimPluginEnvContent,
  manifestsMigration,
  migrate,
  ProjectMigratorMWV3,
  statesMigration,
  updateLaunchJson,
  userdataMigration,
  wrapRunMigration,
} from "../../../../src/core/middleware/projectMigratorV3";
import * as loader from "../../../../src/core/middleware/projectSettingsLoader";
import { getProjectSettingsPath } from "../../../../src/core/middleware/projectSettingsLoader";
import { VersionForMigration } from "../../../../src/core/middleware/types";
import * as debugV3MigrationUtils from "../../../../src/core/middleware/utils/debug/debugV3MigrationUtils";
import {
  backupFolder,
  MigrationContext,
} from "../../../../src/core/middleware/utils/migrationContext";
import * as v3MigrationUtils from "../../../../src/core/middleware/utils/v3MigrationUtils";
import {
  buildEnvUserFileName,
  getTrackingIdFromPath,
  getVersionState,
  migrationNotificationMessage,
  outputCancelMessage,
} from "../../../../src/core/middleware/utils/v3MigrationUtils";
import { CoreHookContext } from "../../../../src/core/types";
import { NotAllowedMigrationError } from "../../../../src/error";
import { MockTools, MockUserInteraction, randomAppName } from "../../utils";
import {
  assertFileContent,
  Constants,
  copyTestProject,
  getManifestPathV2,
  getTestAssetsPath,
  getYmlTemplates,
  loadExpectedYmlFile,
  mockMigrationContext,
  readEnvFile,
  readEnvUserFile,
} from "./utils";
import { pathUtils } from "../../../../src";

let mockedEnvRestore: () => void;
const mockedId = "00000000-0000-0000-0000-000000000000";

describe("ProjectMigratorMW", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
    mockedEnvRestore = mockedEnv({
      TEAMSFX_V3: "true",
    });
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path", async () => {
    sandbox
      .stub(MockUserInteraction.prototype, "showMessage")
      .onCall(0)
      .resolves(ok("More Info"))
      .onCall(1)
      .resolves(ok("Upgrade"));
    sandbox.stub(MockUserInteraction.prototype, "openUrl").resolves(ok(true));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [ProjectMigratorMWV3],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isOk());
    } finally {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("user cancel", async () => {
    sandbox
      .stub(MockUserInteraction.prototype, "showMessage")
      .resolves(err(new Error("user cancel") as FxError));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [ProjectMigratorMWV3],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isErr());
    } finally {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("wrap run error ", async () => {
    const tools = new MockTools();
    setTools(tools);
    sandbox.stub(MigratorV3, "migrate").throws(new Error("mocker error"));
    sandbox.stub(MigratorV3, "rollbackMigration").resolves();
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    try {
      await wrapRunMigration(context, migrate);
    } catch (error) {
      assert.isTrue(error.message.includes("mocker error"));
      return;
    }
    assert.fail("should throw");
  });

  it("wrap run unhandled error ", async () => {
    const tools = new MockTools();
    setTools(tools);
    sandbox.stub(MigratorV3, "preMigration").rejects({
      code: "ENOENT",
      path: "project/mocked_file",
      message: "mocked missing file error",
    });
    MigratorV3.subMigrations[0] = MigratorV3.preMigration;
    sandbox.stub(MigratorV3, "rollbackMigration").resolves();
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    try {
      await wrapRunMigration(context, migrate);
    } catch (error) {
      assert.isTrue(error.message.includes("mocked missing file error"));
      assert.isTrue(context.currentStep === "preMigration");
      return;
    }
    assert.fail("should throw");
  });

  it("happy path run error - notAllowedMigrationError", async () => {
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [ProjectMigratorMWV3],
    });
    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true, nonInteractive: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isErr());
      assert.instanceOf((res as any).error, NotAllowedMigrationError);
    } finally {
      await fs.remove(inputs.projectPath!);
    }
  });
});

describe("ProjectMigratorMW with no TEAMSFX_V3", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
    mockedEnvRestore = mockedEnv({
      TEAMSFX_V3: "false",
    });
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("TEAMSFX_V3 is false", async () => {
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok(""));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [ProjectMigratorMWV3],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isErr());
    } finally {
      await fs.remove(inputs.projectPath!);
    }
  });
});

describe("MigrationContext", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path", async () => {
    const tools = new MockTools();
    setTools(tools);

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    let res = await context.backup(".fx");
    assert.isTrue(res);
    res = await context.backup("no-exist");
    assert.isFalse(res);
    await context.fsWriteFile("a", "test-data");
    await context.fsCopy("a", "a-copy");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "a-copy")));
    await context.fsEnsureDir("b/c");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "b/c")));
    await context.fsCreateFile("d");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "d")));
    const modifiedPaths = context.getModifiedPaths();
    assert.isTrue(modifiedPaths.includes("a"));
    assert.isTrue(modifiedPaths.includes("a-copy"));
    assert.isTrue(modifiedPaths.includes("b"));
    assert.isTrue(modifiedPaths.includes("b/c"));
    assert.isTrue(modifiedPaths.includes("d"));
    await context.fsRemove("d");
    await context.cleanModifiedPaths();
    assert.isEmpty(context.getModifiedPaths());

    context.addReport("test report");
    context.addTelemetryProperties({ testProperrty: "test property" });
    await context.restoreBackup();
    await context.cleanBackup();
  });
});

describe("manifestsMigration", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("happy path: aad manifest exists", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const oldAppPackageFolderPath = path.join(projectPath, "templates", "appPackage");
    assert.isFalse(await fs.pathExists(oldAppPackageFolderPath));

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const manifestPath = path.join(appPackageFolderPath, "manifest.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);

    const aadManifestPath = path.join(projectPath, "aad.manifest.json");
    assert.isTrue(await fs.pathExists(aadManifestPath));
    const aadManifest = (await fs.readFile(aadManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const aadManifestExpected = (
      await fs.readFile(path.join(projectPath, "expected", "aad.manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(aadManifest, aadManifestExpected);
  });

  it("happy path: spfx", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPathSpfx, projectPath);

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const oldAppPackageFolderPath = path.join(projectPath, "templates", "appPackage");
    assert.isFalse(await fs.pathExists(oldAppPackageFolderPath));

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const remoteManifestPath = path.join(appPackageFolderPath, "manifest.json");
    assert.isTrue(await fs.pathExists(remoteManifestPath));
    const remoteManifest = (await fs.readFile(remoteManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const remoteManifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(remoteManifest, remoteManifestExpeceted);

    const localManifestPath = path.join(appPackageFolderPath, "manifest.local.json");
    assert.isTrue(await fs.pathExists(localManifestPath));
    const localManifest = (await fs.readFile(localManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const localManifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.local.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(localManifest, localManifestExpeceted);
  });

  it("happy path: aad manifest does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPathWithoutAad, projectPath);

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const manifestPath = path.join(appPackageFolderPath, "manifest.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);

    const aadManifestPath = path.join(projectPath, "aad.manifest.template.json");
    assert.isFalse(await fs.pathExists(aadManifestPath));
  });

  it("happy path: project created with ttk <= 4.0.0 with single teams app manifest", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPathOld, projectPath);

    try {
      await manifestsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, errorNames.aadManifestTemplateNotExist);
    }
  });

  it("happy path: project created with ttk <= 4.0.0 with two teams app manifest", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPathOld, projectPath);
    await fs.rename(
      path.join(projectPath, "templates", "appPackage", "manifest.template.json"),
      path.join(projectPath, "templates", "appPackage", "manifest.local.template.json")
    );
    await fs.copy(
      path.join(projectPath, "templates", "appPackage", "manifest.local.template.json"),
      path.join(projectPath, "templates", "appPackage", "manifest.remote.template.json")
    );

    try {
      await manifestsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, errorNames.manifestTemplateNotExist);
    }
  });

  it("migrate manifests failed: appPackage does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(false);

    try {
      await manifestsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, errorNames.manifestTemplateNotExist);
    }
  });

  it("migrate manifests success: provision.bicep does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates", "azure", "provision.bicep"));

    // Action
    await manifestsMigration(migrationContext);

    // Assert
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));
  });

  it("migrate manifests failed: teams app manifest does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates/appPackage/manifest.template.json"));

    try {
      await manifestsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, errorNames.manifestTemplateNotExist);
      assert.equal(
        error.innerError.message,
        "templates/appPackage/manifest.template.json does not exist. You may be trying to upgrade a project created by Microsoft 365 Agents Toolkit for Visual Studio Code v3.x / Microsoft 365 Agents Toolkit CLI v0.x / Microsoft 365 Agents Toolkit for Visual Studio v17.3. Please install Microsoft 365 Agents Toolkit for Visual Studio Code v4.x / Microsoft 365 Agents Toolkit CLI v1.x / Microsoft 365 Agents Toolkit for Visual Studio v17.4 and run upgrade first."
      );
    }
  });
});

describe("manifestsMigration valid domain", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("manifest without validDomain", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    const oldManifestPath = getManifestPathV2(projectPath);
    const readRes = await manifestUtils._readAppManifest(oldManifestPath);
    assert.isTrue(readRes.isOk());
    const teamsManifest = readRes._unsafeUnwrap();
    teamsManifest.validDomains = [];
    const WriteRes = await manifestUtils._writeAppManifest(teamsManifest, oldManifestPath);
    assert.isTrue(WriteRes.isOk());
    // Action
    await manifestsMigration(migrationContext);

    // Assert

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    const manifestPath = path.join(appPackageFolderPath, "manifest.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);
  });

  it("manifest without validDomain and bicep has output key validDomain", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    sandbox.stub(v3MigrationUtils, "isValidDomainForBotOutputKey").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    const oldManifestPath = getManifestPathV2(projectPath);
    const readRes = await manifestUtils._readAppManifest(oldManifestPath);
    assert.isTrue(readRes.isOk());
    const teamsManifest = readRes._unsafeUnwrap();
    teamsManifest.validDomains = [];
    const WriteRes = await manifestUtils._writeAppManifest(teamsManifest, oldManifestPath);
    assert.isTrue(WriteRes.isOk());
    // Action
    await manifestsMigration(migrationContext);

    // Assert

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    const manifestPath = path.join(appPackageFolderPath, "manifest.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "")
      .replace(
        "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
        "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__VALIDDOMAIN"
      );
    assert.equal(manifest, manifestExpeceted);
  });

  it("manifest with bot validDomain", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    const oldManifestPath = getManifestPathV2(projectPath);
    const readRes = await manifestUtils._readAppManifest(oldManifestPath);
    assert.isTrue(readRes.isOk());
    const teamsManifest = readRes._unsafeUnwrap();
    teamsManifest.validDomains = [
      "{{state.fx-resource-frontend-hosting.domain}}",
      "{{state.fx-resource-bot.validDomain}}",
    ];
    const WriteRes = await manifestUtils._writeAppManifest(teamsManifest, oldManifestPath);
    assert.isTrue(WriteRes.isOk());
    // Action
    await manifestsMigration(migrationContext);

    // Assert

    const appPackageFolderPath = path.join(projectPath, "appPackage");
    const manifestPath = path.join(appPackageFolderPath, "manifest.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "")
      .replace(
        "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
        "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__VALIDDOMAIN"
      );
    assert.equal(manifest, manifestExpeceted);
  });
});

describe("azureParameterMigration", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("Happy Path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);

    // Action
    await azureParameterMigration(migrationContext);

    // Assert
    const azureParameterDevFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.dev.json"
    );
    const azureParameterTestFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.test.json"
    );
    assert.isTrue(await fs.pathExists(azureParameterDevFilePath));
    assert.isTrue(await fs.pathExists(azureParameterTestFilePath));
    const azureParameterExpected = await fs.readFile(
      path.join(projectPath, "expected", "azure.parameters.json"),
      "utf-8"
    );
    const azureParameterDev = await fs.readFile(azureParameterDevFilePath, "utf-8");
    const azureParameterTest = await fs.readFile(azureParameterTestFilePath, "utf-8");
    assert.equal(azureParameterDev, azureParameterExpected);
    assert.equal(azureParameterTest, azureParameterExpected);
  });

  it("migrate azure.parameter failed: .fx/config does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Action
    await azureParameterMigration(migrationContext);

    // Assert
    const azureParameterDevFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.dev.json"
    );
    assert.isFalse(await fs.pathExists(azureParameterDevFilePath));
  });

  it("migrate azure.parameter failed: provision.bicep does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    await fs.ensureDir(path.join(projectPath, ".fx", "config"));

    try {
      await azureParameterMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, "MigrationReadFileError");
      assert.equal(error.innerError.message, "templates/azure/provision.bicep does not exist");
    }
  });

  it("azure.parameter.dev.json is not existing", async () => {
    // Stub
    sandbox.stub(fs, "pathExistsSync").returns(false);

    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);

    // Action
    await azureParameterMigration(migrationContext);

    // Assert
    const azureParameterDevFilePath = path.join(
      projectPath,
      "templates",
      "azure",
      "azure.parameters.dev.json"
    );

    assert.isFalse(await fs.pathExists(azureParameterDevFilePath));
  });
});

describe("updateLaunchJson", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success in happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);

    await updateLaunchJson(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, backupFolder, ".vscode/launch.json")));
    const updatedLaunchJson = await fs.readJson(path.join(projectPath, Constants.launchJsonPath));
    assert.equal(
      updatedLaunchJson.configurations[0].url,
      "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[1].url,
      "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[2].url,
      "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[3].url,
      "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}"
    );
  });

  ["transparent-m365-tab", "transparent-m365-me"].forEach((testCase) => {
    it(testCase, async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      await copyTestProject(path.join("debug", testCase), projectPath);

      await updateLaunchJson(migrationContext);

      assert.equal(
        await fs.readFile(path.join(projectPath, ".vscode", "launch.json"), "utf-8"),
        await fs.readFile(path.join(projectPath, "expected", "launch.json"), "utf-8")
      );
    });
  });
});

describe("stateMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await statesMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, Constants.environmentFolder)));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "state.dev"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.dev"))
    );
    const testEnvContent_dev = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "dev"
    );
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "state.local"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.local"))
    );
    const testEnvContent_local = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "local"
    );
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("configMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await configsMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, Constants.environmentFolder)));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "config.dev"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.dev"))
    );
    const testEnvContent_dev = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "dev"
    );
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "config.local"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.local"))
    );
    const testEnvContent_local = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "local"
    );
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("userdataMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("happy path for userdata migration", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await userdataMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, Constants.environmentFolder)));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "userdata.dev"
    );
    assert.isTrue(
      await fs.pathExists(
        path.join(projectPath, Constants.environmentFolder, buildEnvUserFileName("dev"))
      )
    );
    const testEnvContent_dev = await readEnvUserFile(
      path.join(projectPath, Constants.environmentFolder),
      "dev"
    );
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "userdata.local"
    );
    assert.isTrue(
      await fs.pathExists(
        path.join(projectPath, Constants.environmentFolder, buildEnvUserFileName("local"))
      )
    );
    const testEnvContent_local = await readEnvUserFile(
      path.join(projectPath, Constants.environmentFolder),
      "local"
    );
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });

  it("Should successfully resolve different EOLs of userdata", async () => {
    sandbox
      .stub(fs, "readFile")
      .resolves(
        "fx-resource-aad-app-for-teams.clientSecret=abcd\nfx-resource-bot.botPassword=1234\n" as unknown as Buffer
      );

    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await userdataMigration(migrationContext);
    sandbox.restore(); // in case that assertFileContent uses readFileSync

    await assertFileContent(
      projectPath,
      path.join(Constants.environmentFolder, buildEnvUserFileName("dev")),
      "userdataenv"
    );
  });
});

describe("generateApimPluginEnvContent", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await generateApimPluginEnvContent(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, Constants.environmentFolder)));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "apimPlugin.dev"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.dev"))
    );
    const testEnvContent_dev = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "dev"
    );
    assert.equal(testEnvContent_dev, trueEnvContent_dev);
  });

  it("checkapimPluginExists: apim exists", () => {
    const pjSettings_1 = {
      appName: "testapp",
      components: [
        {
          name: "teams-tab",
        },
        {
          name: "apim",
        },
      ],
    };
    assert.isTrue(checkapimPluginExists(pjSettings_1));
  });

  it("checkapimPluginExists: apim not exists", () => {
    const pjSettings_2 = {
      appName: "testapp",
      components: [
        {
          name: "teams-tab",
        },
      ],
    };
    assert.isFalse(checkapimPluginExists(pjSettings_2));
  });

  it("checkapimPluginExists: components not exists", () => {
    const pjSettings_3 = {
      appName: "testapp",
    };
    assert.isFalse(checkapimPluginExists(pjSettings_3));
  });

  it("checkapimPluginExists: obj null", () => {
    const pjSettings_4 = null;
    assert.isFalse(checkapimPluginExists(pjSettings_4));
  });
});

describe("allEnvMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path for all env migration", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await configsMigration(migrationContext);
    await statesMigration(migrationContext);
    await userdataMigration(migrationContext);
    await generateApimPluginEnvContent(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, Constants.environmentFolder)));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "all.dev"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.dev"))
    );
    const testEnvContent_dev = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "dev"
    );
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "all.local"
    );
    assert.isTrue(
      await fs.pathExists(path.join(projectPath, Constants.environmentFolder, ".env.local"))
    );
    const testEnvContent_local = await readEnvFile(
      path.join(projectPath, Constants.environmentFolder),
      "local"
    );
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });
});

describe("Migration utils", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("checkVersionForMigration V2", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.upgradeable);
  });

  it("checkVersionForMigration V3", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox
      .stub(settingsUtil, "readSettings")
      .resolves(ok({ trackingId: "mockId", version: "1.0.0" }));
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
    sandbox.stub(fs, "readFile").resolves("version: 1.0.0" as any);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.compatible);
  });

  it("checkVersionForMigration V3 abandoned", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox.stub(loader, "getProjectSettingPathV2").returns("");
    sandbox.stub(loader, "getProjectSettingsPath").returns("");
    sandbox.stub(fs, "pathExists").callsFake(async (path) => {
      return path ? true : false;
    });
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.unsupported);
  });

  it("checkVersionForMigration empty", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox.stub(fs, "pathExists").resolves(false);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state.state, VersionState.unsupported);
  });

  it("getTrackingIdFromPath: V2 ", async () => {
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path === getProjectSettingsPath(projectPath)) {
        return false;
      }
      return true;
    });
    sandbox.stub(fs, "readJson").resolves({ projectId: mockedId });
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, mockedId);
  });

  it("getTrackingIdFromPath: V3 ", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(settingsUtil, "readSettings").resolves(
      ok({
        version: MetadataV3.projectVersion,
        trackingId: mockedId,
      })
    );
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, mockedId);
  });

  it("getTrackingIdFromPath: V3 failed", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(settingsUtil, "readSettings").resolves(err(new Error("mocked error") as FxError));
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, "");
  });

  it("getTrackingIdFromPath: empty", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, "");
  });

  it("getTrackingIdFromPath: empty", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    const trackingId = await getTrackingIdFromPath(projectPath);
    assert.equal(trackingId, "");
  });

  it("getVersionState", () => {
    assert.equal(
      getVersionState({
        version: "2.0.0",
        source: VersionSource.projectSettings,
      }),
      VersionState.upgradeable
    );
    assert.equal(
      getVersionState({
        version: "1.0.0",
        source: VersionSource.teamsapp,
      }),
      VersionState.compatible
    );
    assert.equal(
      getVersionState({
        version: "1.1.0",
        source: VersionSource.teamsapp,
      }),
      VersionState.compatible
    );
    assert.equal(
      getVersionState({
        version: "2.0.0",
        source: VersionSource.teamsapp,
      }),
      VersionState.unsupported
    );
    assert.equal(
      getVersionState({
        version: "v1.2",
        source: VersionSource.teamsapp,
      }),
      VersionState.compatible
    );
    assert.equal(
      getVersionState({
        version: "1.2",
        source: VersionSource.teamsapp,
      }),
      VersionState.compatible
    );
    assert.equal(
      getVersionState({
        version: "",
        source: VersionSource.unknown,
      }),
      VersionState.unsupported
    );
  });

  it("outputCancelMessage", () => {
    outputCancelMessage("2.0.0", Platform.VS);
    outputCancelMessage("2.0.0", Platform.CLI);
    outputCancelMessage("2.0.0", Platform.VSCode);
  });

  it("migrationNotificationMessage", () => {
    const tools = new MockTools();
    setTools(tools);

    const version: VersionForMigration = {
      currentVersion: "2.0.0",
      source: VersionSource.projectSettings,
      state: VersionState.upgradeable,
      platform: Platform.VS,
    };

    migrationNotificationMessage(version);
    version.platform = Platform.VSCode;
    migrationNotificationMessage(version);
    version.platform = Platform.CLI;
    migrationNotificationMessage(version);
  });
});

describe("Migration show notification", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();
  const inputs: Inputs = {
    platform: Platform.VSCode,
    ignoreEnvInfo: true,
    projectPath: projectPath,
  };
  const coreCtx = {
    arguments: [inputs],
  };
  const version: VersionForMigration = {
    currentVersion: "2.0.0",
    source: VersionSource.projectSettings,
    state: VersionState.upgradeable,
    platform: Platform.VSCode,
  };

  beforeEach(async () => {
    inputs["isNonmodalMessage"] = "";
    inputs["confirmOnly"] = "";
    inputs["skipUserConfirm"] = "";
    sandbox.stub(MockUserInteraction.prototype, "openUrl").resolves(ok(true));
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("nonmodal case and click upgrade", async () => {
    inputs.isNonmodalMessage = "true";
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("Upgrade"));
    const res = await MigratorV3.showNotification(coreCtx, version);
    assert.isTrue(res);
  });

  it("nonmodal case and click More Info", async () => {
    inputs.isNonmodalMessage = "true";
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("More Info"));
    const res = await MigratorV3.showNotification(coreCtx, version);
    assert.isFalse(res);
  });

  it("nonmodal case and click nothing", async () => {
    inputs.isNonmodalMessage = "true";
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok(""));
    const res = await MigratorV3.showNotification(coreCtx, version);
    assert.isFalse(res);
  });

  it("confirmOnly case and click OK", async () => {
    inputs.confirmOnly = "true";
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("OK"));
    const res = await MigratorV3.showNotification(coreCtx, version);
    assert.isTrue(res);
  });

  it("confirmOnly case and click cancel", async () => {
    inputs.confirmOnly = "true";
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("cancel"));
    const res = await MigratorV3.showNotification(coreCtx, version);
    assert.isFalse(res);
  });

  it("skipUserConfirm case", async () => {
    inputs.skipUserConfirm = "true";
    const res = await MigratorV3.showNotification(coreCtx, version);
    assert.isTrue(res);
  });
});

describe("debugMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    sinon.stub(debugV3MigrationUtils, "updateLocalEnv").callsFake(async () => {});
    await getYmlTemplates();
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sinon.restore();
  });

  const testCases = [
    "transparent-tab",
    "transparent-sso-tab",
    "transparent-bot",
    "transparent-sso-bot",
    "transparent-notification",
    "transparent-tab-bot-func",
    "transparent-m365-tab",
    "transparent-m365-me",
    "beforeV3.4.0-tab",
    "beforeV3.4.0-bot",
    "beforeV3.4.0-tab-bot-func",
    "V3.5.0-V4.0.6-tab",
    "V3.5.0-V4.0.6-tab-bot-func",
    "V3.5.0-V4.0.6-notification-trigger",
    "V3.5.0-V4.0.6-command",
    "V3.5.0-V4.0.6-tab-bot-func-node18",
    "beforeV3.4.0-tab-bot-func-node18",
    "transparent-notification-node18",
    "V4.0.2-notification-trigger",
  ];

  const simpleAuthPath = path.join(os.homedir(), ".fx", "localauth").replace(/\\/g, "\\\\");
  const simpleAuthAppsettingsPath = path.join(
    os.homedir(),
    ".fx",
    "localauth",
    "appsettings.Development.json"
  );

  testCases.forEach((testCase) => {
    it(testCase, async () => {
      const nodeMajorVersion = testCase.endsWith("node18") ? "18" : "16";
      sinon
        .stub(NodeChecker, "getInstalledNodeVersion")
        .resolves({ version: `${nodeMajorVersion}.0.0`, majorVersion: nodeMajorVersion });
      const migrationContext = await mockMigrationContext(projectPath);

      await copyTestProject(path.join("debug", testCase), projectPath);

      await debugMigration(migrationContext);
      const expectedYmlContent = await loadExpectedYmlFile(
        path.join(projectPath, "expected", "app.local.yml")
      );
      const actualYmlContent = await fs.readFile(
        path.join(projectPath, "teamsapp.local.yml"),
        "utf-8"
      );
      assert.equal(
        actualYmlContent,
        expectedYmlContent.replace("SIMPLE_AUTH_APPSETTINGS_PATH", simpleAuthAppsettingsPath)
      );
      assert.equal(
        await fs.readFile(path.join(projectPath, ".vscode", "tasks.json"), "utf-8"),
        (await fs.readFile(path.join(projectPath, "expected", "tasks.json"), "utf-8")).replace(
          "SIMPLE_AUTH_PATH",
          simpleAuthPath
        )
      );
    });
  });
});

describe("updateGitignore", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should update existing gitignore file", async () => {
    const migrationContext: MigrationContext = await mockMigrationContext(projectPath);
    await copyTestProject("happyPath", projectPath);

    await MigratorV3.updateGitignore(migrationContext);

    await assertFileContent(projectPath, ".gitignore", "whenGitignoreExist");
  });

  it("should create new gitignore file when no gitignore file exists", async () => {
    const migrationContext: MigrationContext = await mockMigrationContext(projectPath);
    await copyTestProject("happyPath", projectPath);

    await fs.remove(path.join(projectPath, ".gitignore"));

    await MigratorV3.updateGitignore(migrationContext);

    await assertFileContent(projectPath, ".gitignore", "whenGitignoreNotExist");
  });
});
