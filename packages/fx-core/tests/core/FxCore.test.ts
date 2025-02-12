// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ErrorType,
  ListAPIResult,
  SpecParser,
  SpecParserError,
  ValidationStatus,
  WarningType,
} from "@microsoft/m365-spec-parser";
import {
  DeclarativeCopilotManifestSchema,
  FxError,
  IQTreeNode,
  InputResult,
  Inputs,
  LogProvider,
  Ok,
  Platform,
  Result,
  Stage,
  SystemError,
  TeamsAppManifest,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { assert, expect } from "chai";
import fs from "fs-extra";
import jsyaml from "js-yaml";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import {
  FxCore,
  PackageService,
  getLocalizedString,
  getUuid,
  teamsDevPortalClient,
} from "../../src";
import { ConstantString } from "../../src/common/constants";
import { FeatureFlagName } from "../../src/common/featureFlags";
import { setTools } from "../../src/common/globalVars";
import * as projectHelper from "../../src/common/projectSettingsHelper";
import {
  TeamsfxConfigType,
  TeamsfxVersionState,
  projectTypeChecker,
} from "../../src/common/projectTypeChecker";
import { MetadataV3, VersionSource, VersionState } from "../../src/common/versionMetadata";
import {
  DriverDefinition,
  DriverInstance,
  ExecutionResult,
  ILifecycle,
  LifecycleName,
  Output,
  ProjectModel,
  UnresolvedPlaceholders,
} from "../../src/component/configManager/interface";
import { YamlParser } from "../../src/component/configManager/parser";
import { coordinator } from "../../src/component/coordinator";
import { UpdateAadAppDriver } from "../../src/component/driver/aad/update";
import * as buildAadManifest from "../../src/component/driver/aad/utility/buildAadManifest";
import { AddWebPartDriver } from "../../src/component/driver/add/addWebPart";
import { DriverContext } from "../../src/component/driver/interface/commonArgs";
import { CreateAppPackageDriver } from "../../src/component/driver/teamsApp/createAppPackage";
import { AppStudioError } from "../../src/component/driver/teamsApp/errors";
import { SyncManifestArgs } from "../../src/component/driver/teamsApp/interfaces/SyncManifest";
import { SyncManifestDriver } from "../../src/component/driver/teamsApp/syncManifest";
import { teamsappMgr } from "../../src/component/driver/teamsApp/teamsappMgr";
import { copilotGptManifestUtils } from "../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { ValidateManifestDriver } from "../../src/component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../../src/component/driver/teamsApp/validateAppPackage";
import { ValidateWithTestCasesDriver } from "../../src/component/driver/teamsApp/validateTestCases";
import { createDriverContext } from "../../src/component/driver/util/utils";
import { WrapDriverContext } from "../../src/component/driver/util/wrapUtil";
import "../../src/component/feature/sso";
import * as CopilotPluginHelper from "../../src/component/generator/apiSpec/helper";
import * as pluginGeneratorHelper from "../../src/component/generator/apiSpec/helper";
import * as copilotExtensionHelper from "../../src/component/generator/copilotExtension//helper";
import { TemplateNames } from "../../src/component/generator/templates/templateNames";
import { LaunchHelper } from "../../src/component/m365/launchHelper";
import { envUtil } from "../../src/component/utils/envUtil";
import { metadataUtil } from "../../src/component/utils/metadataUtil";
import { pathUtils } from "../../src/component/utils/pathUtils";
import * as collaborator from "../../src/core/collaborator";
import { environmentManager } from "../../src/core/environment";
import * as projectMigratorV3 from "../../src/core/middleware/projectMigratorV3";
import * as projMigrator from "../../src/core/middleware/projectMigratorV3";
import * as migrationUtil from "../../src/core/middleware/utils/v3MigrationUtils";
import { CoreHookContext } from "../../src/core/types";
import {
  FileNotFoundError,
  InputValidationError,
  InvalidProjectError,
  MissingEnvironmentVariablesError,
  MissingRequiredInputError,
  NotImplementedError,
  UserCancelError,
} from "../../src/error/common";
import { NoNeedUpgradeError } from "../../src/error/upgrade";
import {
  CapabilityOptions,
  QuestionNames,
  ScratchOptions,
  SyncManifestInputs,
  UninstallInputs,
  questionNodes,
} from "../../src/question";
import { ApiPluginStartOptions, HubOptions } from "../../src/question/constants";
import { ProjectTypeOptions } from "../../src/question/scaffold/vsc/ProjectTypeOptions";
import { validationUtils } from "../../src/ui/validationUtils";
import { MockTools, randomAppName } from "./utils";

const tools = new MockTools();

describe("Core basic APIs", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  beforeEach(() => {
    setTools(tools);
  });
  afterEach(async () => {
    sandbox.restore();
  });

  it("deploy aad manifest happy path with param", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    // sandbox.stub(UpdateAadAppDriver.prototype, "run").resolves(new Ok(new Map()));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      [QuestionNames.TargetEnvName]: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sandbox.spy(UpdateAadAppDriver.prototype, "execute");
    await core.deployAadManifest(inputs);
    sandbox.assert.calledOnce(runSpy);
    assert.isNotNull(runSpy.getCall(0).args[0]);
    assert.strictEqual(
      runSpy.getCall(0).args[0].manifestPath,
      path.join(os.tmpdir(), appName, "aad.manifest.json")
    );
    runSpy.restore();
  });

  it("add web part to SPFx", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    const runSpy = sandbox.spy(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("add web part to SPFx - CLI help", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.CLI_HELP,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    const runSpy = sandbox.spy(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("add web part to SPFx with empty .yo-rc.json", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    sandbox.stub(fs, "pathExists").callsFake(async (directory: string) => {
      if (directory.includes(path.join("webparts", "helloworld"))) {
        return false;
      }
      return true;
    });
    sandbox.stub(fs, "readJson").resolves({});
    const runSpy = sandbox.stub(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("add web part to SPFx with framework", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appPath = path.join(os.tmpdir(), appName);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      "spfx-folder": ".\\src",
      "manifest-path": path.join(appPath, "appPackage\\manifest.json"),
      "local-manifest-path": path.join(appPath, "appPackage\\manifest.local.json"),
      "spfx-webpart-name": "helloworld",
      "spfx-install-latest-package": "true",
      "spfx-load-package-version": "loaded",
      stage: Stage.addWebpart,
      projectPath: appPath,
    };

    sandbox.stub(fs, "pathExists").callsFake(async (directory: string) => {
      if (directory.includes(path.join("webparts", "helloworld"))) {
        return false;
      }
      return true;
    });
    sandbox.stub(fs, "readJson").resolves({
      "@microsoft/generator-sharepoint": {
        template: "react",
      },
    });
    const runSpy = sandbox.stub(AddWebPartDriver.prototype, "run");
    await core.addWebpart(inputs);
    sandbox.assert.calledOnce(runSpy);
    runSpy.restore();
  });

  it("deploy aad manifest happy path", async () => {
    const promtionOnVSC =
      'Your Microsoft Entra app has been deployed successfully. To view that, click "More info"';

    const core = new FxCore(tools);
    const showMessage = sandbox.spy(tools.ui, "showMessage") as unknown as sinon.SinonSpy<
      ["info" | "warn" | "error", string, boolean, ...string[]],
      Promise<Result<string | undefined, FxError>>
    >;
    const openUrl = sandbox.spy(tools.ui, "openUrl");
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "build")));
    await deleteTestProject(appName);
    assert.isTrue(res.isOk());
    assert.isTrue(showMessage.called);
    assert.equal(showMessage.getCall(0).args[0], "info");
    assert.equal(showMessage.getCall(0).args[1], promtionOnVSC);
    assert.isFalse(showMessage.getCall(0).args[2]);
    assert.equal(showMessage.getCall(0).args[3], "More info");
    assert.isFalse(openUrl.called);
  });
  it("deploy aad manifest happy path with click more info", async () => {
    const core = new FxCore(tools);
    sandbox.stub(tools.ui, "showMessage").resolves(ok("More info"));
    sandbox.stub(tools.ui, "openUrl").resolves(ok(true));
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isOk());
    if (res.isErr()) console.error(res.error);
    assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "build")));
    await deleteTestProject(appName);
  });

  it("deploy aad manifest happy path without click learn more", async () => {
    const core = new FxCore(tools);
    sandbox.stub(tools.ui, "showMessage").resolves(err(new UserError("test", "test", "test")));
    sandbox.stub(tools.ui, "openUrl").resolves(ok(true));
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isOk());
    if (res.isErr()) console.error(res.error);
    assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "build")));
    await deleteTestProject(appName);
  });
  it("deploy aad manifest happy path on cli", async () => {
    const core = new FxCore(tools);
    const showMessage = sandbox.spy(tools.ui, "showMessage") as unknown as sinon.SinonSpy<
      ["info" | "warn" | "error", string, boolean, ...string[]],
      Promise<Result<string | undefined, FxError>>
    >;
    const appName = await mockV3Project();
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .resolves({ result: new Ok(new Map()), summaries: [] });
    const inputs: Inputs = {
      platform: Platform.CLI,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    await deleteTestProject(appName);
    assert.isTrue(showMessage.calledOnce);
    assert.equal(showMessage.getCall(0).args[0], "info");
    assert.equal(
      showMessage.getCall(0).args[1],
      "Your Microsoft Entra app has been updated successfully."
    );
    assert.isFalse(showMessage.getCall(0).args[2]);
    assert.isTrue(res.isOk());
  });

  it("deploy aad manifest return err", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appManifestPath = path.join(os.tmpdir(), appName, "aad.manifest.json");
    sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok(["dev", "local"]));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: appManifestPath,
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    sandbox
      .stub(UpdateAadAppDriver.prototype, "execute")
      .throws(new UserError("error name", "fake_error", "fake_err_msg"));
    const errMsg = `AAD manifest doesn't exist in ${appManifestPath}, please use the CLI to specify an AAD manifest to deploy.`;
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.strictEqual(res.error.message, "fake_err_msg");
    }
  });

  it("deploy aad manifest with missing env err", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appManifestPath = path.join(os.tmpdir(), appName, "aad.manifest.json");
    sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok([""]));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: appManifestPath,
      env: undefined,
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    sandbox.stub(UpdateAadAppDriver.prototype, "execute").resolves({
      result: err(
        new MissingEnvironmentVariablesError(
          "aadApp/update",
          "AAD_APP_OBJECT_ID",
          "fake path",
          "https://fake-help-link"
        )
      ),
      summaries: [],
    });
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      // Cannot assert the full message because the mocked code can't get correct env file path
      assert.include(
        res.error.message,
        "Missing environment variables 'AAD_APP_OBJECT_ID' for file: fake path. Please edit the .env file"
      );

      assert.include(
        res.error.message,
        "For new Teams Toolkit projects, make sure you've run provision or debug to set these variables correctly."
      );
    }
  });

  it("deploy aad manifest not exist", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const appManifestPath = path.join(os.tmpdir(), appName, "aad.manifest.json");
    await fs.remove(appManifestPath);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: appName,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Capabilities]: ["Tab", "TabSSO"],
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AadAppManifestFilePath]: path.join(os.tmpdir(), appName, "aad.manifest.json"),
      env: "dev",
      stage: Stage.deployAad,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.deployAadManifest(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof FileNotFoundError);
    }
    await deleteTestProject(appName);
  });

  it("phantomMigrationV3 happy path", async () => {
    const core = new FxCore(tools);
    const appName = await mockV2Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
      skipUserConfirm: true,
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isOk());
    await deleteTestProject(appName);
  });

  it("phantomMigrationV3 return error for invalid V2 project", async () => {
    sandbox.stub(projectMigratorV3, "checkActiveResourcePlugins").resolves(false);

    const core = new FxCore(tools);
    const appName = await mockV2Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
      skipUserConfirm: true,
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(
      res._unsafeUnwrapErr().message.includes(new InvalidProjectError(inputs.projectPath!).message)
    );
    await deleteTestProject(appName);
  });

  it("phantomMigrationV3 return error for non-project", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir()),
      skipUserConfirm: true,
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(
      res._unsafeUnwrapErr().message.includes(new InvalidProjectError(inputs.projectPath!).message)
    );
  });

  it("phantomMigrationV3 return error for V5 project", async () => {
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const res = await core.phantomMigrationV3(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(res._unsafeUnwrapErr().message.includes(new NoNeedUpgradeError().message));
    await deleteTestProject(appName);
  });

  it("permission v3", async () => {
    let res;
    const core = new FxCore(tools);
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    sandbox.stub(questionNodes, "grantPermission").returns({ data: { type: "group" } });
    sandbox.stub(questionNodes, "listCollaborator").returns({ data: { type: "group" } });
    sandbox.stub(collaborator, "listCollaborator").resolves(ok(undefined as any));
    sandbox.stub(collaborator, "checkPermission").resolves(ok(undefined as any));
    sandbox.stub(collaborator, "grantPermission").resolves(ok(undefined as any));

    res = await core.listCollaborator(inputs);
    assert.isTrue(res.isOk());
    res = await core.checkPermission(inputs);
    assert.isTrue(res.isOk());
    res = await core.grantPermission(inputs);
    assert.isTrue(res.isOk());
  });

  it("buildAadManifest method should exist", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failure that when local template not released to GitHub
      NODE_ENV: "development", // workaround test failure that when local template not released to GitHub
      AAD_APP_OBJECT_ID: getUuid(),
      AAD_APP_CLIENT_ID: getUuid(),
      TAB_DOMAIN: "fake",
      TAB_ENDPOINT: "fake",
    });

    const originFunc = envUtil.readEnv;
    try {
      envUtil.readEnv = async () => {
        return ok({
          AAD_APP_OBJECT_ID: getUuid(),
          AAD_APP_CLIENT_ID: getUuid(),
          TAB_DOMAIN: "fake",
          TAB_ENDPOINT: "fake",
        });
      };
      const core = new FxCore(tools);
      const appName = await mockV3Project();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: path.join(os.tmpdir(), appName),
      };
      sandbox.stub(buildAadManifest, "buildAadManifest").resolves({} as any);
      const result = await core.buildAadManifest(inputs);
      assert.isTrue(result.isOk());
    } finally {
      envUtil.readEnv = originFunc;
      restore();
    }
  });

  it("convertAadToNewSchema method should work fine", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failure that when local template not released to GitHub
      NODE_ENV: "development", // workaround test failure that when local template not released to GitHub
    });

    try {
      const core = new FxCore(tools);
      const appName = await mockV3Project();
      const projectPath = path.join(os.tmpdir(), appName);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: projectPath,
        [QuestionNames.AadAppManifestFilePath]: `${projectPath}/aad.manifest.json`,
      };

      sandbox.stub(tools.ui, "showMessage").resolves(ok("Continue"));
      const result = await core.convertAadToNewSchema(inputs);
      assert.isTrue(result.isOk());
    } finally {
      restore();
    }
  });

  it("convertAadToNewSchema should throw file not exist error if aad.manifest.json does not exist", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failure that when local template not released to GitHub
      NODE_ENV: "development", // workaround test failure that when local template not released to GitHub
    });

    try {
      const core = new FxCore(tools);
      const appName = await mockV3Project();
      const projectPath = path.join(os.tmpdir(), appName);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: projectPath,
        [QuestionNames.AadAppManifestFilePath]: `/not-exist-path/aad.manifest.json`,
      };

      const result = await core.convertAadToNewSchema(inputs);
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.isTrue(result.error instanceof FileNotFoundError);
      }
    } finally {
      restore();
    }
  });

  it("convertAadToNewSchema throw user cancel error if not confirmed", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failure that when local template not released to GitHub
      NODE_ENV: "development", // workaround test failure that when local template not released to GitHub
    });

    try {
      const core = new FxCore(tools);
      const appName = await mockV3Project();
      const projectPath = path.join(os.tmpdir(), appName);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: projectPath,
        [QuestionNames.AadAppManifestFilePath]: `${projectPath}/aad.manifest.json`,
      };

      const result = await core.convertAadToNewSchema(inputs);
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.isTrue(result.error instanceof UserCancelError);
      }
    } finally {
      restore();
    }
  });

  it("convertAadToNewSchema show message when manifest is in new schema", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failure that when local template not released to GitHub
      NODE_ENV: "development", // workaround test failure that when local template not released to GitHub
    });

    try {
      const core = new FxCore(tools);
      const appName = await mockV3Project();
      const projectPath = path.join(os.tmpdir(), appName);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: projectPath,
        [QuestionNames.AadAppManifestFilePath]: `${projectPath}/aad.manifest.json`,
      };

      sandbox.stub(fs, "readJson").resolves({ displayName: "displayName" });
      const showMessageStub = sandbox.stub(tools.ui, "showMessage");

      const result = await core.convertAadToNewSchema(inputs);
      sandbox.assert.calledOnceWithExactly(
        showMessageStub,
        "info",
        getLocalizedString("core.convertAadToNewSchema.alreadyNewSchema") as any,
        false
      );
      assert.isTrue(result.isOk());
    } finally {
      restore();
    }
  });

  it("addSso method should exist", async () => {
    const restore = mockedEnv({
      TEAMSFX_DEBUG_TEMPLATE: "true", // workaround test failures when template changed but not release to GitHub alpha template
      NODE_ENV: "development", // workaround test failures when template changed but not release to GitHub alpha template
    });
    try {
      const appName = randomAppName();
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.AppName]: appName,
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.ProjectType]: ProjectTypeOptions.tabOptionId,
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
        [QuestionNames.Folder]: os.tmpdir(),
        [QuestionNames.TemplateName]: TemplateNames.Tab,
        stage: Stage.create,
        projectPath: path.join(os.tmpdir(), appName, "samples-v3"),
      };
      const res = await core.createProject(inputs);
      const projectPath = inputs.projectPath!;
      assert.isTrue(res.isOk() && res.value.projectPath === projectPath);

      const implement = new FxCore(tools);

      const mockFunc = {
        namespace: "mock namespace",
        method: "addSso",
      };

      const result = await implement.executeUserTask(mockFunc, inputs);
      assert.isTrue(result.isOk());
    } finally {
      restore();
    }
  });
  it("uninstall with empty input", async () => {
    const core = new FxCore(tools);
    const inputs: UninstallInputs = {
      platform: Platform.CLI,
    };
    const res = await core.uninstall(inputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall with invalid mode", async () => {
    const core = new FxCore(tools);
    const inputs = {
      platform: Platform.CLI,
      mode: "invalid",
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall by manifest ID - success", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(teamsDevPortalClient, "deleteApp").resolves(true);
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").resolves();
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: [
        "m365-app",
        "app-registration",
        "bot-framework-registration",
      ],
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isOk());
  });
  it("uninstall by manifest ID - missing manifest ID", async () => {
    const core = new FxCore(tools);
    const inputs: UninstallInputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall by manifest ID - empty options", async () => {
    const core = new FxCore(tools);
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isOk());
  });
  it("uninstall by manifest ID - failed to get token", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError("mockedSource", "mockedError", "mockedMessage")));
    const inputs1 = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["m365-app"],
      nonInteractive: true,
    };
    const res1 = await core.uninstall(inputs1 as UninstallInputs);
    assert.isTrue(res1.isErr());

    const inputs2 = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["app-registration"],
      nonInteractive: true,
    };
    const res2 = await core.uninstall(inputs2 as UninstallInputs);
    assert.isTrue(res2.isErr());

    const inputs3 = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["bot-framework-registration"],
      nonInteractive: true,
    };
    const res3 = await core.uninstall(inputs3 as UninstallInputs);
    assert.isTrue(res3.isErr());
  });
  it("uninstall by manifest ID - failed to get title ID", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(PackageService.prototype, "retrieveTitleId").throws("error");
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: [
        "m365-app",
        "app-registration",
        "bot-framework-registration",
      ],
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall by manifest ID - failed to get bot ID", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves(undefined);
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["bot-framework-registration"],
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall by manifest ID - M365 App user cancel", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(tools.ui, "confirm").resolves(ok({ result: false } as InputResult<boolean>));
    sandbox.stub(teamsDevPortalClient, "deleteApp").throws("error");
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").throws("error");
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["m365-app"],
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof UserCancelError);
    }
  });
  it("uninstall by manifest ID - TDP user cancel", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(tools.ui, "confirm").resolves(ok({ result: false } as InputResult<boolean>));
    sandbox.stub(teamsDevPortalClient, "deleteApp").throws("error");
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").throws("error");
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["app-registration"],
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof UserCancelError);
    }
  });
  it("uninstall by manifest ID - Bot user cancel", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(tools.ui, "confirm").resolves(ok({ result: false } as InputResult<boolean>));
    sandbox.stub(teamsDevPortalClient, "deleteApp").throws("error");
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").throws("error");
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeManifestId,
      [QuestionNames.ManifestId]: "valid-manifest-id",
      [QuestionNames.UninstallOptions]: ["bot-framework-registration"],
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof UserCancelError);
    }
  });
  it("uninstall by env - success", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(teamsDevPortalClient, "deleteApp").resolves(true);
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").resolves();
    const appName = await mockCliUninstallProject();
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      env: "dev",
      [QuestionNames.UninstallOptions]: [
        "m365-app",
        "app-registration",
        "bot-framework-registration",
      ],
      nonInteractive: true,
    };

    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isOk());

    const envRes = await envUtil.readEnv(path.join(os.tmpdir(), appName), "dev", false);
    assert.isTrue(envRes.isOk());
    if (envRes.isOk()) {
      const envVars = envRes.value;
      assert.isTrue(envVars["TEAMS_APP_ID"] === "");
      assert.isTrue(envVars["M365_TITLE_ID"] === "");
      assert.isTrue(envVars["BOT_ID"] === "");
    }
    await deleteTestProject(appName);
  });
  it("uninstall by env - missing env", async () => {
    const core = new FxCore(tools);
    const appName = await mockCliUninstallProject();

    const inputs: UninstallInputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      nonInteractive: true,
    };

    const res = await core.uninstall(inputs);
    assert.isTrue(res.isErr());
    await deleteTestProject(appName);
  });
  it("uninstall by env - empty options", async () => {
    const core = new FxCore(tools);
    const appName = await mockCliUninstallProject();

    const inputs: UninstallInputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      nonInteractive: true,
      env: "dev",
    };

    const res = await core.uninstall(inputs);
    assert.isTrue(res.isOk());
    await deleteTestProject(appName);
  });
  it("uninstall by env - invalid yaml", async () => {
    const core = new FxCore(tools);
    const appName = await mockCliUninstallProject();
    sandbox.stub(metadataUtil, "parse").resolves(err(new SystemError("", "", "")));
    const inputs: UninstallInputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      nonInteractive: true,
      env: "dev",
    };
    const res = await core.uninstall(inputs);
    assert.isTrue(res.isErr());
    await deleteTestProject(appName);
  });
  it("uninstall by env - empty provision actions", async () => {
    const core = new FxCore(tools);
    const appName = await mockCliUninstallProject();
    sandbox.stub(metadataUtil, "parse").resolves(ok({} as ProjectModel));
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError("mockedSource", "mockedError", "mockedMessage")));
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      nonInteractive: true,
      env: "dev",
      [QuestionNames.UninstallOptions]: [
        "m365-app",
        "app-registration",
        "bot-framework-registration",
      ],
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isOk());
    await deleteTestProject(appName);
  });
  it("uninstall by env - empty env key name", async () => {
    const core = new FxCore(tools);
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        provision: {
          name: "provision",
          driverDefs: [
            {
              uses: "teamsApp/create",
            },
            {
              uses: "botFramework/create",
            },
            {
              uses: "teamsApp/extendToM365",
            },
          ],
        },
      } as ProjectModel)
    );
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(teamsDevPortalClient, "deleteApp").resolves(true);
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").resolves();
    const appName = await mockCliUninstallProject();
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      env: "dev",
      [QuestionNames.UninstallOptions]: [
        "m365-app",
        "app-registration",
        "bot-framework-registration",
      ],
      nonInteractive: true,
    };

    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isOk());

    const envRes = await envUtil.readEnv(path.join(os.tmpdir(), appName), "dev", false);
    assert.isTrue(envRes.isOk());
    if (envRes.isOk()) {
      const envVars = envRes.value;
      assert.isTrue(envVars["TEAMS_APP_ID"] === "");
      assert.isTrue(envVars["M365_TITLE_ID"] === "");
      assert.isTrue(envVars["BOT_ID"] === "");
    }
    await deleteTestProject(appName);
  });
  it("uninstall by env - failed to get token", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError("mockedSource", "mockedError", "mockedMessage")));
    sandbox.stub(teamsDevPortalClient, "deleteApp").resolves(true);
    sandbox.stub(teamsDevPortalClient, "getBotId").resolves("mocked-bot-id");
    sandbox.stub(teamsDevPortalClient, "deleteBot").resolves();
    sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("mocked-title-id");
    sandbox.stub(PackageService.prototype, "unacquire").resolves();
    const appName = await mockCliUninstallProject();
    const inputs1 = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      env: "dev",
      [QuestionNames.UninstallOptions]: ["m365-app"],
      nonInteractive: true,
    };

    const res1 = await core.uninstall(inputs1 as UninstallInputs);
    assert.isTrue(res1.isErr());

    const inputs2 = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      env: "dev",
      [QuestionNames.UninstallOptions]: ["app-registration"],
      nonInteractive: true,
    };

    const res2 = await core.uninstall(inputs2 as UninstallInputs);
    assert.isTrue(res2.isErr());

    const inputs3 = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeEnv,
      projectPath: path.join(os.tmpdir(), appName),
      env: "dev",
      [QuestionNames.UninstallOptions]: ["bot-framework-registration"],
      nonInteractive: true,
    };

    const res3 = await core.uninstall(inputs3 as UninstallInputs);
    assert.isTrue(res3.isErr());
  });
  it("uninstall by title ID - success", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(PackageService.prototype, "unacquire").resolves();
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeTitleId,
      [QuestionNames.TitleId]: "mocked-title-id",
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isOk());
  });
  it("uninstall by title ID - missing title ID", async () => {
    const core = new FxCore(tools);
    sandbox
      .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mocked-token"));
    sandbox.stub(PackageService.prototype, "unacquire").resolves();
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeTitleId,
      nonInteractive: true,
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall by title ID - failed", async () => {
    const core = new FxCore(tools);
    sandbox.stub(core, "uninstallM365App").resolves(err(new SystemError("", "", "")));
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.UninstallMode]: QuestionNames.UninstallModeTitleId,
      nonInteractive: true,
      [QuestionNames.TitleId]: "mocked-title-id",
    };
    const res = await core.uninstall(inputs as UninstallInputs);
    assert.isTrue(res.isErr());
  });
  it("uninstall M365 App - invalid input", async () => {
    const core = new FxCore(tools);
    const res = await core.uninstallM365App(undefined, undefined);
    assert.isTrue(res.isErr());
  });
  it("uninstall Bot Framework Registration - invalid input", async () => {
    const core = new FxCore(tools);
    const res = await core.uninstallBotFrameworRegistration(undefined, undefined);
    assert.isTrue(res.isErr());
  });
  it("reset env var - happy path", async () => {
    const core = new FxCore(tools);
    const ctx: CoreHookContext = { arguments: [], envVars: { testKey: "oldValue" } };
    core.resetEnvVar("testKey", ctx);
    expect(ctx.envVars).to.deep.equal({ testKey: "" });
  });
  it("reset env var - undefine ctx", async () => {
    const core = new FxCore(tools);
    const ctx: CoreHookContext | undefined = undefined;
    core.resetEnvVar("testKey", ctx);
    assert.isUndefined(ctx);
  });
  it("reset env var - initialize envVars if it is undefined", async () => {
    const core = new FxCore(tools);
    const ctx: CoreHookContext = { arguments: [], envVars: undefined };
    core.resetEnvVar("testKey", ctx, false);
    expect(ctx.envVars).to.deep.equal({ testKey: "" });
  });
  it("reset env var - skipIfNotExist is true", async () => {
    const core = new FxCore(tools);
    const ctx: CoreHookContext = { arguments: [], envVars: { existingKey: "value" } };
    core.resetEnvVar("testKey", ctx);
    expect(ctx.envVars).to.deep.equal({ existingKey: "value" });
  });
  it("reset env var - skipIfNotExist is false", async () => {
    const core = new FxCore(tools);
    const ctx: CoreHookContext = { arguments: [], envVars: { existingKey: "value" } };
    core.resetEnvVar("testKey", ctx, false);
    expect(ctx.envVars).to.deep.equal({ existingKey: "value", testKey: "" });
  });
});

describe("apply yaml template", async () => {
  const tools = new MockTools();
  beforeEach(() => {
    setTools(tools);
  });
  describe("when run with missing input", async () => {
    it("should return error when projectPath is undefined", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: undefined,
      };
      const res = await core.apply(inputs, "", "provision");
      assert.isTrue(res.isErr() && res.error instanceof InputValidationError);
    });

    it("should return error when env is undefined", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: undefined,
      };
      const res = await core.apply(inputs, "", "provision");
      assert.isTrue(res.isErr() && res.error instanceof InputValidationError);
    });
  });

  describe("when readEnv returns error", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(err(mockedError));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error too", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });

  describe("when YamlParser returns error", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(err(mockedError));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error too", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });

  describe("when running against an empty yaml file", async () => {
    const sandbox = sinon.createSandbox();

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(ok({ version: "1.0.0" }));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return ok", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isOk());
    });
  });

  describe("when lifecycle returns error", async () => {
    const sandbox = sinon.createSandbox();
    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    class MockedProvision implements ILifecycle {
      name: LifecycleName = "provision";
      driverDefs: DriverDefinition[] = [];
      public async run(ctx: DriverContext): Promise<Result<Output, FxError>> {
        return err(mockedError);
      }

      public resolvePlaceholders(): UnresolvedPlaceholders {
        return [];
      }

      public async execute(ctx: DriverContext): Promise<ExecutionResult> {
        return {
          result: err({
            kind: "Failure",
            error: mockedError,
          }),
          summaries: [],
        };
      }

      public resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
        return ok([]);
      }
    }

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(
        ok({
          version: "1.0.0",
          provision: new MockedProvision(),
        })
      );
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });
  describe("runLifecycle", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");
    class MockedProvision implements ILifecycle {
      name: LifecycleName = "provision";
      driverDefs: DriverDefinition[] = [];
      public async run(ctx: DriverContext): Promise<Result<Output, FxError>> {
        return err(mockedError);
      }

      public resolvePlaceholders(): UnresolvedPlaceholders {
        return [];
      }

      public async execute(ctx: DriverContext): Promise<ExecutionResult> {
        return {
          result: ok(new Map()),
          summaries: [],
        };
      }

      public resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
        return ok([]);
      }
    }

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      const context = createDriverContext(inputs);
      const lifecycle = new MockedProvision();
      const res = await core.runLifecycle(lifecycle, context, "dev");
      assert.isTrue(res.isOk());
    });

    it("partial success", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      const lifecycle = new MockedProvision();
      sandbox.stub(lifecycle, "execute").resolves({
        result: err({
          kind: "PartialSuccess",
          env: new Map(),
          reason: {
            kind: "UnresolvedPlaceholders",
            failedDriver: { uses: "t", with: {} },
            unresolvedPlaceHolders: ["TEST_VAR"],
          },
        }),
        summaries: [],
      });
      const context = createDriverContext(inputs);
      const res = await core.runLifecycle(lifecycle, context, "dev");
      assert.isTrue(res.isOk());
    });

    it("DriverError", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      const lifecycle = new MockedProvision();
      sandbox.stub(lifecycle, "execute").resolves({
        result: err({
          kind: "PartialSuccess",
          env: new Map(),
          reason: {
            kind: "DriverError",
            failedDriver: { uses: "t", with: {} },
            error: mockedError,
          },
        }),
        summaries: [],
      });
      const context = createDriverContext(inputs);
      const res = await core.runLifecycle(lifecycle, context, "dev");
      assert.isTrue(res.isErr());
    });
  });
});

async function mockV3Project(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  // await fs.move(path.join(__dirname, "../sampleV3"), path.join(os.tmpdir(), appName));
  await fs.copy(path.join(__dirname, "../samples/sampleV3/"), path.join(projectPath));
  return appName;
}

async function mockV2Project(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  await fs.copy(
    path.join(__dirname, "../core/middleware/testAssets/v3Migration/happyPath"),
    path.join(projectPath)
  );
  return appName;
}

async function mockCliUninstallProject(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  await fs.copy(path.join(__dirname, "../samples/uninstall/"), path.join(projectPath));
  return appName;
}

async function deleteTestProject(appName: string) {
  await fs.remove(path.join(os.tmpdir(), appName));
}

describe("createEnvCopyV3", async () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  const sourceEnvContent = [
    "# this is a comment",
    "TEAMSFX_ENV=dev",
    "APP_NAME_SUFFIX=dev",
    "",
    "_KEY1=value1",
    "KEY2=value2",
    "SECRET_KEY3=xxxx",
  ];
  const sourceEnvStr = sourceEnvContent.join(os.EOL);

  const writeStreamContent: string[] = [];
  // fs.WriteStream's full interface is too hard to mock. We only use write() and end() so we just mock them here.
  class MockedWriteStream {
    write(chunk: any, callback?: ((error: Error | null | undefined) => void) | undefined): boolean {
      writeStreamContent.push(chunk);
      return true;
    }
    end(): boolean {
      return true;
    }
  }

  beforeEach(() => {
    sandbox.stub(fs, "readFile").resolves(Buffer.from(sourceEnvStr, "utf8"));
    sandbox.stub<any, any>(fs, "createWriteStream").returns(new MockedWriteStream());
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should create new .env file with desired content", async () => {
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("./env/.env.dev"));
    sandbox.stub(fs, "pathExists").resolves(true);
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isOk());
    assert(
      writeStreamContent[0] === `${sourceEnvContent[0]}${os.EOL}`,
      "comments should be copied"
    );
    assert(
      writeStreamContent[1] === `TEAMSFX_ENV=newEnv${os.EOL}`,
      "TEAMSFX_ENV's value should be new env name"
    );
    assert(
      writeStreamContent[2] === `APP_NAME_SUFFIX=newEnv${os.EOL}`,
      "APP_NAME_SUFFIX's value should be new env name"
    );
    assert(writeStreamContent[3] === `${os.EOL}`, "empty line should be coped");
    assert(
      writeStreamContent[4] === `_KEY1=${os.EOL}`,
      "key starts with _ should be copied with empty value"
    );
    assert(
      writeStreamContent[5] === `KEY2=${os.EOL}`,
      "key not starts with _ should be copied with empty value"
    );
    assert(
      writeStreamContent[6] === `SECRET_KEY3=${os.EOL}`,
      "key not starts with SECRET_ should be copied with empty value"
    );
  });

  it("should failed case 1", async () => {
    sandbox
      .stub(pathUtils, "getEnvFilePath")
      .onFirstCall()
      .resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isErr());
  });

  it("should failed case 2", async () => {
    sandbox
      .stub(pathUtils, "getEnvFilePath")
      .onFirstCall()
      .resolves(ok("./env"))
      .onSecondCall()
      .resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isErr());
  });
});

describe("publishInDeveloperPortal", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();

  before(() => {
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
  });
  afterEach(() => {
    sandbox.restore();
  });

  it("success", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      env: "local",
      projectPath: "project-path",
      platform: Platform.VSCode,
      [QuestionNames.AppPackagePath]: "path",
      ignoreLockByUT: true,
    };
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(coordinator, "publishInDeveloperPortal").resolves(ok(undefined));
    const res = await core.publishInDeveloperPortal(inputs);

    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
});

describe("Teams app APIs", async () => {
  const tools = new MockTools();
  const core = new FxCore(tools);

  afterEach(() => {
    sinon.restore();
  });

  it("validate app package", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppPackageFilePath]: ".\\build\\appPackage\\appPackage.dev.zip",
      [QuestionNames.ValidateMethod]: "validateAgainstAppPackage",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sinon.spy(ValidateAppPackageDriver.prototype, "execute");
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    await core.validateApplication(inputs);
    sinon.assert.calledOnce(runSpy);
  });

  it("validate manifest", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: ".\\appPackage\\manifest.json",
      [QuestionNames.ValidateMethod]: "validateAgainstSchema",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sinon.spy(ValidateManifestDriver.prototype, "execute");
    await core.validateApplication(inputs);
    sinon.assert.calledOnce(runSpy);
  });

  it("validate with test cases", async () => {
    const appName = await mockV3Project();

    const mockedEnvRestore = mockedEnv({
      [FeatureFlagName.AsyncAppValidation]: "true",
    });

    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppPackageFilePath]: ".\\build\\appPackage\\appPackage.dev.zip",
      [QuestionNames.ValidateMethod]: "validateWithTestCases",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const runSpy = sinon.spy(ValidateWithTestCasesDriver.prototype, "execute");
    await core.validateApplication(inputs);
    sinon.assert.calledOnce(runSpy);

    mockedEnvRestore();
  });

  it("create app package", async () => {
    setTools(tools);
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: ".\\appPackage\\manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
      [QuestionNames.OutputZipPathParamName]: ".\\build\\appPackage\\appPackage.dev.zip",
    };

    sinon.stub(process, "platform").value("win32");
    const runStub = sinon
      .stub(CreateAppPackageDriver.prototype, "execute")
      .resolves({ result: ok(new Map()), summaries: [] });
    const showMessageStub = sinon.stub(tools.ui, "showMessage");
    await core.createAppPackage(inputs);
    sinon.assert.calledOnce(runStub);
    sinon.assert.calledOnce(showMessageStub);
  });

  it("publish application", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      projectPath: path.join(os.tmpdir(), appName),
    };

    sinon
      .stub(coordinator, "publish")
      .resolves(err(new SystemError("mockedSource", "mockedError", "mockedMessage")));
    await core.publishApplication(inputs);
  });
});

describe("previewWithManifest", () => {
  const tools = new MockTools();
  const core = new FxCore(tools);

  afterEach(() => {
    sinon.restore();
  });

  it("getManifestV3 error", async () => {
    sinon.stub(manifestUtils, "getManifestV3").resolves(err({ foo: "bar" } as any));
    const appName = await mockV3Project();
    const inputs: Inputs = {
      [QuestionNames.M365Host]: HubOptions.teams().id,
      [QuestionNames.TeamsAppManifestFilePath]: path.join(
        os.tmpdir(),
        appName,
        "appPackage",
        "manifest.template.json"
      ),
      env: "dev",
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const result = await core.previewWithManifest(inputs);
    assert.isTrue(result.isErr());
    assert.deepEqual((result as any).error, { foo: "bar" });
  });

  it("getLaunchUrl error", async () => {
    const appName = await mockV3Project();
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(new TeamsAppManifest()));
    sinon.stub(LaunchHelper.prototype, "getLaunchUrl").resolves(err({ foo: "bar" } as any));
    const inputs: Inputs = {
      [QuestionNames.M365Host]: HubOptions.teams().id,
      [QuestionNames.TeamsAppManifestFilePath]: path.join(
        os.tmpdir(),
        appName,
        "appPackage",
        "manifest.template.json"
      ),
      env: "dev",
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const result = await core.previewWithManifest(inputs);
    assert.isTrue(result.isErr());
    assert.deepEqual((result as any).error, { foo: "bar" });
  });

  it("happy path", async () => {
    const appName = await mockV3Project();
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(new TeamsAppManifest()));
    sinon.stub(LaunchHelper.prototype, "getLaunchUrl").resolves(ok("test-url"));
    const inputs: Inputs = {
      [QuestionNames.M365Host]: HubOptions.teams().id,
      [QuestionNames.TeamsAppManifestFilePath]: path.join(
        os.tmpdir(),
        appName,
        "appPackage",
        "manifest.template.json"
      ),
      env: "dev",
      platform: Platform.VSCode,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const result = await core.previewWithManifest(inputs);
    assert.isTrue(result.isOk());
    assert.deepEqual((result as any).value, "test-url");
  });
});

describe("getProjectId", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    const core = new FxCore(tools);
    sandbox.stub(core, "getProjectMetadata").resolves(
      ok({
        projectId: "12345",
        version: "1.1.1",
      })
    );
    const res = await core.getProjectId(".");
    assert.isTrue(res.isOk() && res.value === "12345");
  });
  it("return empty value", async () => {
    const core = new FxCore(tools);
    sandbox.stub(core, "getProjectMetadata").resolves(ok({}));
    const res = await core.getProjectId(".");
    assert.isTrue(res.isOk() && res.value === "");
  });
});
describe("getProjectMetadata", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(fs, "pathExistsSync").returns(true);
    sandbox.stub(fs, "readFileSync").returns("version: 1.1.1\nprojectId: 12345" as any);
    const core = new FxCore(tools);
    const res = await core.getProjectMetadata(".");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {
        projectId: "12345",
        version: "1.1.1",
      });
    }
  });
  it("yml not exist", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(fs, "pathExistsSync").resolves(false);
    const core = new FxCore(tools);
    const res = await core.getProjectMetadata(".");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {});
    }
  });
  it("throw error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(fs, "pathExistsSync").throws(new Error("mocked error"));
    const core = new FxCore(tools);
    const res = await core.getProjectMetadata(".");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {});
    }
  });
});
describe("getTeamsAppName", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {
      projectId: "12345",
      provision: {
        name: "provision",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: {
              name: "testappname-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
        ],
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isOk() && res.value === "testappname-");
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {
      projectId: "12345",
      provision: {
        name: "provision",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: {
              name: "testappname${{APP_NAME_SUFFIX}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
        ],
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isOk() && res.value === "testappname");
  });
  it("return empty value", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {};
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isOk() && res.value === "");
  });
  it("parse yml error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(metadataUtil, "parse").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.getTeamsAppName(".");
    assert.isTrue(res.isErr());
  });
});

describe("getProjectInfo", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    const mockProjectModel: any = {
      projectId: "mock-project-id",
      provision: {
        name: "provision",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: {
              name: "testappname-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
        ],
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "readEnv").resolves(
      ok({
        TEAMS_APP_ID: "mock-team-app-id",
        TEAMS_APP_TENANT_ID: "mock-tenant-id",
      })
    );
    const core = new FxCore(tools);
    const res = await core.getProjectInfo(".", "dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {
        projectId: "mock-project-id",
        teamsAppId: "mock-team-app-id",
        m365TenantId: "mock-tenant-id",
        teamsAppName: "testappname",
      });
    }
  });
  it("parse yml error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(metadataUtil, "parse").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.getProjectInfo(".", "dev");
    assert.isTrue(res.isErr());
  });
  it("read env error", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("./teamsapp.yml");
    sandbox.stub(metadataUtil, "parse").resolves(ok({} as any));
    sandbox.stub(envUtil, "readEnv").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.getProjectInfo(".", "dev");
    assert.isTrue(res.isErr());
  });
});

describe("checkProjectType", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy 1", async () => {
    sandbox.stub(projectTypeChecker, "checkProjectType").resolves({
      isTeamsFx: false,
      lauguages: [],
      hasTeamsManifest: false,
      dependsOnTeamsJs: false,
    });
    const core = new FxCore(tools);
    const res = await core.checkProjectType("");
    assert.isTrue(res.isOk());
  });

  it("happy 2", async () => {
    sandbox.stub(projectTypeChecker, "checkProjectType").resolves({
      isTeamsFx: true,
      teamsfxConfigType: TeamsfxConfigType.teamsappYml,
      teamsfxConfigVersion: "1.0.0",
      teamsfxVersionState: TeamsfxVersionState.Compatible,
      teamsfxProjectId: "xxxx-xxxx-xxxx",
      lauguages: [],
      hasTeamsManifest: true,
      manifestCapabilities: ["bot"],
      manifestAppId: "xxx",
      manifestVersion: "1.17",
      dependsOnTeamsJs: true,
    });
    const core = new FxCore(tools);
    const res = await core.checkProjectType("");
    assert.isTrue(res.isOk());
  });
});

describe("isEnvFile", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("file patten not match", async () => {
    const core = new FxCore(tools);
    const res = await core.isEnvFile(".", ".abc.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });
  it("getEnvFolderPath return error", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(err(new UserError({})));
    const core = new FxCore(tools);
    const res = await core.isEnvFile(".", ".env.dev");
    assert.isTrue(res.isErr());
  });
  it("getEnvFolderPath return undefined", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok(undefined));
    const core = new FxCore(tools);
    const res = await core.isEnvFile(".", ".env.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });
  it("folder not match", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("/tmp"));
    const core = new FxCore(tools);
    const res = await core.isEnvFile("/tmp", "/tmp1/.env.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });
  it("match", async () => {
    sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("/tmp"));
    const core = new FxCore(tools);
    const res = await core.isEnvFile("/tmp", "/tmp/.env.dev");
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value);
    }
  });
});
describe("getQuestions", async () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn = () => {};
  afterEach(() => {
    sandbox.restore();
    mockedEnvRestore();
  });
  it("happy path", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "with-plugin",
        "api-plugin-type",
        "plugin-manifest-path",
        "plugin-opeanapi-spec-path",
        "api-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "azure-openai-deployment-name",
        "openai-key",
        "folder",
        "app-name",
      ]);
    }
  });
  it("happy path with runtime", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "true",
    });
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "runtime",
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "with-plugin",
        "api-plugin-type",
        "plugin-manifest-path",
        "plugin-opeanapi-spec-path",
        "api-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "azure-openai-deployment-name",
        "openai-key",
        "folder",
        "app-name",
      ]);
    }
  });

  it("happy path: API Copilot plugin enabled", async () => {
    const restore = mockedEnv({});
    const core = new FxCore(tools);
    const res = await core.getQuestions(Stage.create, { platform: Platform.CLI_HELP });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const node = res.value;
      const names: string[] = [];
      collectNodeNames(node!, names);
      assert.deepEqual(names, [
        "capabilities",
        "bot-host-type-trigger",
        "spfx-solution",
        "spfx-install-latest-package",
        "spfx-framework-type",
        "spfx-webpart-name",
        "spfx-folder",
        "me-architecture",
        "with-plugin",
        "api-plugin-type",
        "plugin-manifest-path",
        "plugin-opeanapi-spec-path",
        "api-auth",
        "custom-copilot-rag",
        "openapi-spec-location",
        "api-operation",
        "custom-copilot-agent",
        "programming-language",
        "llm-service",
        "azure-openai-key",
        "azure-openai-endpoint",
        "azure-openai-deployment-name",
        "openai-key",
        "folder",
        "app-name",
      ]);
    }
    restore();
  });

  function collectNodeNames(node: IQTreeNode, names: string[]) {
    if (node.data.type !== "group") {
      names.push(node.data.name);
    }
    if (node.children) {
      for (const child of node.children) {
        collectNodeNames(child, names);
      }
    }
  }
});
describe("copilotPlugin", async () => {
  let mockedEnvRestore: RestoreFn = () => {};

  afterEach(() => {
    sinon.restore();
    mockedEnvRestore();
  });

  it("add API - json", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
  });

  it("add API - VS platform", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VS,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    const showMessage = sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(showMessage.calledOnce);
  });

  it("add API - Copilot plugin", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      [QuestionNames.DestinationApiSpecFilePath]: "destination.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      plugins: [
        {
          file: "ai-plugin.json",
          id: "plugin1",
        },
      ],
    };

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generateForCopilot").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(manifestUtils, "getPluginFilePath").resolves(ok("ai-plugin.json"));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves([]);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sinon.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    const result = await core.copilotPluginAddAPI(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
  });

  it("add API missing required input - Copilot plugin", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      plugins: [
        {
          file: "ai-plugin.json",
          id: "plugin1",
        },
      ],
    };
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generateForCopilot").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves([]);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.isTrue(result.error instanceof MissingRequiredInputError);
    }
  });

  it("add API error when getting plugin path - Copilot plugin", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: "manifest.json",
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      [QuestionNames.DestinationApiSpecFilePath]: "destination.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      plugins: [
        {
          file: "ai-plugin.json",
          id: "plugin1",
        },
      ],
    };
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generateForCopilot").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon
      .stub(manifestUtils, "getPluginFilePath")
      .resolves(err(new SystemError("testError", "testError", "", "")));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves([]);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "testError");
    }
  });
  it("add API - return multiple auth", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getUserById2",
          server: "https://server2",
          api: "GET /user/{userId2}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          auth: {
            name: "oauth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 3,
      allAPICount: 3,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sinon.stub(pluginGeneratorHelper, "injectAuthAction").resolves(undefined as any);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
  });

  it("add API - return multiple server error", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server2",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "MultipleServerError");
    }
  });

  it("add API - no provision section in teamsapp yaml file - OAuth", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "oauthAuth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "oauthAuth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      version: "1.0.0",
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectOAuthActionFailedError");
    }
  });

  it("add API - no provision section in teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      version: "1.0.0",
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectAPIKeyActionFailedError");
    }
  });

  it("add API - no teamsApp/create action in teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectAPIKeyActionFailedError");
    }
  });

  it("add API - no teams app id in teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            otherEnv: "OtherEnv",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectAPIKeyActionFailedError");
    }
  });

  it("add API - no teams app id in teamsapp yaml file - OAuth", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "oauthAuth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "oauthAuth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            otherEnv: "OtherEnv",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal((result.error as FxError).name, "InjectOAuthActionFailedError");
    }
  });

  it("add API - should inject api key action to teamsapp yaml file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should not inject api key action to teamsapp yaml file when already exists", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            name: "bearerAuth1",
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "BEARERAUTH1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
  });

  it("add API - should inject api key action to teamsapp yaml file when name does not match", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            name: "api_key2",
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "API_KEY1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "api_key2",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "API_KEY1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should inject api key action to teamsapp yaml file when missing with in yaml", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          writeToEnvironmentFile: {
            registrationId: "BEARERAUTH1_REGISTRATION_ID1",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "apiKey/register",
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID1",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should inject api key action to teamsapp yaml file when missing name in yaml", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "API_KEY1_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return !path.endsWith("yml");
    });
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);

      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "API_KEY1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 1);
  });

  it("add API - should inject api key action to teamsapp yaml file with local teamsapp file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);
      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 2);
  });

  it("add API - should inject oauth action to teamsapp yaml file with local teamsapp file", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "oauthAuth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "oauthAuth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationUrl",
                  tokenUrl: "mockedTokenUrl",
                  scopes: {
                    mockedScope: "description for mocked scope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);
      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "oauth/register",
            with: {
              name: "oauthAuth",
              flow: "authorizationCode",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              configurationId: "OAUTHAUTH_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 2);
  });

  it("add API - should filter unknown api key action", async () => {
    const appName = await mockV3Project();
    mockedEnvRestore = mockedEnv({
      TEAMSFX_CLI_DOTNET: "false",
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}", "GET /store/order"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server1",
          api: "GET /user/{userId}",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server1",
          api: "GET /store/order",
          auth: {
            name: "bearerAuth1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [],
      allSuccess: true,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const teamsappObject = {
      provision: [
        {
          uses: "teamsApp/create",
          with: {
            name: "dfefeef-${{TEAMSFX_ENV}}",
          },
          writeToEnvironmentFile: {
            teamsAppId: "TEAMS_APP_ID",
          },
        },
        {
          uses: "apiKey/register",
          with: {
            name: "api_key_unknown",
            appId: "${{TEAMS_APP_ID}}",
            apiSpecPath: ".appPackage/apiSpecificationFiles/openapi.json",
          },
          writeToEnvironmentFile: {
            registrationId: "API_KEY_UNKNOWN_REGISTRATION_ID",
          },
        },
        {
          uses: "teamsApp/zipAppPackage",
          with: {
            manifestPath: "./appPackage/manifest.json",
            outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
            outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
          },
        },
      ],
    };
    const yamlString = jsyaml.dump(teamsappObject);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(yamlString as any);

    let writeYamlObjectTriggeredTimes = 0;
    sinon.stub(fs, "writeFile").callsFake((_, yamlString) => {
      writeYamlObjectTriggeredTimes++;
      const yamlObject = jsyaml.load(yamlString);
      assert.deepEqual(yamlObject, {
        provision: [
          {
            uses: "teamsApp/create",
            with: {
              name: "dfefeef-${{TEAMSFX_ENV}}",
            },
            writeToEnvironmentFile: {
              teamsAppId: "TEAMS_APP_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "bearerAuth1",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: "./appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "BEARERAUTH1_REGISTRATION_ID",
            },
          },
          {
            uses: "apiKey/register",
            with: {
              name: "api_key_unknown",
              appId: "${{TEAMS_APP_ID}}",
              apiSpecPath: ".appPackage/apiSpecificationFiles/openapi.json",
            },
            writeToEnvironmentFile: {
              registrationId: "API_KEY_UNKNOWN_REGISTRATION_ID",
            },
          },
          {
            uses: "teamsApp/zipAppPackage",
            with: {
              manifestPath: "./appPackage/manifest.json",
              outputZipPath: "./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip",
              outputJsonPath: "./appPackage/build/manifest.${{TEAMSFX_ENV}}.json",
            },
          },
        ],
      });
    });

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeYamlObjectTriggeredTimes === 2);
  });

  it("add API - warnings", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [
          {
            id: "getUserById",
            title: "Get User By Id",
          },
          {
            id: "notexist",
            title: "Get User By Id",
          },
        ],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [
        {
          type: WarningType.OperationOnlyContainsOptionalParam,
          content: "fakeMessage",
          data: { commandId: "fakeId", parameterName: "fakeName" },
        },
      ],
      allSuccess: false,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("warning message");
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const logSpy = sinon.spy(tools.logProvider, "info");
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(logSpy.calledOnce);
  });

  it("add API - unknown warning not show log", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ManifestPath]: path.join(os.tmpdir(), appName, "appPackage/manifest.json"),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [
          {
            id: "getUserById",
            title: "Get User By Id",
          },
          {
            id: "notexist",
            title: "Get User By Id",
          },
        ],
      },
    ];

    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };

    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").resolves({
      warnings: [
        {
          type: "unknown" as any,
          content: "fakeMessage",
          data: { commandId: "fakeId", parameterName: "fakeName" },
        },
      ],
      allSuccess: false,
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const logSpy = sinon.spy(tools.logProvider, "info");
    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(logSpy.notCalled);
  });

  it("add API - readManifestFailed", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "generate").throws(new Error("fakeError"));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
  });

  it("add API - assembleError", async () => {
    const appName = await mockV3Project();
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon.stub(SpecParser.prototype, "list").throws(new Error("fakeError"));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
  });

  it("add API - SpecParserError", async () => {
    const appName = await mockV3Project();
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const listResult: ListAPIResult = {
      APIs: [
        {
          operationId: "getUserById",
          server: "https://server",
          api: "GET /user/{userId}",
          isValid: true,
          reason: [],
        },
        {
          operationId: "getStoreOrder",
          server: "https://server",
          api: "GET /store/order",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 2,
      allAPICount: 2,
    };
    sinon.stub(SpecParser.prototype, "validate").resolves({
      warnings: [],
      status: ValidationStatus.Valid,
      errors: [],
    });
    sinon.stub(SpecParser.prototype, "list").resolves(listResult);
    sinon
      .stub(SpecParser.prototype, "generate")
      .throws(new SpecParserError("", ErrorType.FilterSpecFailed));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
  });

  it("add API - ui error", async () => {
    const appName = await mockV3Project();
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon
      .stub(tools.ui, "showMessage")
      .resolves(err(new UserError("testSource", "testError", "", "")));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "testError");
    }
  });

  it("add API - not 'add' when confirm", async () => {
    const appName = await mockV3Project();
    const manifest = new TeamsAppManifest();
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "apiSpecificationFiles/openapi.json",
        commands: [],
      },
    ];
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["testOperation"],
      [QuestionNames.ManifestPath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    sinon
      .stub(SpecParser.prototype, "generate")
      .throws(new SpecParserError("fakeMessage", ErrorType.SpecNotValid));
    sinon.stub(validationUtils, "validateInputs").resolves(undefined);
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sinon.stub(tools.ui, "showMessage").resolves(ok(""));

    const result = await core.copilotPluginAddAPI(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.isTrue(result.error instanceof UserCancelError);
    }
  });

  it("isDeclarativeAgentApp - invalid project path", async () => {
    const core = new FxCore(tools);
    const inputs = { projectPath: "invalid" } as Inputs;
    const res = await core.isDelcarativeAgentApp(inputs);
    assert.isTrue(res.isErr());
  });

  it("isDeclarativeAgentApp - true", async () => {
    const core = new FxCore(tools);
    const manifest = new TeamsAppManifest();
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          id: "1",
          file: "file",
        },
      ],
    };
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    const inputs = { projectPath: "mock" } as Inputs;
    const res = await core.isDelcarativeAgentApp(inputs);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value);
    }
  });

  it("isDeclarativeAgentApp - false", async () => {
    const core = new FxCore(tools);
    const manifest = new TeamsAppManifest();
    sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    const inputs = { projectPath: "mock" } as Inputs;
    const res = await core.isDelcarativeAgentApp(inputs);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isFalse(res.value);
    }
  });

  describe("listPluginApiSpecs", async () => {
    it("success", async () => {
      const inputs = {
        [QuestionNames.ManifestPath]: "manifest.json",
        platform: Platform.VS,
      };
      const manifest = new TeamsAppManifest();
      manifest.copilotExtensions = {
        plugins: [
          {
            file: "ai-plugin.json",
            id: "plugin1",
          },
        ],
      };
      sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
      sinon
        .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
        .resolves(ok(["apispec.json"]));

      const core = new FxCore(tools);
      const res = await core.listPluginApiSpecs(inputs);

      assert.isTrue(res.isOk());
    });

    it("read manifest error", async () => {
      const inputs = {
        [QuestionNames.ManifestPath]: "manifest.json",
        platform: Platform.VS,
      };
      sinon
        .stub(manifestUtils, "_readAppManifest")
        .resolves(err(new SystemError("read manifest error", "read manifest error", "", "")));

      const core = new FxCore(tools);
      const res = await core.listPluginApiSpecs(inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "read manifest error");
      }
    });

    it("get api spec error", async () => {
      const inputs = {
        [QuestionNames.ManifestPath]: "manifest.json",
        platform: Platform.VS,
      };
      const manifest = new TeamsAppManifest();
      manifest.copilotExtensions = {
        plugins: [
          {
            file: "ai-plugin.json",
            id: "plugin1",
          },
        ],
      };
      sinon.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
      sinon
        .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
        .resolves(err(new SystemError("get plugin error", "get plugin error", "", "")));

      const core = new FxCore(tools);
      const res = await core.listPluginApiSpecs(inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "get plugin error");
      }
    });
  });

  it("load operations - should return a list of operations when given valid inputs", async () => {
    const core = new FxCore(tools);
    const inputs = {
      manifest: {},
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
    };
    const expectedResult = [
      {
        id: "operation1",
        label: "operation1",
        groupName: "1",
        data: { serverUrl: "https://server1" },
      },
      {
        id: "operation2",
        label: "operation2",
        groupName: "2",
        data: { serverUrl: "https://server2" },
      },
    ];
    sinon.stub(CopilotPluginHelper, "listOperations").returns(Promise.resolve(ok(expectedResult)));
    const result = await core.copilotPluginListOperations(inputs as any);
    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.deepEqual(result.value, expectedResult);
    }
  });

  it("load operations - should return an error when an exception is thrown", async () => {
    const core = new FxCore(tools);
    const inputs = {
      manifest: {},
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
    };
    sinon.stub(CopilotPluginHelper, "listOperations").returns(Promise.resolve(err([])));
    const result = await core.copilotPluginListOperations(inputs as any);
    assert.isTrue(result.isErr());
  });

  it("load operations - no manifest in inputs", async () => {
    const core = new FxCore(tools);
    const inputs = {
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
      includeExistingAPIs: false,
    };

    sinon
      .stub(SpecParser.prototype, "validate")
      .resolves({ status: ValidationStatus.Valid, warnings: [], errors: [] });
    sinon
      .stub(SpecParser.prototype, "list")
      .resolves({ APIs: [], allAPICount: 0, validAPICount: 0 });

    try {
      await core.copilotPluginListOperations(inputs as any);
    } catch (e: any) {
      assert.equal(e.name, MissingRequiredInputError.name);
    }
  });

  it("load operations - invalid manifest", async () => {
    const core = new FxCore(tools);
    const inputs = {
      apiSpecUrl: "https://example.com/api-spec",
      shouldLogWarning: true,
      includeExistingAPIs: false,
      "manifest-path": "fakePath",
    };

    sinon
      .stub(manifestUtils, "_readAppManifest")
      .returns(Promise.resolve(err(new FileNotFoundError("file", "fakePath"))));
    sinon
      .stub(SpecParser.prototype, "validate")
      .resolves({ status: ValidationStatus.Valid, warnings: [], errors: [] });
    sinon
      .stub(SpecParser.prototype, "list")
      .resolves({ APIs: [], allAPICount: 0, validAPICount: 0 });

    try {
      await core.copilotPluginListOperations(inputs as any);
    } catch (e: any) {
      assert.equal(e.name, FileNotFoundError.name);
    }
  });

  it("teamsapp management APIs", async () => {
    const core = new FxCore(tools);
    const inputs = {
      platform: Platform.CLI,
    };
    sinon.stub(teamsappMgr, "updateTeamsApp").resolves(ok(undefined));
    sinon
      .stub(teamsappMgr, "packageTeamsApp")
      .resolves(ok({ manifestPath: "", outputJsonPath: "", outputZipPath: "" }));
    sinon.stub(teamsappMgr, "validateTeamsApp").resolves(ok(undefined));
    sinon.stub(teamsappMgr, "publishTeamsApp").resolves(ok(undefined));
    const res1 = await core.updateTeamsAppCLIV3(inputs as any);
    const res2 = await core.packageTeamsAppCLIV3(inputs as any);
    const res3 = await core.validateTeamsAppCLIV3(inputs as any);
    const res4 = await core.publishTeamsAppCLIV3(inputs as any);
    assert.isTrue(res1.isOk());
    assert.isTrue(res2.isOk());
    assert.isTrue(res3.isOk());
    assert.isTrue(res4.isOk());
  });
});

describe("addPlugin", async () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    setTools(tools);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("from API spec: add action success", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    assert.isTrue(openFileStub.calledOnce);

    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from API spec: add action success with bearer token auth and teamsapp.local.yaml exist", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();
    sandbox.stub(fs, "readFile").resolves("{{test}}" as any);
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      if (path.endsWith(MetadataV3.localConfigFile)) {
        return true;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));
    sandbox
      .stub(CopilotPluginHelper, "injectAuthAction")
      .resolves({ defaultRegistrationIdEnvName: "test", registrationIdEnvName: "test2" });

    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
          auth: {
            name: "bearerAuth",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    assert.isTrue(openFileStub.calledOnce);
    assert.equal(writeFileStub.args[0][1], "{{test2}}");

    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from API spec: add action success with oauth token auth and without teamsapp.local.yaml", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();
    sandbox.stub(fs, "readFile").resolves("{{test}}" as any);
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      if (path.endsWith(MetadataV3.localConfigFile)) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));
    sandbox
      .stub(CopilotPluginHelper, "injectAuthAction")
      .resolves({ defaultRegistrationIdEnvName: "test", registrationIdEnvName: "test2" });

    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
          auth: {
            name: "oauth2",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://example.com/auth",
                  tokenUrl: "https://example.com/token",
                  scopes: {
                    "user.read": "Read user profile",
                  },
                },
              },
            },
          },
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    assert.isTrue(openFileStub.calledOnce);
    assert.equal(writeFileStub.args[0][1], "{{test2}}");

    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from API spec: empty declarativeCopilots 1", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {};
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.isTrue(result.error instanceof UserError);
    }
  });

  it("from API spec: empty declarativeCopilots 2", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.isTrue(result.error instanceof UserError);
    }
  });

  it("from API spec: add action success - copilot agent", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    assert.isTrue(openFileStub.calledOnce);

    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from API spec: add action with warnings from CLI", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.CLI,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.yaml",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("warning message");
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const core = new FxCore(tools);
    sandbox
      .stub(CopilotPluginHelper, "generateFromApiSpec")
      .resolves(
        ok({ warnings: [{ type: WarningType.OperationOnlyContainsOptionalParam, content: "" }] })
      );

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.addPlugin(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from existing plugin: add action success and not view plugin manifest", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.CLI,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.PluginManifestFilePath]: "ai-plugin.json",
      [QuestionNames.PluginOpenApiSpecFilePath]: "openapi.json",
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.existingPlugin().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));

    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotExtensionHelper, "addExistingPlugin")
      .resolves(ok({ destinationPluginManifestPath: "ai-plugin.json", warnings: [] }));

    const core = new FxCore(tools);

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }

    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from existing plugin: add action error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.CLI,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.PluginManifestFilePath]: "ai-plugin.json",
      [QuestionNames.PluginOpenApiSpecFilePath]: "openapi.json",
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.existingPlugin().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));

    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotExtensionHelper, "addExistingPlugin")
      .resolves(err(new SystemError("fakeError", "fakeError", "", "")));

    sandbox.stub(tools.ui, "showMessage").resolves(ok("Add"));

    const core = new FxCore(tools);

    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }

    assert.isTrue(result.isErr() && result.error.name === "fakeError");
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("from API Spec: generateForCopilot error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.json")) {
        return false;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return false;
      }
      return true;
    });
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(tools.ui, "showMessage").resolves(ok("Add"));
    sandbox
      .stub(CopilotPluginHelper, "generateFromApiSpec")
      .resolves(err(new SystemError("", "", "", "")));
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
  });

  it("from API spec: add action error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.json")) {
        return false;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(err(new SystemError("addActionError", "addActionError", "", "")));

    const core = new FxCore(tools);
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));

    sandbox.stub(tools.ui, "showMessage").resolves(ok("Add"));
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "addActionError");
    }
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("error: read Teams manifest error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new SystemError("manifestError", "manifestError", "", "")));
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "manifestError");
    }
  });

  it("error: get declarative copilot manifest path error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(err(new SystemError("getError", "getError", "", "")));
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "getError");
    }
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("error: read GPT manifest error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(err(new SystemError("readError", "readError", "", "")));
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "readError");
    }
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("error: not copilot GPT project", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, AppStudioError.TeamsAppRequiredPropertyMissingError.name);
    }
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("error: cancel", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(tools.ui, "showMessage").resolves(ok("Cancel"));
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.isTrue(result.error instanceof UserCancelError);
    }
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("error: confirm UI error", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiSpecLocation]: "test.json",
      [QuestionNames.ApiOperation]: ["GET /user/{userId}"],
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(tools.ui, "showMessage")
      .resolves(err(new SystemError("uiError", "uiError", "", "")));
    const core = new FxCore(tools);
    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal("uiError", result.error.name);
    }
    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }
  });

  it("call kiota: success redirect to Kiota", async () => {
    const mockedEnvRestore = mockedEnv({
      [FeatureFlagName.KiotaIntegration]: "true",
    });

    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const core = new FxCore(tools);

    const result = await core.addPlugin(inputs);
    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.equal(result.value.lastCommand, "addPlugin");
      assert.equal(result.value.manifestPath, "manifest.json");
    }

    mockedEnvRestore();
  });

  it("call kiota: success create project with input", async () => {
    const mockedEnvRestore = mockedEnv({
      [FeatureFlagName.KiotaIntegration]: "true",
    });

    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "aiplugin-apiplugin.json",
      [QuestionNames.ApiSpecLocation]: "spec-apimanifest.yaml",
      [QuestionNames.ApiOperation]: "aiplugin-apiplugin.json",
      projectPath: path.join(os.tmpdir(), appName),
    };

    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");
    sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
      if (path.endsWith("openapi_1.yaml")) {
        return true;
      }
      if (path.endsWith("ai-plugin_1.json")) {
        return true;
      }
      if (path.endsWith("openapi_2.yaml")) {
        return false;
      }
      if (path.endsWith("ai-plugin_2.json")) {
        return false;
      }
      if (path.endsWith("aiplugin-apiplugin.json")) {
        return true;
      }
      if (path.endsWith("spec-apimanifest.yaml")) {
        return true;
      }
      if (path.endsWith("aiplugin_1-apiplugin.json")) {
        return false;
      }
      if (path.endsWith("spec_1-apimanifest.yaml")) {
        return false;
      }
      return true;
    });
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));

    const core = new FxCore(tools);
    sandbox
      .stub(CopilotPluginHelper, "generateFromApiSpec")
      .callsFake(
        async (
          specParser,
          teamsManifestPath,
          inputs,
          context,
          component,
          projectType,
          outputFilePath
        ) => {
          assert.isTrue(outputFilePath.destinationApiSpecFilePath.includes("apimanifest.yaml"));
          assert.isTrue(outputFilePath.pluginManifestFilePath?.includes("apiplugin.json"));
          return ok({ warnings: [] });
        }
      );

    const showMessageStub = sandbox
      .stub(tools.ui, "showMessage")
      .callsFake((level, message, modal, items) => {
        if (level == "info") {
          return Promise.resolve(
            ok(getLocalizedString("core.addPlugin.success.viewPluginManifest"))
          );
        } else if (level === "warn") {
          return Promise.resolve(ok("Add"));
        } else {
          throw new NotImplementedError("TEST", "showMessage");
        }
      });

    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.addPlugin(inputs);
    if (result.isErr()) {
      console.log(result.error);
    }
    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledTwice);
    assert.isTrue(openFileStub.calledOnce);

    if (await fs.pathExists(inputs.projectPath!)) {
      await fs.remove(inputs.projectPath!);
    }

    mockedEnvRestore();
  });

  describe("projectVersionCheck", async () => {
    it("invalid project", async () => {
      sandbox.stub(projectHelper, "isValidProjectV3").returns(false);
      sandbox.stub(projectHelper, "isValidProjectV2").returns(false);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: os.tmpdir(),
        projectPath: "./",
      };
      const core = new FxCore(tools);
      const result = await core.projectVersionCheck(inputs);
      assert.isTrue(result.isErr());
    });
    it("version is undefined", async () => {
      sandbox.stub(projectHelper, "isValidProjectV3").returns(true);
      sandbox
        .stub(migrationUtil, "getProjectVersionFromPath")
        .resolves({ version: "", source: VersionSource.teamsapp });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: os.tmpdir(),
        projectPath: "./",
      };
      const core = new FxCore(tools);
      const result = await core.projectVersionCheck(inputs);
      assert.isTrue(result.isErr());
    });
    it("no plugin", async () => {
      sandbox.stub(projectHelper, "isValidProjectV3").returns(true);
      sandbox
        .stub(migrationUtil, "getProjectVersionFromPath")
        .resolves({ version: "1.0", source: VersionSource.teamsapp });
      sandbox.stub(migrationUtil, "getTrackingIdFromPath").resolves("xxxx-xxxx");
      sandbox.stub(migrationUtil, "getVersionState").returns(VersionState.upgradeable);
      sandbox.stub(projMigrator, "checkActiveResourcePlugins").resolves(false);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: os.tmpdir(),
        projectPath: "./",
      };
      const core = new FxCore(tools);
      const result = await core.projectVersionCheck(inputs);
      assert.isTrue(result.isErr());
    });
    it("sync Manifest - success", async () => {
      const core = new FxCore(tools);
      const inputs = {
        platform: Platform.CLI_HELP,
        projectPath: "fake",
        env: "dev",
        nonInteractive: true,
      };
      sandbox.stub(SyncManifestDriver.prototype, "sync").resolves(ok(new Map<string, string>()));
      const res = await core.syncManifest(inputs as SyncManifestInputs);
      assert.isTrue(res.isOk());
    });
    it("sync Manifest - default CLI project path", async () => {
      const core = new FxCore(tools);
      const inputs = {
        platform: Platform.CLI_HELP,
        env: "dev",
        nonInteractive: true,
        ignoreLockByUT: true,
      };
      const defaultProjectPath = "./";
      sandbox
        .stub(SyncManifestDriver.prototype, "sync")
        .callsFake(async (args: SyncManifestArgs, context: WrapDriverContext) => {
          assert.isTrue(args.projectPath === defaultProjectPath);
          return ok(new Map<string, string>());
        });
      const res = await core.syncManifest(inputs as SyncManifestInputs);
      assert.isTrue(res.isOk());
    });
    it("sync Manifest - default VSC project path", async () => {
      const core = new FxCore(tools);
      const inputs = {
        platform: Platform.VSCode,
        env: "dev",
        nonInteractive: true,
        ignoreLockByUT: true,
      };
      const defaultProjectPath = path.join(os.homedir(), ConstantString.RootFolder);
      sandbox
        .stub(SyncManifestDriver.prototype, "sync")
        .callsFake(async (args: SyncManifestArgs, context: WrapDriverContext) => {
          assert.isTrue(args.projectPath === defaultProjectPath);
          return ok(new Map<string, string>());
        });
      const res = await core.syncManifest(inputs as SyncManifestInputs);
      assert.isTrue(res.isOk());
    });
  });
});

describe("kiotaRegenerate", async () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn;

  beforeEach(() => {
    setTools(tools);
    mockedEnvRestore = mockedEnv({
      [FeatureFlagName.KiotaIntegration]: "true",
    });
  });

  afterEach(() => {
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path: successfully regenerate", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(
      ok({
        actions: [
          {
            id: "action_1",
            file: "test-aiplugin.json",
          },
        ],
      } as DeclarativeCopilotManifestSchema)
    );
    sandbox.stub(fs, "readJson").resolves({
      capabilities: {
        conversation_starters: [],
      },
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "apiSpecificationFile/openapi.json",
          },
          run_for_functions: ["listRepairs"],
        },
      ],
      functions: [
        {
          name: "listRepairs",
          description: "List all repairs",
        },
      ],
    } as any);

    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
          auth: {
            name: "bearerAuth",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox
      .stub(CopilotPluginHelper, "injectAuthAction")
      .resolves({ defaultRegistrationIdEnvName: "test", registrationIdEnvName: "test" });
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isOk());
  });

  it("happy path: add new action", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };

    sandbox.stub(fs, "readJson").resolves({
      capabilities: {
        conversation_starters: [],
      },
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "apiSpecificationFile/openapi.json",
          },
          run_for_functions: ["listRepairs"],
        },
      ],
      functions: [
        {
          name: "listRepairs",
          description: "List all repairs",
        },
      ],
    } as any);
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(
      ok({
        actions: [
          {
            id: "action_1",
            file: "test-aiplugin1.json",
          },
        ],
      } as DeclarativeCopilotManifestSchema)
    );

    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
          auth: {
            name: "bearerAuth",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox
      .stub(CopilotPluginHelper, "injectAuthAction")
      .resolves({ defaultRegistrationIdEnvName: "test", registrationIdEnvName: "test" });
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(CopilotPluginHelper, "generateFromApiSpec").resolves(ok({ warnings: [] }));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isOk());
  });

  it("should throw error if fail to update plugin manifest", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(
      ok({
        actions: [
          {
            id: "action_1",
            file: "test-aiplugin.json",
          },
        ],
      } as DeclarativeCopilotManifestSchema)
    );
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
          auth: {
            name: "bearerAuth",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox
      .stub(CopilotPluginHelper, "injectAuthAction")
      .resolves({ defaultRegistrationIdEnvName: "test", registrationIdEnvName: "test" });
    sandbox
      .stub(CopilotPluginHelper, "generateFromApiSpec")
      .resolves(err(new UserError("fake-error-source", "fake-error-name", "fake-error-message")));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());
  });

  it("should throw error if no project path", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
    };

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());
  });

  it("should throw error if failed to read app manifest", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new UserError("fakeError", "fakeError", "", "")));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "fakeError");
    }
  });

  it("should throw error if no copilotAgents in manifest", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "TeamsAppMissingRequiredCapability");
    }
  });

  it("should throw error if failed to getManifestPath", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(err(new UserError("fakeError", "fakeError", "", "")));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "fakeError");
    }
  });

  it("should throw error if failed to readCopilotGptManifestFile", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(err(new UserError("fakeError", "fakeError", "", "")));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "fakeError");
    }
  });

  it("should throw error if failed to add action", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.TeamsAppManifestFilePath]: "manifest.json",
      [QuestionNames.ApiPluginManifestPath]: "test-aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          file: "test1.json",
          id: "action_1",
        },
      ],
    };

    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok("dcManifest.json"));
    sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(
      ok({
        actions: [
          {
            id: "action_1",
            file: "test-aiplugin1.json",
          },
        ],
      } as DeclarativeCopilotManifestSchema)
    );
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /user/{userId}",
          server: "https://example.com",
          operationId: "getExample",
          isValid: true,
          reason: [],
          auth: {
            name: "bearerAuth",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sandbox
      .stub(CopilotPluginHelper, "injectAuthAction")
      .resolves({ defaultRegistrationIdEnvName: "test", registrationIdEnvName: "test" });
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(err(new UserError("fakeError", "fakeError", "", "")));

    const core = new FxCore(tools);
    const result = await core.kiotaRegenerate(inputs);
    assert.isTrue(result.isErr());

    if (result.isErr()) {
      assert.equal(result.error.name, "fakeError");
    }
  });
});

describe("addAuthAction", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  it("happy path: successfully add auth action for api key", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.PluginManifestFilePath]: "aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      [QuestionNames.ApiOperation]: ["operation1"],
      [QuestionNames.AuthName]: "mockAuthName",
      [QuestionNames.ApiAuth]: "api-key",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const pluginManifest = {
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(pluginGeneratorHelper, "injectAuthAction").resolves({
      defaultRegistrationIdEnvName: "test",
      registrationIdEnvName: "test",
    });
    sandbox.stub(path, "normalize").returns("normalizedPath");
    sandbox.stub(path, "join").returns("joinedPath");
    sandbox.stub(fs, "readJson").resolves(pluginManifest);
    sandbox.stub(fs, "writeJson").callsFake(async (path: string, data: any) => {
      assert.equal(data.runtimes.length, 2);
    });
    const core = new FxCore(tools);
    const result = await core.addAuthAction(inputs);
    assert.isTrue(result.isOk());
  });

  it("happy path: successfully add auth action for oauth", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.PluginManifestFilePath]: "aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      [QuestionNames.ApiOperation]: ["operation1"],
      [QuestionNames.AuthName]: "mockAuthName",
      [QuestionNames.ApiAuth]: "oauth",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const pluginManifest = {
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(pluginGeneratorHelper, "injectAuthAction").resolves({
      defaultRegistrationIdEnvName: "test",
      registrationIdEnvName: "test",
    });
    sandbox.stub(path, "normalize").returns("normalizedPath");
    sandbox.stub(path, "join").returns("joinedPath");
    sandbox.stub(fs, "readJson").resolves(pluginManifest);
    sandbox.stub(fs, "writeJson").callsFake(async (path: string, data: any) => {
      assert.equal(data.runtimes.length, 2);
    });
    const core = new FxCore(tools);
    const result = await core.addAuthAction(inputs);
    assert.isTrue(result.isOk());
  });

  it("happy path: successfully add auth action for microsoft entra", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.PluginManifestFilePath]: "aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      [QuestionNames.ApiOperation]: ["operation1"],
      [QuestionNames.AuthName]: "mockAuthName",
      [QuestionNames.ApiAuth]: "microsoft-entra",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const pluginManifest = {
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(pluginGeneratorHelper, "injectAuthAction").resolves({
      defaultRegistrationIdEnvName: "test",
      registrationIdEnvName: "test",
    });
    sandbox.stub(path, "normalize").returns("normalizedPath");
    sandbox.stub(path, "join").returns("joinedPath");
    sandbox.stub(fs, "readJson").resolves(pluginManifest);
    sandbox.stub(fs, "writeJson").callsFake(async (path: string, data: any) => {
      assert.equal(data.runtimes.length, 2);
    });
    const core = new FxCore(tools);
    const result = await core.addAuthAction(inputs);
    assert.isTrue(result.isOk());
  });

  it("happy path: should do nothing if auth action not added", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.PluginManifestFilePath]: "aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      [QuestionNames.ApiOperation]: ["operation1"],
      [QuestionNames.AuthName]: "mockAuthName",
      [QuestionNames.ApiAuth]: "api-key",
      projectPath: path.join(os.tmpdir(), appName),
    };
    const pluginManifest = {
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(pluginGeneratorHelper, "injectAuthAction").resolves(undefined);
    sandbox.stub(path, "normalize").returns("normalizedPath");
    sandbox.stub(path, "join").returns("joinedPath");
    sandbox.stub(fs, "readJson").resolves(pluginManifest);
    const writeJsonStub = sandbox.stub(fs, "writeJson").resolves();
    const core = new FxCore(tools);
    const result = await core.addAuthAction(inputs);
    assert.isTrue(result.isOk());
    assert.isTrue(writeJsonStub.notCalled);
  });

  it("should throw error when missing project path", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.PluginManifestFilePath]: "aiplugin.json",
      [QuestionNames.ApiSpecLocation]: "test-openapi.yaml",
      [QuestionNames.ApiOperation]: ["operation1"],
      [QuestionNames.AuthName]: "mockAuthName",
      [QuestionNames.ApiAuth]: "api-key",
    };
    const pluginManifest = {
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    };
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
    sandbox
      .stub(copilotGptManifestUtils, "readCopilotGptManifestFile")
      .resolves(ok({} as DeclarativeCopilotManifestSchema));
    sandbox.stub(pluginGeneratorHelper, "injectAuthAction").resolves({
      defaultRegistrationIdEnvName: "test",
      registrationIdEnvName: "test",
    });
    sandbox.stub(path, "normalize").returns("normalizedPath");
    sandbox.stub(path, "join").returns("joinedPath");
    sandbox.stub(fs, "readJson").resolves(pluginManifest);
    sandbox.stub(fs, "writeJson").callsFake(async (path: string, data: any) => {
      assert.equal(data.runtimes.length, 2);
    });
    const core = new FxCore(tools);
    const result = await core.addAuthAction(inputs);
    assert.isTrue(result.isErr());
  });
});

describe("addKnowledge", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  it("happy path: successfully add knowledge", async () => {
    const appName = await mockV3Project();
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: os.tmpdir(),
      projectPath: path.join(os.tmpdir(), appName),
    };
    const core = new FxCore(tools);
    const result = await core.addKnowledge(inputs);
    assert.isTrue(result.isOk());
  });
});
