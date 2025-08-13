// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Context,
  InputsWithProjectPath,
  ManifestUtil,
  Platform,
  TeamsAppManifest,
  TeamsManifest,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import * as chai from "chai";
import fs from "fs-extra";
import "mocha";
import { RestoreFn } from "mocked-env";
import sinon from "sinon";
import Container from "typedi";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { ExecutionResult } from "../../../../src/component/driver/interface/stepDriver";
import {
  checkIfAppInDifferentAcountSameTenant,
  getAppPackage,
  updateManifestV3,
  updateTeamsAppV3ForPublish,
} from "../../../../src/component/driver/teamsApp/appStudio";
import { ConfigureTeamsAppDriver } from "../../../../src/component/driver/teamsApp/configure";
import { CreateAppPackageDriver } from "../../../../src/component/driver/teamsApp/createAppPackage";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { RetryHandler } from "../../../../src/component/driver/teamsApp/utils/utils";
import { envUtil } from "../../../../src/component/utils/envUtil";
import { QuestionNames } from "../../../../src/question";
import {
  MockLogProvider,
  MockTools,
  MockedAzureAccountProvider,
  MockedM365Provider,
  randomAppName,
} from "../../../core/utils";
import { getAzureProjectRoot } from "../../../plugins/resource/appstudio/helper";

describe.skip("appStudio", () => {
  const tools = new MockTools();
  setTools(tools);
  const sandbox = sinon.createSandbox();
  describe("checkIfAppInDifferentAcountSameTenant", () => {
    const logger = new MockLogProvider();
    const teamsAppId = "teams";
    const m365TokenProvider = new MockedM365Provider();

    afterEach(() => {
      sandbox.restore();
    });

    it("get app successfully: returns false", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));
      sandbox.stub(teamsDevPortalClient, "getApp").resolves();

      const res = await checkIfAppInDifferentAcountSameTenant(
        teamsAppId,
        m365TokenProvider,
        logger
      );
      chai.assert.isTrue(res.isOk());

      if (res.isOk()) {
        chai.assert.isFalse(res.value);
      }
    });

    it("get token error: returns error", async () => {
      m365TokenProvider.getAccessToken = sandbox
        .stub()
        .returns(err(new UserError("token", "token", "", "")));

      const res = await checkIfAppInDifferentAcountSameTenant(
        teamsAppId,
        m365TokenProvider,
        logger
      );
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "token");
      }
    });

    it("app in tenant but different account: returns true", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));
      sandbox.stub(teamsDevPortalClient, "getApp").throws({ message: "404" });
      sandbox.stub(teamsDevPortalClient, "checkExistsInTenant").returns(Promise.resolve(true));
      const res = await checkIfAppInDifferentAcountSameTenant(
        teamsAppId,
        m365TokenProvider,
        logger
      );
      chai.assert.isTrue(res.isOk());

      if (res.isOk()) {
        chai.assert.isTrue(res.value);
      }
    });

    it("get app error (not 404): returns false", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));
      sandbox.stub(teamsDevPortalClient, "getApp").throws({ message: "401" });
      const res = await checkIfAppInDifferentAcountSameTenant(
        teamsAppId,
        m365TokenProvider,
        logger
      );
      chai.assert.isTrue(res.isOk());

      if (res.isOk()) {
        chai.assert.isFalse(res.value);
      }
    });
  });

  describe("getAppPackage", () => {
    const logger = new MockLogProvider();
    const teamsAppId = "teams";
    const m365TokenProvider = new MockedM365Provider();

    afterEach(() => {
      sandbox.restore();
    });

    it("get package successfully", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(""));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outline.png", new Buffer(""));
      zip.addFile("zh-cn.json", new Buffer(""));
      const archivedFile = zip.toBuffer();
      sandbox.stub(RetryHandler, "Retry").resolves({
        data: archivedFile,
      });

      const res = await getAppPackage(teamsAppId, m365TokenProvider, logger);
      chai.assert.isTrue(res.isOk());

      if (res.isOk()) {
        chai.assert.isTrue(res.value.manifest !== undefined);
        chai.assert.isTrue(res.value.icons !== undefined);
        chai.assert.isTrue(res.value.icons?.color !== undefined);
        chai.assert.isTrue(res.value.icons?.outline !== undefined);
        chai.assert.isTrue(res.value.languages !== undefined);
      }
    });

    it("get package successfully with unsupported file", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(""));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outline.png", new Buffer(""));
      zip.addFile("idk.json", new Buffer(""));
      const archivedFile = zip.toBuffer();
      sandbox.stub(RetryHandler, "Retry").resolves({
        data: archivedFile,
      });
      const loggerSpy = sandbox.stub(logger, "warning").resolves();

      const res = await getAppPackage(teamsAppId, m365TokenProvider, logger);
      chai.assert.isTrue(res.isOk());

      if (res.isOk()) {
        chai.assert.isTrue(loggerSpy.called);
        chai.assert.isUndefined(res.value.languages);
      }
    });

    it("get token error: returns error", async () => {
      m365TokenProvider.getAccessToken = sandbox
        .stub()
        .returns(err(new UserError("token", "token", "", "")));

      const res = await getAppPackage(teamsAppId, m365TokenProvider, logger);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "token");
      }
    });

    it("get package failed due to api", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));

      sandbox.stub(RetryHandler, "Retry").throws();

      const res = await getAppPackage(teamsAppId, m365TokenProvider, logger);
      chai.assert.isTrue(res.isErr());
    });

    it("get package empty response", async () => {
      m365TokenProvider.getAccessToken = sandbox.stub().returns(ok("token"));

      sandbox.stub(RetryHandler, "Retry").resolves({});

      const res = await getAppPackage(teamsAppId, m365TokenProvider, logger);
      chai.assert.isTrue(res.isErr());
    });
  });

  describe("updateTeamsAppV3ForPublish", () => {
    let mockedEnvRestore: RestoreFn | undefined;
    afterEach(() => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("not valid json", async () => {
      const ctx = createContext();
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(""));
      const info = zip.toBuffer();

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "ManifestValidationFailed");
      }
    });

    it("no manifest file", async () => {
      const ctx = createContext();
      const zip = new AdmZip();
      const info = zip.toBuffer();

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };
      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "FileNotFoundError");
      }
    });

    it("manifest without id", async () => {
      const ctx = createContext();
      const json = {
        $schema: "schema",
      };
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(JSON.stringify(json)));
      const info = zip.toBuffer();

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "ManifestValidationFailed");
      }
    });

    it("manifest invalid id", async () => {
      const ctx = createContext();
      const json = {
        id: "fe58d257",
      };
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(JSON.stringify(json)));
      const info = zip.toBuffer();

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        console.log(res.error);
        chai.assert.equal(res.error.name, "ManifestValidationFailed");
      }
    });

    it.skip("manifest no schema", async () => {
      const ctx = createContext();
      const json = {
        id: "fe58d257-4ce6-427e-a388-496c89633774",
      };
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(JSON.stringify(json)));
      const info = zip.toBuffer();

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "ManifestValidationFailed");
      }
    });

    it.skip("manifest validation failed", async () => {
      const ctx = createContext();

      const json = {
        $schema: "schema",
        id: "fe58d257-4ce6-427e-a388-496c89633774",
      };
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(JSON.stringify(json)));
      const info = zip.toBuffer();

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };

      const errors: string[] = ["error1"];
      sandbox.stub(ManifestUtil, "validateManifest").resolves(errors);

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "ManifestValidationFailed");
        chai.assert.isTrue(res.error.message.includes("error1"));
      }
    });

    it("update teams app error", async () => {
      const ctx = createContext();
      const json = {
        $schema: "schema",
        id: "fe58d257-4ce6-427e-a388-496c89633774",
      };
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(JSON.stringify(json)));
      const info = zip.toBuffer();
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };
      const updateDriver = new ConfigureTeamsAppDriver();
      sandbox.stub(Container, "get").callsFake((name) => {
        if (name === "teamsApp/update") {
          return updateDriver;
        } else {
          throw new Error("not implemented");
        }
      });
      sandbox
        .stub(updateDriver, "execute")
        .resolves({ result: err(new UserError("apiError", "apiError", "", "")), summaries: [] });

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "apiError");
      }
    });

    it("happy path", async () => {
      const ctx = createContext();
      const json = {
        $schema: "schema",
        id: "fe58d257-4ce6-427e-a388-496c89633774",
      };
      const zip = new AdmZip();
      zip.addFile("manifest.json", new Buffer(JSON.stringify(json)));
      const info = zip.toBuffer();
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);

      const inputs: InputsWithProjectPath = {
        [QuestionNames.AppPackagePath]: info,
        platform: Platform.VSCode,
        projectPath: "projectPath",
      };
      const updateDriver = new ConfigureTeamsAppDriver();
      sandbox.stub(Container, "get").callsFake((name) => {
        if (name === "teamsApp/update") {
          return updateDriver;
        } else {
          throw new Error("not implemented");
        }
      });
      sandbox.stub(updateDriver, "execute").resolves({ result: ok(new Map([])), summaries: [] });

      const res = await updateTeamsAppV3ForPublish(ctx, inputs);
      chai.assert.isTrue(res.isOk());
    });
  });
});

describe("App-manifest Component - v3", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  const appName = randomAppName();
  const inputs: InputsWithProjectPath = {
    projectPath: getAzureProjectRoot(),
    platform: Platform.VSCode,
    "app-name": appName,
    appPackagePath: "fakePath",
  };
  const cliInputs = {
    projectPath: getAzureProjectRoot(),
    platform: Platform.CLI,
    "app-name": appName,
    appPackagePath: "fakePath",
  };
  const mockDriverRes: ExecutionResult = { result: ok(new Map()), summaries: [] };
  let context: Context;
  setTools(tools);

  beforeEach(() => {
    context = createContext();
    sandbox.stub(tools.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("fakeToken"));
    sandbox.stub(tools.tokenProvider.m365TokenProvider, "getJsonObject").resolves(
      ok({
        unique_name: "fakename",
      })
    );

    context.logProvider = new MockLogProvider();
    context.tokenProvider = {
      m365TokenProvider: new MockedM365Provider(),
      azureAccountProvider: new MockedAzureAccountProvider(),
    };
    sandbox
      .stub(Container, "get")
      .withArgs(sandbox.match("teamsApp/zipAppPackage"))
      .returns(new CreateAppPackageDriver())
      .withArgs(sandbox.match("teamsApp/update"))
      .returns(new ConfigureTeamsAppDriver());
    sandbox.stub(envUtil, "readEnv").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("updateManifestV3 - preview only", async function () {
    const manifest = new TeamsAppManifest();
    manifest.id = "";
    manifest.icons.color = "resources/color.png";
    manifest.icons.outline = "resources/outline.png";
    const updatedManifest = { ...manifest };
    updatedManifest.version = "2.0.0";
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "getManifestV3").resolves(ok(manifest as unknown as TeamsManifest));
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(updatedManifest);
    sandbox.stub(fs, "readFile").resolves(new Buffer(JSON.stringify(manifest)));
    sandbox.stub(context.userInteraction, "showMessage").resolves(ok("Preview only"));
    sandbox.stub(ConfigureTeamsAppDriver.prototype, "execute").resolves(mockDriverRes);
    sandbox.stub(CreateAppPackageDriver.prototype, "execute").resolves(mockDriverRes);

    await updateManifestV3(context, cliInputs);
  });

  it("updateManifestV3 - happy path", async function () {
    const manifest = new TeamsAppManifest();
    manifest.id = "";
    manifest.icons.color = "resources/color.png";
    manifest.icons.outline = "resources/outline.png";
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "getManifestV3").resolves(ok(manifest as unknown as TeamsManifest));
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(manifest);
    sandbox.stub(fs, "readFile").resolves(new Buffer(JSON.stringify(manifest)));
    sandbox.stub(context.userInteraction, "showMessage").resolves(ok("View in Developer Portal"));
    sandbox.stub(ConfigureTeamsAppDriver.prototype, "execute").resolves();

    await updateManifestV3(context, inputs);
  });

  it("updateManifestV3 - rebuild", async function () {
    const manifest = new TeamsAppManifest();
    manifest.id = "";
    manifest.icons.color = "resources/color.png";
    manifest.icons.outline = "resources/outline.png";
    const updatedManifest = { ...manifest };
    updatedManifest.version = "2.0.0";
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "getManifestV3").resolves(ok(manifest as unknown as TeamsManifest));
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "readJSON").resolves(updatedManifest);
    sandbox.stub(fs, "readFile").resolves(new Buffer(JSON.stringify(manifest)));
    sandbox.stub(context.userInteraction, "showMessage").resolves(ok("Preview and update"));
    sandbox.stub(ConfigureTeamsAppDriver.prototype, "execute").resolves(mockDriverRes);
    sandbox.stub(CreateAppPackageDriver.prototype, "execute").resolves(mockDriverRes);

    await updateManifestV3(context, inputs);
  });

  it("updateManifestV3 - getManifestV3 Error", async () => {
    sandbox.stub(manifestUtils, "getTeamsAppManifestPath").resolves("");
    sandbox.stub(manifestUtils, "getManifestV3").resolves(err(new UserError({})));
    const ctx = createContext();
    const inputs: InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: "projectPath",
    };
    const res = await updateManifestV3(ctx, inputs);
    chai.assert.isTrue(res.isErr());
  });
});
