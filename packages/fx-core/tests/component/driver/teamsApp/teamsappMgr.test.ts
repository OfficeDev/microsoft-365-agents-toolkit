// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, TeamsAppManifest, err, ok } from "@microsoft/teamsfx-api";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import * as sinon from "sinon";
import { teamsappMgr } from "../../../../src/component/driver/teamsApp/teamsappMgr";
import {
  FileNotFoundError,
  MissingRequiredInputError,
  UserCancelError,
} from "../../../../src/error";
import { envUtil } from "../../../../src/component/utils/envUtil";
import { pathUtils } from "../../../../src/component/utils/pathUtils";
import { CreateAppPackageDriver } from "../../../../src/component/driver/teamsApp/createAppPackage";
import { TOOLS, setTools } from "../../../../src/common/globalVars";
import { MockTools } from "../../../core/utils";
import { ValidateManifestDriver } from "../../../../src/component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../../../../src/component/driver/teamsApp/validateAppPackage";
import { ConfigureTeamsAppDriver } from "../../../../src/component/driver/teamsApp/configure";
import { PublishAppPackageDriver } from "../../../../src/component/driver/teamsApp/publishAppPackage";
import * as CommonTools from "../../../../src/common/tools";

describe("TeamsAppMgr", async () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  describe("ensureAppPackageFile", async () => {
    it("sucess", async () => {
      sandbox
        .stub(teamsappMgr, "packageTeamsApp")
        .resolves(ok({ manifestPath: "", outputJsonPath: "", outputZipPath: "" }));
      sandbox.stub(fs, "pathExists").resolves(true);
      const result = await teamsappMgr.ensureAppPackageFile({
        projectPath: "",
        platform: Platform.CLI,
      });
      chai.assert(result.isOk());
    });
    it("file not found", async () => {
      sandbox
        .stub(teamsappMgr, "packageTeamsApp")
        .resolves(ok({ manifestPath: "", outputJsonPath: "", outputZipPath: "" }));
      sandbox.stub(fs, "pathExists").resolves(false);
      const result = await teamsappMgr.ensureAppPackageFile({
        projectPath: "",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr() && result.error instanceof FileNotFoundError);
    });
    it("packageTeamsApp returns error", async () => {
      sandbox.stub(teamsappMgr, "packageTeamsApp").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.ensureAppPackageFile({
        projectPath: "",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr() && result.error instanceof UserCancelError);
    });
  });

  describe("readManifestFromZip", async () => {
    it("sucess", async () => {
      const result = await teamsappMgr.readManifestFromZip(
        "./tests/component/driver/teamsApp/success.zip"
      );
      chai.assert(result.isOk());
    });
    it("fail", async () => {
      const result = await teamsappMgr.readManifestFromZip(
        "./tests/component/driver/teamsApp/fail.zip"
      );
      chai.assert(result.isErr());
    });
  });

  describe("checkAndTryToLoadEnv", async () => {
    it("no need to resolve", async () => {
      sandbox.stub(fs, "readFile").resolves("abc" as any);
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isOk() && result.value === undefined);
    });

    it("with env-file", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "loadEnvFile").resolves(ok({}));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
        "env-file": "xxx",
      });
      chai.assert(result.isOk() && result.value === undefined);
    });

    it("with env-file but load fail", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "loadEnvFile").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
        "env-file": "xxx",
      });
      chai.assert(result.isErr() && result.error instanceof UserCancelError);
    });

    it("no env-file and list default envs fail", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr() && result.error instanceof UserCancelError);
    });

    it("no env-file and get default env folder fail", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr() && result.error instanceof UserCancelError);
    });

    it("no env-file and get default env folder returns undefined", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok(undefined));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isOk() && result.value === undefined);
    });

    it("has env input, success load target env file", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("abc"));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
        env: "dev",
      });
      chai.assert(result.isOk() && result.value === "dev");
    });

    it("has env input, but not target env file not found", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("abc"));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
        env: "dev2",
      });
      chai.assert(result.isErr() && result.error instanceof FileNotFoundError);
    });

    it("no env input, more than one env available", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "dev2"]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("abc"));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr() && result.error instanceof MissingRequiredInputError);
    });

    it("no env input, only one env available, just use it", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("abc"));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isOk() && result.value === "dev");
    });

    it("no env input, no env file found in default location, do nothing", async () => {
      sandbox.stub(fs, "readFile").resolves("${{APP_NAME}}" as any);
      sandbox.stub(envUtil, "listEnv").resolves(ok([]));
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("abc"));
      const result = await teamsappMgr.checkAndTryToLoadEnv({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isOk() && result.value === undefined);
    });
  });

  describe("packageTeamsApp", async () => {
    const tools = new MockTools();
    setTools(tools);
    it("no manifest file input, default does not exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const result = await teamsappMgr.packageTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });
    it("has manifest file input, but not exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const result = await teamsappMgr.packageTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr());
    });
    it("has manifest file and exists, checkAndTryToLoadEnv fail", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(teamsappMgr, "checkAndTryToLoadEnv").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.packageTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr());
    });
    it("driver fail", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(teamsappMgr, "checkAndTryToLoadEnv").resolves(ok("dev"));
      sandbox.stub(CommonTools, "runForTypeSpecProject").resolves();
      sandbox
        .stub(CreateAppPackageDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.packageTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr());
    });
    it("driver success", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(teamsappMgr, "checkAndTryToLoadEnv").resolves(ok(undefined));
      sandbox.stub(CommonTools, "runForTypeSpecProject").resolves();
      sandbox
        .stub(CreateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      const result = await teamsappMgr.packageTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isOk());
    });
  });

  describe("validateTeamsApp", async () => {
    const tools = new MockTools();
    setTools(tools);
    it("no manifest file and package file input, default does not exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const result = await teamsappMgr.validateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });
    it("input manifest file, load env fail", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(teamsappMgr, "checkAndTryToLoadEnv").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.validateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr());
    });
    it("input manifest file, run driver fail", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(teamsappMgr, "checkAndTryToLoadEnv").resolves(ok(undefined));
      sandbox
        .stub(ValidateManifestDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.validateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isErr());
    });
    it("input manifest file, run driver success", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(teamsappMgr, "checkAndTryToLoadEnv").resolves(ok(undefined));
      sandbox
        .stub(ValidateManifestDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      const result = await teamsappMgr.validateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "manifest-file": "xxx",
      });
      chai.assert(result.isOk());
    });
    it("input package file, run driver success", async () => {
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      const result = await teamsappMgr.validateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "package-file": "xxx",
      });
      chai.assert(result.isOk());
    });
    it("input package file, run driver fail", async () => {
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.validateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
        "package-file": "xxx",
      });
      chai.assert(result.isErr());
    });
  });

  describe("updateTeamsApp", async () => {
    const tools = new MockTools();
    setTools(tools);
    it("ensureAppPackageFile fail", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.updateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });

    it("ValidateAppPackageDriver fail", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.updateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });

    it("ConfigureTeamsAppDriver fail", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      sandbox
        .stub(ConfigureTeamsAppDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.updateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });

    it("readManifestFromZip fail", async () => {
      sandbox
        .stub(TOOLS.tokenProvider.m365TokenProvider, "getJsonObject")
        .resolves(ok({ scope: [] }));
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox.stub(teamsappMgr, "readManifestFromZip").resolves(err(new UserCancelError()));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      sandbox
        .stub(ConfigureTeamsAppDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      const result = await teamsappMgr.updateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });

    it("success", async () => {
      sandbox
        .stub(TOOLS.tokenProvider.m365TokenProvider, "getJsonObject")
        .resolves(ok({ scope: [] }));
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox.stub(teamsappMgr, "readManifestFromZip").resolves(ok(new TeamsAppManifest()));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      sandbox
        .stub(ConfigureTeamsAppDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      const result = await teamsappMgr.updateTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isOk());
    });
  });

  describe("publishTeamsApp", async () => {
    const tools = new MockTools();
    setTools(tools);
    it("ensureAppPackageFile fail", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(err(new UserCancelError()));
      const result = await teamsappMgr.publishTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });

    it("ValidateAppPackageDriver fail", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.publishTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });

    it("PublishAppPackageDriver fail", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      sandbox
        .stub(PublishAppPackageDriver.prototype, "execute")
        .resolves({ result: err(new UserCancelError()), summaries: [] });
      const result = await teamsappMgr.publishTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isErr());
    });
    it("success", async () => {
      sandbox.stub(teamsappMgr, "ensureAppPackageFile").resolves(ok(undefined));
      sandbox.stub(teamsappMgr, "readManifestFromZip").resolves(ok(new TeamsAppManifest()));
      sandbox
        .stub(ValidateAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      sandbox
        .stub(PublishAppPackageDriver.prototype, "execute")
        .resolves({ result: ok(new Map()), summaries: [] });
      const result = await teamsappMgr.publishTeamsApp({
        projectPath: "xxx",
        platform: Platform.CLI,
      });
      chai.assert(result.isOk());
    });
  });
});
