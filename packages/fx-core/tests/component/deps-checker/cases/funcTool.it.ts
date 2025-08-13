// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */

import chai from "chai";
import spies from "chai-spies";
import * as fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import semver from "semver";
import * as sinon from "sinon";
import * as uuid from "uuid";
import { FuncToolChecker } from "../../../../src/component/deps-checker/internal/funcToolChecker";
import { isLinux } from "../../../../src/component/deps-checker/util/system";
import * as funcUtils from "../utils/funcTool";

chai.use(spies);
const expect = chai.expect;
const assert = chai.assert;

describe("FuncToolChecker E2E Test", async () => {
  const sandbox = sinon.createSandbox();
  let baseFolder: string | undefined = undefined;
  beforeEach(async function () {
    sandbox.restore();
    baseFolder = path.join(os.homedir(), "func-e2e", uuid.v4().substring(0, 6));
  });

  afterEach(async function () {
    if (baseFolder) {
      await fs.remove(baseFolder);
    }
  });

  const mockFunc = (homeDir: string): FuncToolChecker => {
    sandbox
      .stub(FuncToolChecker, <any>"getDefaultInstallPath")
      .returns(path.join(homeDir, "./.fx/bin/azfunc")) as unknown as FuncToolChecker;
    return new FuncToolChecker();
  };

  it("not install + special character dir", async function () {
    if ((await funcUtils.getGlobalFunc()) || isLinux()) {
      this.skip();
    }

    const projectPath = path.join(baseFolder!, "project path");
    const symlinkPath = path.join(projectPath, "./devTools/func");
    const homePath = path.join(baseFolder!, "Aarón García", "for test");
    await fs.ensureDir(homePath);
    const funcToolChecker = mockFunc(homePath);
    const spyChecker = sandbox.spy(funcToolChecker, "getInstallationInfo");

    const installOptions = {
      projectPath: projectPath,
      symlinkDir: "./devTools/func",
      version: "^4.0.5174",
    };
    const res = await funcToolChecker.resolve(installOptions);
    if (res.error) {
      console.log(res.error);
    }
    assert.isTrue(spyChecker.calledOnce);

    expect(res.isInstalled).to.be.equal(true);
    expect(res.details.binFolders?.length).to.be.equal(1);
    expect(res.details.binFolders?.[0]).to.be.equal(symlinkPath);

    const installationInfo = await funcToolChecker.getInstallationInfo(installOptions);
    expect(installationInfo.isInstalled).to.be.equal(true);
    expect(installationInfo.details.binFolders?.length).to.be.equal(1);
    expect(installationInfo.details.binFolders?.[0]).to.be.equal(symlinkPath);
    expect(res.command).to.be.equal("func");
    await assertFuncStart(symlinkPath);
  });

  it("not install + throw error when installing", async function () {
    if ((await funcUtils.getGlobalFunc()) || isLinux()) {
      this.skip();
    }

    const projectPath = path.join(baseFolder!, "projectDir");
    const symlinkPath = path.join(projectPath, "./devTools/func");
    const homePath = path.join(baseFolder!, "homeDir");
    const funcToolChecker = mockFunc(homePath);
    await fs.ensureFile(path.join(homePath, ".fx/bin/azfunc"));
    const spyChecker = sandbox.spy(funcToolChecker, "getInstallationInfo");

    const installOptions = {
      projectPath: projectPath,
      symlinkDir: "./devTools/func",
      version: "^4.0.5174",
    };
    const res = await funcToolChecker.resolve(installOptions);
    assert.isFalse(res.isInstalled);
    const installationInfo = await funcToolChecker.getInstallationInfo(installOptions);
    assert.isFalse(installationInfo.isInstalled);

    // second: still works well
    await fs.remove(path.join(homePath, ".fx/bin/azfunc"));
    const retryRes = await funcToolChecker.resolve(installOptions);
    if (retryRes.error) {
      console.log(retryRes.error);
    }
    assert.isTrue(retryRes.isInstalled);
    const retryInstallationInfo = await funcToolChecker.getInstallationInfo(installOptions);
    assert.isTrue(retryInstallationInfo.isInstalled, "second run, should success");
    await assertFuncStart(symlinkPath);
  });

  it("not install + linux + user cancel", async function () {
    if ((await funcUtils.getGlobalFunc()) || !isLinux()) {
      this.skip();
    }
    const projectPath = path.join(baseFolder!, "projectDir");
    const homePath = path.join(baseFolder!, "homeDir");
    const funcToolChecker = mockFunc(homePath);

    const installOptions = {
      projectPath: projectPath,
      symlinkDir: "./devTools/func",
      version: "^4.0.5174",
    };
    const depsInfo = await funcToolChecker.resolve(installOptions);

    expect(depsInfo.details.isLinuxSupported).to.be.equal(false);
    expect(depsInfo.command).to.be.equal("func");
    expect(depsInfo.details.binFolders).to.be.equal(undefined);
    expect(depsInfo.error?.message).to.contains(
      "Unable to install https://aka.ms/teamsfx-actions/devtool-install.",
      `Expect error message contains 'Unable to install https://aka.ms/teamsfx-actions/devtool-install.'. Actual error message: ${depsInfo.error?.message}`
    );
  });

  it("already install + linux", async function () {
    const funcVersion = await funcUtils.getGlobalFunc();
    if (!funcVersion || !isLinux()) {
      this.skip();
    }
    if (!semver.satisfies(funcVersion, "^4.0.5174")) {
      this.skip();
    }

    const projectPath = path.join(baseFolder!, "projectDir");
    const homePath = path.join(baseFolder!, "homeDir");
    const funcToolChecker = mockFunc(homePath);

    const installOptions = {
      projectPath: projectPath,
      symlinkDir: "./devTools/func",
      version: "^4.0.5174",
    };
    const depsInfo = await funcToolChecker.resolve(installOptions);
    if (depsInfo.error) {
      console.log(depsInfo.error);
    }
    expect(depsInfo.isInstalled).to.be.equal(true);
    expect(depsInfo.command).to.be.equal("func");
    await assertFuncStart();
  });

  it("already install + old func version", async function () {
    const funcVersion = await funcUtils.getGlobalFunc();
    if (isLinux()) {
      this.skip();
    }
    if (!funcVersion || semver.satisfies(funcVersion, "^4.0.5174")) {
      this.skip();
    }

    const projectPath = path.join(baseFolder!, "projectDir");
    const symlinkPath = path.join(projectPath, "./devTools/func");
    const homePath = path.join(baseFolder!, "homeDir");
    const funcToolChecker = mockFunc(homePath);

    const spyChecker = sandbox.spy(funcToolChecker, "getInstallationInfo");
    const installOptions = {
      projectPath: projectPath,
      symlinkDir: "./devTools/func",
      version: "^4.0.5174",
    };
    const res = await funcToolChecker.resolve(installOptions);
    if (res.error) {
      console.log(res.error);
    }
    assert.isTrue(spyChecker.calledOnce);
    assert.isTrue(res.isInstalled);
    assert.equal(res.details.binFolders?.length, 1);
    assert.equal(res.details.binFolders?.[0], symlinkPath);

    const installationInfo = await funcToolChecker.resolve(installOptions);
    expect(installationInfo.isInstalled).to.be.equal(true);
    assert.equal(installationInfo.command, "func");
    await assertFuncStart(symlinkPath);
  });

  it("already install", async function () {
    const funcVersion = await funcUtils.getGlobalFunc();
    if (isLinux()) {
      this.skip();
    }
    if (!funcVersion || !semver.satisfies(funcVersion, "^4.0.5174")) {
      this.skip();
    }

    const projectPath = path.join(baseFolder!, "projectDir");
    const homePath = path.join(baseFolder!, "homeDir");
    const funcToolChecker = mockFunc(homePath);

    const spyChecker = sandbox.spy(funcToolChecker, "getInstallationInfo");
    const installOptions = {
      projectPath: projectPath,
      symlinkDir: "./devTools/func",
      version: "^4.0.5174",
    };
    const res = await funcToolChecker.resolve(installOptions);
    if (res.error) {
      console.log(res.error);
    }
    assert.isTrue(spyChecker.calledOnce);
    assert.isTrue(res.isInstalled);
    assert.equal(res.details.binFolders, undefined);

    const installationInfo = await funcToolChecker.resolve(installOptions);
    expect(installationInfo.isInstalled).to.be.equal(true);
    assert.equal(installationInfo.command, "func");
    await assertFuncStart();
  });
});

async function assertFuncStart(binFolder?: string): Promise<void> {
  const funcStartResult = await funcUtils.funcStart(binFolder);
  // func start can work: "Unable to find project root. Expecting to find one of host.json, local.settings.json in project root."
  expect(funcStartResult.cmdOutputIncludingStderr).to.includes(
    "Unable to find project root",
    `func start should return error message that contains "Unable to find project root", but actual output: "${funcStartResult.cmdOutputIncludingStderr}"`
  );
}
