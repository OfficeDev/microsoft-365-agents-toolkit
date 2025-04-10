// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TokenCredential } from "@azure/identity";
import { AzureAccountProvider, Settings, SubscriptionInfo } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv from "mocked-env";
import os from "os";
import * as path from "path";
import sinon from "sinon";
import { isFeatureFlagEnabled } from "../../src/common/featureFlags";
import { execPowerShell, execShell } from "../../src/component/local/process";
import { TaskDefinition } from "../../src/component/local/taskDefinition";
import {
  isValidOfficeAddInProject,
  isValidProject,
  isValidProjectV3,
} from "../../src/common/projectSettingsHelper";
import { cpUtils } from "../../src/component/utils/depsChecker/cpUtils";
import { randomAppName } from "./utils";

describe("Other test case", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("isFeatureFlagEnabled: return true when related environment variable is set to 1 or true", () => {
    const featureFlagName = "FEATURE_FLAG_UNIT_TEST";

    let restore = mockedEnv({
      [featureFlagName]: "1",
    });
    assert.isTrue(isFeatureFlagEnabled(featureFlagName));
    assert.isTrue(isFeatureFlagEnabled(featureFlagName, false)); // default value should be override
    restore();

    restore = mockedEnv({
      [featureFlagName]: "true",
    });
    assert.isTrue(isFeatureFlagEnabled(featureFlagName));
    restore();

    restore = mockedEnv({
      [featureFlagName]: "TruE", // should allow some characters be upper case
    });
    assert.isTrue(isFeatureFlagEnabled(featureFlagName));
    restore();
  });

  it("isFeatureFlagEnabled: return default value when related environment variable is not set", () => {
    const featureFlagName = "FEATURE_FLAG_UNIT_TEST";

    const restore = mockedEnv({
      [featureFlagName]: undefined, // delete it from process.env
    });
    assert.isFalse(isFeatureFlagEnabled(featureFlagName));
    assert.isFalse(isFeatureFlagEnabled(featureFlagName, false));
    assert.isTrue(isFeatureFlagEnabled(featureFlagName, true));
    restore();
  });

  it("isFeatureFlagEnabled: return false when related environment variable is set to non 1 or true value", () => {
    const featureFlagName = "FEATURE_FLAG_UNIT_TEST";

    let restore = mockedEnv({
      [featureFlagName]: "one",
    });
    assert.isFalse(isFeatureFlagEnabled(featureFlagName));
    assert.isFalse(isFeatureFlagEnabled(featureFlagName, true)); // default value should be override
    restore();

    restore = mockedEnv({
      [featureFlagName]: "",
    });
    assert.isFalse(isFeatureFlagEnabled(featureFlagName));
    restore();
  });

  it("executeCommand", async () => {
    {
      try {
        const res = await cpUtils.executeCommand(undefined, undefined, undefined, "ls");
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
    {
      try {
        const res = await cpUtils.tryExecuteCommand(undefined, undefined, undefined, "ls");
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
    {
      try {
        const res = await execShell("ls");
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
    {
      try {
        const res = await execPowerShell("ls");
        assert.isTrue(res !== undefined);
      } catch (e) {}
    }
  });
  it("TaskDefinition", async () => {
    const appName = randomAppName();
    const projectPath = path.resolve(os.tmpdir(), appName);
    {
      const res = TaskDefinition.frontendStart(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendStart(projectPath, "javascript", "echo", true);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendWatch(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.authStart(projectPath, "");
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.botStart(projectPath, "javascript", true);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.ngrokStart(projectPath, true, []);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.frontendInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.backendExtensionsInstall(projectPath, "");
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.botInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.spfxInstall(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.gulpCert(projectPath);
      assert.isTrue(res !== undefined);
    }
    {
      const res = TaskDefinition.gulpServe(projectPath);
      assert.isTrue(res !== undefined);
    }
  });
  it("isValidProject: true", async () => {
    const projectSettings: any = {
      appName: "myapp",
      version: "1.0.0",
      projectId: "123",
    };
    sandbox.stub(fs, "readJsonSync").returns(projectSettings);
    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "readdirSync").returns([]);
    const isValid = isValidProject("aaa");
    assert.isTrue(isValid);
  });
  it("isValidProject v3: true", async () => {
    const mockedEnvRestore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const settings: Settings = {
        version: "1.0.0",
        trackingId: "123",
      };
      sandbox.stub(fs, "readJsonSync").returns(settings);
      sandbox.stub(fs, "existsSync").returns(true);
      sandbox.stub(fs, "readdirSync").returns([]);
      const isValid = isValidProject("aaa");
      assert.isTrue(isValid);
    } finally {
      mockedEnvRestore();
    }
  });
  it("isValidProject v3: false case 1", async () => {
    const mockedEnvRestore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const settings: any = {
        version: "1.0.0",
      };
      sandbox.stub(fs, "readJsonSync").returns(settings);
      const isValid = isValidProject("aaa");
      assert.isFalse(isValid);
    } finally {
      mockedEnvRestore();
    }
  });
  it("isValidProject v3: false case 2", async () => {
    const mockedEnvRestore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const settings: any = {
        projectId: "123",
      };
      sandbox.stub(fs, "readJsonSync").returns(settings);
      const isValid = isValidProject("aaa");
      assert.isFalse(isValid);
    } finally {
      mockedEnvRestore();
    }
  });
  it("projectSettingsHelper - isValidProjectV3 - office add-in", () => {
    sandbox.stub(fs, "readdirSync").returns(["manifest.xml"] as any);
    assert.equal(isValidProjectV3("test"), false);
  });
  it("projectSettingsHelper - isValidOfficeAddInProject - metaos add-in", () => {
    sandbox.stub(fs, "readdirSync").returns(["manifest.json", "manifest.xml"] as any);
    assert.equal(isValidOfficeAddInProject("test"), false);
  });
});
