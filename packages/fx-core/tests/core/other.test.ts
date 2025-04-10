// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TokenCredential } from "@azure/identity";
import {
  AzureAccountProvider,
  ok,
  OptionItem,
  Settings,
  SingleSelectQuestion,
  SubscriptionInfo,
  TextInputQuestion,
} from "@microsoft/teamsfx-api";
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
import {
  shareNode,
  selectUsersToRemoveSharedAccess,
  removeSharedAccessNode,
} from "../../src/question/other";
import { CollaborationUtil } from "../../src/core/collaborator";
import { teamsDevPortalClient } from "../../src/client/teamsDevPortalClient";
import * as shareUtils from "../../src/component/driver/share/utils";
import { QuestionNames } from "../../src/question/constants";

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

  describe("shareNode and removeSharedAccessNode", () => {
    beforeEach(() => {
      sandbox.restore();
    });

    it("shareNode should return IQTreeNode with correct children", () => {
      const result = shareNode();

      // Verify the main node structure
      assert.isObject(result);
      assert.property(result, "data");
      assert.property(result, "children");
      assert.isArray(result.children);
      assert.lengthOf(result.children!, 1);

      // Verify share option question node
      const shareOptionNode = result.children![0];
      assert.property(shareOptionNode.data, "name");
      assert.equal(shareOptionNode.data.name, "option");
      assert.property(shareOptionNode.data, "type");
      assert.equal(shareOptionNode.data.type, "singleSelect");

      // Verify children of share option question
      assert.property(shareOptionNode, "children");
      assert.isArray(shareOptionNode.children);
      assert.lengthOf(shareOptionNode.children!, 1);
    });

    it("shareOption question should have correct options", () => {
      const result = shareNode();
      const shareOptionQuestion = result.children![0].data as SingleSelectQuestion;

      assert.isArray(shareOptionQuestion.staticOptions);
      assert.lengthOf(shareOptionQuestion.staticOptions, 2);

      const [shareApp, shareToUser] = shareOptionQuestion.staticOptions as OptionItem[];
      assert.equal(shareApp.id, "share-app");
      assert.equal(shareToUser.id, "share-to-users");
    });

    it("shareToUser question should be shown when shareOption is shareToUser", () => {
      const result = shareNode();
      const shareToUserNode = result.children![0].children![0] as any;

      // Test condition function
      const inputs = { [QuestionNames.ShareOption]: QuestionNames.ShareOptionShareToUser };
      assert.isTrue(shareToUserNode.condition(inputs));

      // Test when not shareToUser
      const inputs2 = { shareOption: "share-app" };
      assert.isFalse(shareToUserNode.condition(inputs2));
    });

    it("shareToUser validation should work correctly", () => {
      const result = shareNode();
      const shareToUserQuestion = result.children![0].children![0].data as any;

      // Test validation function
      const validation = shareToUserQuestion.validation!.validFunc;

      // Empty input should return error message
      assert.isString(validation(""));
      assert.isString(validation("   "));

      // Valid input should return undefined
      assert.isUndefined(validation("user@example.com"));
    });

    it("removeSharedAccessNode should return IQTreeNode with correct structure", () => {
      const result = removeSharedAccessNode();

      assert.isObject(result);
      assert.property(result, "data");
      assert.property(result, "children");
      assert.isArray(result.children);
      assert.lengthOf(result.children!, 1);

      // Check selectUsersToRemoveSharedAccess question
      const selectUsersQuestion = result.children![0].data as any;
      assert.equal(selectUsersQuestion.name, "users");
      assert.equal(selectUsersQuestion.type, "multiSelect");
      assert.property(selectUsersQuestion, "dynamicOptions");
      assert.isTrue(selectUsersQuestion.skipValidation);
    });

    it("selectUsersToRemoveSharedAccess should throw error when project path is undefined", async () => {
      const result = selectUsersToRemoveSharedAccess();

      try {
        await result.dynamicOptions!({} as any);
        assert.fail("Should have thrown error");
      } catch (error: any) {
        assert.equal(error.message, "Project path is not defined");
      }
    });

    it("selectUsersToRemoveSharedAccess should return correct options", async () => {
      // Mock TOOLS and required responses
      const mockTokenProvider = {
        m365TokenProvider: {
          getAccessToken: sandbox.stub().resolves({ isOk: () => true, value: "mock-token" }),
        },
      };
      (global as any).TOOLS = mockTokenProvider;

      // Mock CollaborationUtil.getCurrentUserInfo
      sandbox
        .stub(CollaborationUtil, "getCurrentUserInfo")
        .resolves(ok({ aadId: "current-user-id" } as any));

      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
        .resolves(ok(["mock-teams-app-id"] as any));

      // Mock teamsDevPortalClient.getApp
      sandbox.stub(teamsDevPortalClient, "getApp").resolves({
        userList: [
          {
            aadId: "current-user-id",
            displayName: "Current User",
            userPrincipalName: "current@example.com",
          },
          {
            aadId: "other-user-id",
            displayName: "Other User",
            userPrincipalName: "other@example.com",
          },
        ],
      } as any);

      const result = selectUsersToRemoveSharedAccess();
      const options = await result.dynamicOptions!({ projectPath: "mockPath" } as any);

      assert.isArray(options);
      assert.lengthOf(options, 1); // Should only include the other user, not current user
      assert.equal((options[0] as any).id, "other@example.com");
      assert.equal((options[0] as any).label, "Other User");
    });

    it("selectUsersToRemoveSharedAccess should handle empty user list", async () => {
      // Mock TOOLS and required responses
      const mockTokenProvider = {
        m365TokenProvider: {
          getAccessToken: sandbox.stub().resolves({ isOk: () => true, value: "mock-token" }),
        },
      };
      (global as any).TOOLS = mockTokenProvider;

      // Mock parseShareAppActionYamlConfig
      sandbox.stub(shareUtils, "parseShareAppActionYamlConfig").resolves(ok(["mock-teams-app-id"]));

      // Mock teamsDevPortalClient.getApp with empty user list
      sandbox.stub(teamsDevPortalClient, "getApp").resolves({
        userList: [],
      });

      const result = selectUsersToRemoveSharedAccess();

      try {
        await result.dynamicOptions!({ projectPath: "mockPath" } as any);
        assert.fail("Should have thrown error");
      } catch (error: any) {
        assert.equal(error.message, "No owner found in the app");
      }
    });
  });
});
